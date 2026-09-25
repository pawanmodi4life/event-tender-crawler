import os
import re
import io
import json
import time
import hashlib
import logging
import zipfile
import datetime as dt
import urllib.parse

from dataclasses import dataclass, asdict, field
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup
import pandas as pd
import gspread

from google.oauth2 import service_account
from openai import OpenAI

from pydantic import BaseModel, Field
from pypdf import PdfReader
from docx import Document as DocxDocument

try:
    from playwright.sync_api import sync_playwright
except Exception:
    sync_playwright = None


# =============================================================================
# 1. CONFIGURATION
# =============================================================================

APP_NAME = "Pan-India Tender Intelligence Agent"

IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

SERVICE_ACCOUNT_FILE = os.getenv(
    "SERVICE_ACCOUNT_FILE",
    "service_account.json",
)

SPREADSHEET_ID = os.getenv(
    "SPREADSHEET_ID",
    "",
)

EXCEL_MASTER_FILE = os.getenv(
    "EXCEL_MASTER_FILE",
    "Pan_India_Tender_URL_Master.xlsx",
)

OPENAI_API_KEY = os.getenv(
    "OPENAI_API_KEY",
    "",
)

# Current cost-sensitive OpenAI API model.
OPENAI_MODEL = os.getenv(
    "OPENAI_MODEL",
    "gpt-5.6-luna",
)

GEM_LISTING_URL = os.getenv(
    "GEM_LISTING_URL",
    "https://bidplus-global.gem.gov.in/",
)

DOWNLOAD_DIR = Path(
    os.getenv(
        "DOWNLOAD_DIR",
        "tender_downloads",
    )
)

DOWNLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

REQUEST_TIMEOUT = int(
    os.getenv(
        "REQUEST_TIMEOUT",
        "20",
    )
)

CRAWL_WORKERS = int(
    os.getenv(
        "CRAWL_WORKERS",
        "10",
    )
)

MAX_DOC_BYTES = int(
    os.getenv(
        "MAX_DOC_BYTES",
        str(30 * 1024 * 1024),
    )
)

MAX_DOC_TEXT_CHARS = int(
    os.getenv(
        "MAX_DOC_TEXT_CHARS",
        "120000",
    )
)

MAX_GENERIC_LINKS = int(
    os.getenv(
        "MAX_GENERIC_LINKS",
        "35",
    )
)

MAX_DOCUMENTS_PER_TENDER = int(
    os.getenv(
        "MAX_DOCUMENTS_PER_TENDER",
        "20",
    )
)


MAX_PAGES_PER_PORTAL = int(
    os.getenv(
        "MAX_PAGES_PER_PORTAL",
        "100",
    )
)

MAX_GEM_PAGES_PER_SEARCH = int(
    os.getenv(
        "MAX_GEM_PAGES_PER_SEARCH",
        "50",
    )
)

MAX_GENERIC_PAGES = int(
    os.getenv(
        "MAX_GENERIC_PAGES",
        "50",
    )
)


CLEAN_EXISTING_ACTIVE_TENDERS = (
    os.getenv(
        "CLEAN_EXISTING_ACTIVE_TENDERS",
        "true",
    ).strip().lower()
    in {"1", "true", "yes", "y"}
)


# =============================================================================
# 2. LOGGING
# =============================================================================

logging.basicConfig(
    level=os.getenv(
        "LOG_LEVEL",
        "INFO",
    ).upper(),
    format="%(asctime)s | %(levelname)s | %(message)s",
)

log = logging.getLogger(
    "tender-agent",
)


# =============================================================================
# 3. COMPANY PROFILE
# =============================================================================

COMPANY_PROFILE = {
    "agency_name": "Soul Events and Consultancy",
    "constitution": "Proprietorship",
    "avg_3yr_turnover_inr": 38_046_000,
    "fy_2025_26_turnover_inr": 50_411_575,
    "audited_net_worth_inr": 4_680_004,
    "max_single_past_work_order_inr": 7_898_550,
    "max_two_works_threshold_inr": 7_223_000,
    "max_three_works_threshold_inr": 4_628_443,
}


# =============================================================================
# 4. EVENT RELEVANCE
# =============================================================================

EVENT_KEYWORDS = [
    "Event Management",
    "Event Agency",
    "Event Management Agency",
    "Exhibition",
    "Exhibition Stall",
    "Exhibition Pavilion",
    "Stall Design",
    "Stall Fabrication",
    "Expo",
    "Trade Fair",
    "Mela",
    "Festival",
    "Conference",
    "Seminar",
    "Convention",
    "Summit",
    "Conclave",
    "Roadshow",
    "Dealer Meet",
    "Annual Day",
    "Foundation Day",
    "Award Ceremony",
    "Cultural Programme",
    "Corporate Event",
    "Activation",
    "Brand Activation",
    "Experiential Marketing",
    "Publicity Campaign",
    "Outreach Campaign",
    "IEC Campaign",
    "Media Campaign",
    "Advertising Agency",
    "Creative Agency",
    "Communication Agency",
    "PR Agency",
    "AV Production",
    "Audio Visual",
    "Sound and Light",
    "Stage Setup",
    "Stage Production",
    "Tentage",
    "Event Decoration",
    "Venue Management",
    "Event Hospitality",
    "Launch Event",
    "Product Launch",
    "Inauguration",
    "Tourism Event",
    "Sports Event",
    "Government Function",
    "Empanelment of Event Agency",
    "Event Empanelment",
]

STRONG_EVENT_TERMS = [
    item.lower()
    for item in EVENT_KEYWORDS
]

NON_EVENT_EXCLUSIONS = [
    "fabrication and supply",
    "fabrication, load testing",
    "machine fabrication",
    "machining",
    "mechanical work",
    "civil work",
    "civil construction",
    "construction work",
    "pipeline",
    "pipe line",
    "solar panel",
    "solar module",
    "optical",
    "prism",
    "lens",
    "electrical maintenance",
    "electrical repair",
    "customs clearance",
    "freight forwarding",
    "spare parts",
    "vehicle repair",
    "vehicle maintenance",
    "material supply",
    "supply of material",
    "equipment supply",
    "equipment maintenance",
    "scientific equipment",
    "laboratory equipment",
    "network equipment",
    "housekeeping",
    "security services",
    "manpower supply",
    "outsourcing manpower",
    "canteen",
    "catering services only",
    "medical equipment",
    "pharmaceutical",
]

TENDERISH_TERMS = (
    "tender",
    "bid",
    "rfp",
    "eoi",
    "procurement",
    "quotation",
    "corrigendum",
    "nit",
    "empanelment",
    "notice inviting",
)

DOWNLOADABLE_EXTENSIONS = (
    ".pdf",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".zip",
)

BAD_TENDER_TITLES = {
    "tender",
    "tender title",
    "tender/e-auction",
    "tender / e-auction",
    "e-auction",
    "e auction",
    "procurement",
    "procurement details",
    "bid",
    "bid title",
    "latest tenders",
    "active tenders",
    "tenders",
    "title",
    "tenders published",
    "tender document",
}


# =============================================================================
# 5. DATA MODELS
# =============================================================================

@dataclass
class Portal:
    portal: str
    category: str
    state: str
    url: str
    priority: str = "P2"
    source_type: str = ""
    active: str = "Yes"


@dataclass
class CrawlHealth:
    portal: str
    url: str
    status: str
    discovered_count: int = 0
    message: str = ""
    checked_at: str = field(
        default_factory=lambda: dt.datetime.now(
            tz=IST
        ).strftime(
            "%Y-%m-%d %H:%M:%S IST"
        )
    )


@dataclass
class TenderCandidate:
    portal: str
    state: str
    source_url: str
    tender_id: str = ""
    organization: str = ""
    title: str = ""
    category: str = ""
    deadline_raw: str = "NOT VERIFIED"
    estimated_value_raw: str = "NOT VERIFIED"
    detail_url: str = ""
    discovered_doc_urls: list[str] = field(
        default_factory=list
    )
    keyword_hits: list[str] = field(
        default_factory=list
    )

    def stable_key(self):
        raw = "|".join([
            (self.tender_id or "").strip().lower(),
            (self.organization or "").strip().lower(),
            (self.title or "").strip().lower(),
            (self.deadline_raw or "").strip().lower(),
            (
                self.detail_url
                or self.source_url
                or ""
            ).strip().lower(),
        ])

        return hashlib.sha256(
            raw.encode("utf-8")
        ).hexdigest()[:24]


@dataclass
class DownloadedDocument:
    url: str
    local_path: str
    mime_type: str
    text: str
    sha256: str
    error: str = ""


@dataclass
class QualificationDecision:
    status: str
    reason: str
    action_plan: str


# =============================================================================
# 6. STRUCTURED OUTPUT MODEL
# =============================================================================

class EligibilityEvidence(BaseModel):
    tender_reference: str | None = None
    organisation: str | None = None
    scope_summary: str | None = None

    estimated_value_inr: float | None = None

    emd_inr: float | None = None
    emd_exemption_explicit: bool | None = None
    emd_exemption_text: str | None = None

    average_turnover_required_inr: float | None = None
    turnover_years: int | None = None
    msme_turnover_relaxation_explicit: bool | None = None

    single_similar_work_required_inr: float | None = None
    two_similar_works_each_required_inr: float | None = None
    three_similar_works_each_required_inr: float | None = None
    experience_requirement_text: str | None = None

    min_net_worth_required_inr: float | None = None

    entity_type_restriction: str | None = None
    consortium_or_jv_allowed: bool | None = None

    physical_submission_required: bool | None = None
    qcbs_or_technical_pitch: bool | None = None
    named_celebrity_or_artist_mandate: bool | None = None

    submission_deadline: str | None = None
    pre_bid_date: str | None = None

    non_core_scope: bool | None = None
    non_core_scope_reason: str | None = None

    evidence_quotes: list[str] = Field(
        default_factory=list
    )

    missing_or_unclear_fields: list[str] = Field(
        default_factory=list
    )


# =============================================================================
# 7. HELPERS
# =============================================================================

def normalize_space(value):
    return re.sub(
        r"\s+",
        " ",
        value or "",
    ).strip()


def keyword_hits(text):
    text_lower = (
        text or ""
    ).lower()

    return [
        keyword
        for keyword in EVENT_KEYWORDS
        if keyword.lower() in text_lower
    ]


def infer_category(title):
    text = normalize_space(
        title or ""
    ).lower()

    if any(
        item in text
        for item in [
            "exhibition stall",
            "exhibition pavilion",
            "stall design",
            "stall fabrication",
            "pavilion design",
            "pavilion fabrication",
            "expo",
            "trade fair",
        ]
    ):
        return "Exhibition / Stall"

    if "empanel" in text and any(
        item in text
        for item in [
            "event",
            "exhibition",
            "advertising",
            "creative",
            "publicity",
            "media",
            "communication",
        ]
    ):
        return "Empanelment"

    if any(
        item in text
        for item in [
            "conference management",
            "organising conference",
            "organizing conference",
            "conduct of conference",
            "summit management",
            "organising summit",
            "organizing summit",
            "conclave management",
            "organising conclave",
            "organizing conclave",
        ]
    ):
        return "Conference / Conclave"

    if any(
        item in text
        for item in [
            "advertising agency",
            "creative agency",
            "media campaign",
            "publicity campaign",
            "outreach campaign",
            "iec campaign",
            "communication agency",
            "pr agency",
        ]
    ):
        return "Advertising / Creative / Outreach"

    if any(
        item in text
        for item in [
            "festival management",
            "mela management",
            "organising mela",
            "organizing mela",
            "cultural event",
            "cultural programme",
        ]
    ):
        return "Festival / Mela / Cultural"

    if any(
        item in text
        for item in [
            "event management",
            "event agency",
            "event production",
            "event execution",
            "corporate event",
            "annual day",
            "foundation day",
            "award ceremony",
            "brand activation",
            "experiential marketing",
            "roadshow",
            "dealer meet",
            "launch event",
            "product launch",
        ]
    ):
        return "Event / Experiential"

    return "Unclassified"


def safe_urljoin(base, href):
    return urllib.parse.urljoin(
        base,
        href or "",
    )


def is_downloadable_url(url):
    try:
        path = urllib.parse.urlsplit(
            url
        ).path.lower()

        return path.endswith(
            DOWNLOADABLE_EXTENSIONS
        )

    except Exception:
        return False


def valid_tender_title(title):
    title = normalize_space(
        title
    )

    if not title:
        return False

    lower = title.lower()

    if lower in BAD_TENDER_TITLES:
        return False

    if len(title) < 8:
        return False

    bad_patterns = [
        r"^click here$",
        r"^view details$",
        r"^download$",
        r"^more$",
        r"^read more$",
        r"^details$",
    ]

    for pattern in bad_patterns:
        if re.match(
            pattern,
            lower,
        ):
            return False

    return True


FALSE_POSITIVE_PHRASES = [
    "pre-tender conference",
    "pre tender conference",
    "pre-bid conference",
    "pre bid conference",
    "pre-bid meeting",
    "pre bid meeting",
    "pre nit conference",
    "pre-nit conference",
    "public procurement seminar",
    "procurement seminar",
    "tender conference",
    "tender creation",
    "tender terms",
    "procurement plan",
    "government initiatives",
    "news updates",
    "latest news",
    "search eoi / pre nit conference",
    "search eoi/pre nit conference",
]

GENERIC_NAV_TITLES = {
    "dae secretariat matters",
    "dae unit tenders",
    "dae procurement plan",
    "list of internal dae tenders",
    "public sector units",
    "government initiatives",
    "news updates",
    "downloads",
    "e-tender",
    "e tender",
    "tender creation",
    "tender terms",
    "procurement plan",
    "latest tenders",
    "active tenders",
    "tenders",
    "tender document",
    "tender notices",
    "notice",
    "notices",
    "tender portal",
    "procurement",
    "tender",
}

EVENT_SERVICE_PHRASES = [
    "event management",
    "event management agency",
    "event agency",
    "event organiser",
    "event organizer",
    "event production",
    "event execution",
    "event logistics",
    "event coordination",
    "event planning",
    "event services",
    "event partner",

    "exhibition stall",
    "exhibition pavilion",
    "stall design",
    "stall fabrication",
    "stall construction",
    "pavilion design",
    "pavilion fabrication",
    "pavilion construction",
    "exhibition design",
    "exhibition management",
    "expo management",
    "trade fair management",

    "conference management",
    "organising conference",
    "organizing conference",
    "organisation of conference",
    "organization of conference",
    "conduct of conference",

    "summit management",
    "organising summit",
    "organizing summit",
    "organisation of summit",
    "organization of summit",

    "conclave management",
    "organising conclave",
    "organizing conclave",
    "organisation of conclave",
    "organization of conclave",

    "seminar management",
    "organising seminar",
    "organizing seminar",

    "mela management",
    "organising mela",
    "organizing mela",
    "festival management",
    "organising festival",
    "organizing festival",
    "cultural event",
    "cultural programme",

    "brand activation",
    "experiential marketing",
    "roadshow",
    "dealer meet",
    "annual day",
    "foundation day",
    "award ceremony",
    "launch event",
    "product launch",
    "inauguration event",

    "publicity campaign",
    "outreach campaign",
    "iec campaign",
    "media campaign",
    "advertising agency",
    "creative agency",
    "communication agency",
    "pr agency",

    "audio visual",
    "av production",
    "stage setup",
    "stage production",
    "sound and light",
    "tentage",
    "event decoration",
]

EVENT_OBJECT_TERMS = [
    "event",
    "exhibition",
    "expo",
    "trade fair",
    "conference",
    "conclave",
    "summit",
    "seminar",
    "mela",
    "festival",
    "roadshow",
    "annual day",
    "foundation day",
    "award ceremony",
    "launch",
    "pavilion",
    "stall",
]

SERVICE_ACTION_TERMS = [
    "organise",
    "organize",
    "organising",
    "organizing",
    "organisation",
    "organization",
    "conduct",
    "manage",
    "management",
    "execute",
    "execution",
    "design",
    "fabricate",
    "fabrication",
    "setup",
    "set up",
    "production",
    "agency",
    "empanelment",
    "empanel",
    "selection of agency",
    "appointment of agency",
]


def is_relevant_event_title(title):
    """
    Decide whether a title represents an actual event/exhibition/creative
    service opportunity, not merely a procurement page or a procurement
    process meeting.
    """

    title = normalize_space(
        title or ""
    ).lower()

    if not title:
        return False

    if title in GENERIC_NAV_TITLES:
        return False

    if any(
        phrase in title
        for phrase in FALSE_POSITIVE_PHRASES
    ):
        return False

    if any(
        exclusion in title
        for exclusion in NON_EVENT_EXCLUSIONS
    ):
        return False

    # Explicit high-confidence service phrases.
    if any(
        phrase in title
        for phrase in EVENT_SERVICE_PHRASES
    ):
        return True

    # Empanelment must explicitly concern our service area.
    if (
        "empanel" in title
        and any(
            term in title
            for term in [
                "event",
                "exhibition",
                "advertising",
                "creative",
                "media",
                "publicity",
                "communication",
                "experiential",
            ]
        )
    ):
        return True

    # General service-action + event-object combination.
    has_event_object = any(
        term in title
        for term in EVENT_OBJECT_TERMS
    )

    has_service_action = any(
        term in title
        for term in SERVICE_ACTION_TERMS
    )

    if has_event_object and has_service_action:
        return True

    return False


def is_relevant_event_tender(candidate):
    # Relevance must be based on the actual tender title.
    # Do NOT use inferred category or full webpage text here.
    return is_relevant_event_title(
        candidate.title
    )


def extract_ref_candidates(text):
    patterns = [
        r"GEM/\d{4}/B/\d+",
        r"\b\d{4}_[A-Za-z0-9_-]+_\d+_\d+\b",
        (
            r"\b"
            r"(?:RFP|EOI|NIT|TENDER|BID)"
            r"[\s:/-]*"
            r"[A-Za-z0-9._/-]{3,}"
            r"\b"
        ),
    ]

    output = []

    for pattern in patterns:
        output.extend(
            re.findall(
                pattern,
                text or "",
                flags=re.I,
            )
        )

    return list(
        dict.fromkeys(
            normalize_space(item)
            for item in output
        )
    )


def session_with_headers():
    session = requests.Session()

    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/128.0 Safari/537.36"
        ),
        "Accept-Language": "en-IN,en;q=0.9",
    })

    return session


# =============================================================================
# 8. ENVIRONMENT VALIDATION
# =============================================================================

def validate_environment():
    problems = []

    if not SPREADSHEET_ID:
        problems.append(
            "SPREADSHEET_ID is missing."
        )

    if not os.path.exists(
        SERVICE_ACCOUNT_FILE
    ):
        problems.append(
            f"Service account file missing: "
            f"{SERVICE_ACCOUNT_FILE}"
        )

    if not os.path.exists(
        EXCEL_MASTER_FILE
    ):
        problems.append(
            f"Tender master Excel missing: "
            f"{EXCEL_MASTER_FILE}"
        )

    # OPENAI_API_KEY is optional.
    # The crawler can still run and write tenders to the sheet without AI.
    if not OPENAI_API_KEY:
        log.warning(
            "OPENAI_API_KEY is missing. "
            "Crawler will continue, but AI eligibility extraction "
            "will be marked unavailable."
        )

    if problems:
        for item in problems:
            log.error(item)

        raise RuntimeError(
            "Environment validation failed."
        )


# =============================================================================
# 9. LOAD MASTER PORTALS
# =============================================================================

def load_master_portals():
    df = pd.read_excel(
        EXCEL_MASTER_FILE,
        sheet_name="Tender URL Master",
        skiprows=2,
    )

    required_columns = {
        "Portal / Organisation",
        "Category",
        "State/Region",
        "URL",
        "Priority",
        "Source Type",
        "Active",
    }

    missing_columns = (
        required_columns
        - set(df.columns)
    )

    if missing_columns:
        raise RuntimeError(
            "Master Excel is missing columns: "
            + ", ".join(
                sorted(missing_columns)
            )
        )

    portals = []

    for _, row in df.iterrows():
        active = str(
            row.get(
                "Active",
                "",
            )
        ).strip().lower()

        if active != "yes":
            continue

        url = str(
            row.get(
                "URL",
                "",
            )
        ).strip()

        if not url.startswith(
            (
                "http://",
                "https://",
            )
        ):
            continue

        portals.append(
            Portal(
                portal=str(
                    row.get(
                        "Portal / Organisation",
                        "",
                    )
                ).strip(),

                category=str(
                    row.get(
                        "Category",
                        "",
                    )
                ).strip(),

                state=str(
                    row.get(
                        "State/Region",
                        "",
                    )
                ).strip(),

                url=url,

                priority=str(
                    row.get(
                        "Priority",
                        "P2",
                    )
                ).strip().upper()
                or "P2",

                source_type=str(
                    row.get(
                        "Source Type",
                        "",
                    )
                ).strip(),
            )
        )

    return portals


# =============================================================================
# 10. CRAWLER
# =============================================================================

class TenderCrawler:
    def __init__(self):
        self.http = session_with_headers()

    def crawl(self, portal):
        try:
            if "gem.gov.in" in portal.url.lower():
                return (
                    [],
                    CrawlHealth(
                        portal.portal,
                        portal.url,
                        "DEDICATED_ADAPTER",
                    ),
                )

            if self._looks_gepnic(
                portal
            ):
                items = self.crawl_gepnic(
                    portal
                )
            else:
                items = self.crawl_generic(
                    portal
                )

            return (
                items,
                CrawlHealth(
                    portal.portal,
                    portal.url,
                    "SUCCESS",
                    len(items),
                ),
            )

        except requests.exceptions.SSLError as error:
            return (
                [],
                CrawlHealth(
                    portal.portal,
                    portal.url,
                    "SSL_ERROR",
                    0,
                    str(error)[:250],
                ),
            )

        except requests.exceptions.Timeout:
            return (
                [],
                CrawlHealth(
                    portal.portal,
                    portal.url,
                    "TIMEOUT",
                    0,
                    "HTTP request timed out",
                ),
            )

        except requests.exceptions.ConnectionError as error:
            message = str(error)

            if (
                "NameResolutionError" in message
                or
                "Failed to resolve" in message
                or
                "Temporary failure in name resolution" in message
            ):
                status = "DNS_ERROR"
            else:
                status = "CONNECTION_ERROR"

            return (
                [],
                CrawlHealth(
                    portal.portal,
                    portal.url,
                    status,
                    0,
                    message[:250],
                ),
            )

        except requests.exceptions.HTTPError as error:
            code = getattr(
                error.response,
                "status_code",
                None,
            )

            status = (
                "LOGIN_OR_BLOCKED"
                if code in (
                    401,
                    403,
                )
                else
                "HTTP_ERROR"
            )

            return (
                [],
                CrawlHealth(
                    portal.portal,
                    portal.url,
                    status,
                    0,
                    str(error)[:250],
                ),
            )

        except RuntimeError as error:
            message = str(error)

            if (
                "CAPTCHA" in message.upper()
                or
                "LOGIN" in message.upper()
            ):
                status = "CAPTCHA_OR_LOGIN"
            else:
                status = "PARSER_FAILED"

            return (
                [],
                CrawlHealth(
                    portal.portal,
                    portal.url,
                    status,
                    0,
                    message[:250],
                ),
            )

        except Exception as error:
            log.exception(
                "Crawler failed: %s",
                portal.portal,
            )

            return (
                [],
                CrawlHealth(
                    portal.portal,
                    portal.url,
                    "PARSER_FAILED",
                    0,
                    str(error)[:250],
                ),
            )

    def _looks_gepnic(
        self,
        portal,
    ):
        url = portal.url.lower()

        return any(
            item in url
            for item in [
                "eprocure.gov.in",
                "etenders.gov.in",
                "tenders.gov.in",
                "tenders.nic.in",
                "etender",
                "eproc",
                ".nic.in",
            ]
        )

    def crawl_gepnic(
        self,
        portal,
    ):
        """
        Crawl GePNIC/CPPP-style listings with pagination.

        Strategy:
        - load latest active tenders page
        - parse current page
        - discover/follow Next/page links
        - stop on repeat page, disabled Next, no new records, or max pages
        """

        base = portal.url.rstrip("/")

        candidate_urls = []

        if "/app" in base.lower():
            separator = (
                "&"
                if "?" in base
                else "?"
            )

            candidate_urls.append(
                base
                + separator
                + "page=FrontEndLatestActiveTenders"
                + "&service=page"
            )

        candidate_urls.extend([
            (
                f"{base}/nicgep/app"
                "?page=FrontEndLatestActiveTenders"
                "&service=page"
            ),
            (
                f"{base}/eprocure/app"
                "?page=FrontEndLatestActiveTenders"
                "&service=page"
            ),
            (
                f"{base}/app"
                "?page=FrontEndLatestActiveTenders"
                "&service=page"
            ),
            base,
        ])

        first_html = ""
        first_url = ""

        for url in dict.fromkeys(
            candidate_urls
        ):
            try:
                response = self.http.get(
                    url,
                    timeout=REQUEST_TIMEOUT,
                    allow_redirects=True,
                )

                if (
                    response.status_code == 200
                    and
                    len(response.text) > 500
                ):
                    first_html = response.text
                    first_url = response.url
                    break

            except requests.RequestException:
                continue

        if not first_html:
            raise RuntimeError(
                "No usable GePNIC page returned."
            )

        if (
            "captcha" in first_html.lower()
            and
            len(first_html) < 100000
        ):
            raise RuntimeError(
                "CAPTCHA_OR_LOGIN_REQUIRED"
            )

        def parse_page(
            html,
            page_url,
        ):
            soup = BeautifulSoup(
                html,
                "html.parser",
            )

            page_results = []

            for row in soup.find_all(
                "tr"
            ):
                row_text = normalize_space(
                    row.get_text(
                        " ",
                        strip=True,
                    )
                )

                hits = keyword_hits(
                    row_text
                )

                if not hits:
                    continue

                cells = [
                    normalize_space(
                        td.get_text(
                            " ",
                            strip=True,
                        )
                    )
                    for td in row.find_all(
                        "td"
                    )
                ]

                if len(cells) < 2:
                    continue

                anchor_titles = [
                    normalize_space(
                        a.get_text(
                            " ",
                            strip=True,
                        )
                    )
                    for a in row.find_all(
                        "a"
                    )
                    if valid_tender_title(
                        normalize_space(
                            a.get_text(
                                " ",
                                strip=True,
                            )
                        )
                    )
                ]

                if anchor_titles:
                    title = max(
                        anchor_titles,
                        key=len,
                    )
                else:
                    usable_cells = [
                        cell
                        for cell in cells
                        if valid_tender_title(
                            cell
                        )
                    ]

                    if not usable_cells:
                        continue

                    title = max(
                        usable_cells,
                        key=len,
                    )

                if not valid_tender_title(
                    title
                ):
                    continue

                refs = extract_ref_candidates(
                    row_text
                )

                tender_id = (
                    refs[0]
                    if refs
                    else ""
                )

                detail_url = page_url
                documents = []

                for anchor in row.find_all(
                    "a",
                    href=True,
                ):
                    href = safe_urljoin(
                        page_url,
                        anchor.get(
                            "href"
                        ),
                    )

                    if not href:
                        continue

                    if href.lower().startswith(
                        "javascript:"
                    ):
                        continue

                    anchor_text = normalize_space(
                        anchor.get_text(
                            " ",
                            strip=True,
                        )
                    )

                    if is_downloadable_url(
                        href
                    ):
                        documents.append(
                            href
                        )

                    elif any(
                        item
                        in (
                            anchor_text
                            + " "
                            + href
                        ).lower()
                        for item in TENDERISH_TERMS
                    ):
                        detail_url = href

                page_results.append(
                    TenderCandidate(
                        portal=portal.portal,
                        state=portal.state,
                        source_url=page_url,
                        tender_id=tender_id,
                        organization=portal.portal,
                        title=title[:500],
                        category=infer_category(
                            title
                        ),
                        detail_url=detail_url,
                        discovered_doc_urls=list(
                            dict.fromkeys(
                                documents
                            )
                        ),
                        keyword_hits=hits,
                    )
                )

            # Pagination candidates: explicit Next plus numbered page links.
            next_urls = []

            for anchor in soup.find_all(
                "a",
                href=True,
            ):
                label = normalize_space(
                    anchor.get_text(
                        " ",
                        strip=True,
                    )
                ).lower()

                href = safe_urljoin(
                    page_url,
                    anchor.get(
                        "href"
                    ),
                )

                if not href:
                    continue

                if href.lower().startswith(
                    "javascript:"
                ):
                    continue

                if (
                    label in {
                        "next",
                        "next >",
                        "next »",
                        ">",
                        "»",
                    }
                    or
                    "next" in label
                    or
                    re.fullmatch(
                        r"\d+",
                        label or "",
                    )
                ):
                    next_urls.append(
                        href
                    )

            return (
                page_results,
                list(
                    dict.fromkeys(
                        next_urls
                    )
                ),
            )

        results = []
        visited_pages = set()
        queued_pages = [first_url]
        html_cache = {
            first_url: first_html
        }

        page_no = 0
        no_new_streak = 0
        seen_tender_keys = set()

        while (
            queued_pages
            and
            page_no < MAX_PAGES_PER_PORTAL
        ):
            current_url = queued_pages.pop(0)

            if current_url in visited_pages:
                continue

            visited_pages.add(
                current_url
            )

            try:
                if current_url in html_cache:
                    html = html_cache.pop(
                        current_url
                    )
                else:
                    response = self.http.get(
                        current_url,
                        timeout=REQUEST_TIMEOUT,
                        allow_redirects=True,
                    )

                    if response.status_code != 200:
                        continue

                    html = response.text

            except requests.RequestException:
                continue

            page_no += 1

            page_results, next_urls = parse_page(
                html,
                current_url,
            )

            new_count = 0

            for tender in page_results:
                key = tender.stable_key()

                if key in seen_tender_keys:
                    continue

                seen_tender_keys.add(
                    key
                )
                results.append(
                    tender
                )
                new_count += 1

            log.info(
                "%s | GePNIC page %d | %d new matching tenders | total %d",
                portal.portal,
                page_no,
                new_count,
                len(results),
            )

            if new_count == 0:
                no_new_streak += 1
            else:
                no_new_streak = 0

            # Stop if multiple pages in a row add nothing.
            if no_new_streak >= 3:
                break

            for next_url in next_urls:
                if (
                    next_url not in visited_pages
                    and
                    next_url not in queued_pages
                ):
                    queued_pages.append(
                        next_url
                    )

        log.info(
            "%s | GePNIC pagination complete | pages=%d | tenders=%d",
            portal.portal,
            page_no,
            len(results),
        )

        return self._dedupe(
            results
        )

    def crawl_generic(
        self,
        portal,
    ):
        """
        Crawl generic tender websites with best-effort pagination.

        We follow:
        - explicit Next links
        - common numbered page links
        - tender/procurement listing links

        We still avoid unlimited crawling.
        """

        first_response = self.http.get(
            portal.url,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
        )

        first_response.raise_for_status()

        lower_page = first_response.text.lower()

        if (
            "captcha" in lower_page
            and
            len(first_response.text) < 50000
        ):
            raise RuntimeError(
                "CAPTCHA_OR_LOGIN_REQUIRED"
            )

        queue = [
            first_response.url
        ]

        html_cache = {
            first_response.url:
                first_response.text
        }

        visited = set()
        results = []
        seen_tender_keys = set()

        page_no = 0
        no_new_streak = 0

        while (
            queue
            and
            page_no < MAX_GENERIC_PAGES
        ):
            page_url = queue.pop(0)

            if page_url in visited:
                continue

            visited.add(
                page_url
            )

            try:
                if page_url in html_cache:
                    html = html_cache.pop(
                        page_url
                    )
                else:
                    response = self.http.get(
                        page_url,
                        timeout=REQUEST_TIMEOUT,
                        allow_redirects=True,
                    )

                    if response.status_code != 200:
                        continue

                    html = response.text

            except requests.RequestException:
                continue

            page_no += 1

            soup = BeautifulSoup(
                html,
                "html.parser",
            )

            links = []
            pagination_links = []

            for anchor in soup.find_all(
                "a",
                href=True,
            ):
                label = normalize_space(
                    anchor.get_text(
                        " ",
                        strip=True,
                    )
                )

                href = safe_urljoin(
                    page_url,
                    anchor.get(
                        "href"
                    ),
                )

                if not href:
                    continue

                combined = (
                    label
                    + " "
                    + href
                ).lower()

                if any(
                    term in combined
                    for term in TENDERISH_TERMS
                ):
                    links.append(
                        (
                            label,
                            href,
                        )
                    )

                label_lower = label.lower()

                if (
                    label_lower in {
                        "next",
                        "next >",
                        "next »",
                        ">",
                        "»",
                    }
                    or
                    "next" in label_lower
                    or
                    re.fullmatch(
                        r"\d+",
                        label_lower or "",
                    )
                ):
                    pagination_links.append(
                        href
                    )

            links = list(
                dict.fromkeys(
                    links
                )
            )

            new_count = 0

            for (
                label,
                href,
            ) in links[:MAX_GENERIC_LINKS]:

                if is_downloadable_url(
                    href
                ):
                    hits = keyword_hits(
                        label
                    )

                    if (
                        hits
                        and
                        valid_tender_title(
                            label
                        )
                    ):
                        tender = TenderCandidate(
                            portal=portal.portal,
                            state=portal.state,
                            source_url=page_url,
                            organization=portal.portal,
                            title=label[:500],
                            category=infer_category(
                                label
                            ),
                            detail_url=href,
                            discovered_doc_urls=[
                                href
                            ],
                            keyword_hits=hits,
                        )

                        key = tender.stable_key()

                        if key not in seen_tender_keys:
                            seen_tender_keys.add(
                                key
                            )
                            results.append(
                                tender
                            )
                            new_count += 1

                    continue

                try:
                    child_response = self.http.get(
                        href,
                        timeout=REQUEST_TIMEOUT,
                        allow_redirects=True,
                    )

                    if child_response.status_code != 200:
                        continue

                    child = BeautifulSoup(
                        child_response.text,
                        "html.parser",
                    )

                    # Important: relevance comes from actual link/title,
                    # not footer/menu text from the entire page.
                    hits = keyword_hits(
                        label
                    )

                    if not hits:
                        continue

                    title = label

                    if not valid_tender_title(
                        title
                    ):
                        heading = child.find(
                            [
                                "h1",
                                "h2",
                                "h3",
                            ]
                        )

                        if heading:
                            title = normalize_space(
                                heading.get_text(
                                    " ",
                                    strip=True,
                                )
                            )

                    if not valid_tender_title(
                        title
                    ):
                        continue

                    documents = []

                    for child_anchor in child.find_all(
                        "a",
                        href=True,
                    ):
                        child_url = safe_urljoin(
                            child_response.url,
                            child_anchor.get(
                                "href"
                            ),
                        )

                        if is_downloadable_url(
                            child_url
                        ):
                            documents.append(
                                child_url
                            )

                    child_text = normalize_space(
                        child.get_text(
                            " ",
                            strip=True,
                        )
                    )

                    refs = extract_ref_candidates(
                        child_text
                    )

                    tender = TenderCandidate(
                        portal=portal.portal,
                        state=portal.state,
                        source_url=page_url,
                        tender_id=(
                            refs[0]
                            if refs
                            else ""
                        ),
                        organization=portal.portal,
                        title=title[:500],
                        category=infer_category(
                            title
                        ),
                        detail_url=child_response.url,
                        discovered_doc_urls=list(
                            dict.fromkeys(
                                documents
                            )
                        )[:MAX_DOCUMENTS_PER_TENDER],
                        keyword_hits=hits,
                    )

                    key = tender.stable_key()

                    if key not in seen_tender_keys:
                        seen_tender_keys.add(
                            key
                        )
                        results.append(
                            tender
                        )
                        new_count += 1

                except requests.RequestException:
                    continue

            log.info(
                "%s | generic page %d | %d new matching tenders | total %d",
                portal.portal,
                page_no,
                new_count,
                len(results),
            )

            if new_count == 0:
                no_new_streak += 1
            else:
                no_new_streak = 0

            if no_new_streak >= 3:
                break

            for next_url in list(
                dict.fromkeys(
                    pagination_links
                )
            ):
                if (
                    next_url not in visited
                    and
                    next_url not in queue
                ):
                    queue.append(
                        next_url
                    )

        log.info(
            "%s | generic pagination complete | pages=%d | tenders=%d",
            portal.portal,
            page_no,
            len(results),
        )

        return self._dedupe(
            results
        )

    def _dedupe(
        self,
        items,
    ):
        seen = set()
        output = []

        for tender in items:
            if not valid_tender_title(
                tender.title
            ):
                continue

            key = tender.stable_key()

            if key in seen:
                continue

            seen.add(key)
            output.append(tender)

        return output


# =============================================================================
# 11. GEM CRAWLER
# =============================================================================

def crawl_gem():
    """
    Crawl the live GeM GTE listing surface with pagination.

    For each event-related search term:
    - search
    - parse visible bids
    - capture real hrefs
    - click/follow Next until exhausted or max pages reached
    """

    if sync_playwright is None:
        return (
            [],
            CrawlHealth(
                "Government e-Marketplace (GeM)",
                GEM_LISTING_URL,
                "DEPENDENCY_MISSING",
                0,
                "Playwright is not installed.",
            ),
        )

    search_terms = [
        "Event Management",
        "Exhibition",
        "Conference",
        "Conclave",
        "Summit",
        "Mela",
        "Festival",
        "Creative Agency",
        "Advertising Agency",
        "Brand Activation",
        "Audio Visual",
        "Sound and Light",
        "Empanelment Event",
        "Outreach Campaign",
        "Experiential Marketing",
        "Foundation Day",
    ]

    results = []

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                headless=True,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ],
            )

            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 "
                    "(Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 "
                    "(KHTML, like Gecko) "
                    "Chrome/128.0 Safari/537.36"
                )
            )

            page = context.new_page()

            page.goto(
                GEM_LISTING_URL,
                timeout=45000,
                wait_until="domcontentloaded",
            )

            page.wait_for_timeout(
                3000
            )

            search_ui_found = False

            for term in search_terms:
                search_input = None

                selectors = [
                    'input[placeholder*="Keyword" i]',
                    'input[placeholder*="Search" i]',
                    'input[type="search"]',
                    'input#search_by',
                ]

                for selector in selectors:
                    locator = page.locator(
                        selector
                    ).first

                    if (
                        locator.count()
                        and
                        locator.is_visible()
                    ):
                        search_input = locator
                        search_ui_found = True
                        break

                if search_input is None:
                    continue

                try:
                    search_input.fill("")
                    search_input.fill(
                        term
                    )
                    page.keyboard.press(
                        "Enter"
                    )

                    page.wait_for_timeout(
                        2500
                    )

                    seen_page_signatures = set()
                    seen_bid_ids = set()

                    for page_number in range(
                        1,
                        MAX_GEM_PAGES_PER_SEARCH + 1,
                    ):
                        body_text = page.locator(
                            "body"
                        ).inner_text()

                        bid_numbers = list(
                            dict.fromkeys(
                                re.findall(
                                    r"GEM/\d{4}/B/\d+",
                                    body_text,
                                )
                            )
                        )

                        signature = hashlib.sha256(
                            (
                                "|".join(
                                    bid_numbers
                                )
                                + "|"
                                + page.url
                            ).encode(
                                "utf-8"
                            )
                        ).hexdigest()

                        if signature in seen_page_signatures:
                            log.info(
                                "GeM | %s | repeated page detected; stopping.",
                                term,
                            )
                            break

                        seen_page_signatures.add(
                            signature
                        )

                        anchors = page.locator(
                            "a[href]"
                        )

                        hrefs = []

                        for index in range(
                            min(
                                anchors.count(),
                                1500,
                            )
                        ):
                            try:
                                anchor = anchors.nth(
                                    index
                                )

                                href = anchor.get_attribute(
                                    "href"
                                )

                                if not href:
                                    continue

                                label = normalize_space(
                                    anchor.inner_text()
                                )

                                absolute = safe_urljoin(
                                    page.url,
                                    href,
                                )

                                combined = (
                                    f"{label} {absolute}"
                                ).lower()

                                if (
                                    "showbiddocument" in combined
                                    or
                                    "bidplus" in combined
                                    or
                                    "gem/" in combined
                                ):
                                    hrefs.append(
                                        (
                                            label,
                                            absolute,
                                        )
                                    )

                            except Exception:
                                continue

                        new_count = 0

                        for bid_no in bid_numbers:
                            if bid_no in seen_bid_ids:
                                continue

                            seen_bid_ids.add(
                                bid_no
                            )

                            numeric_bid = bid_no.split(
                                "/"
                            )[-1]

                            matching_urls = []

                            for label, href in hrefs:
                                combined = (
                                    f"{label} {href}"
                                ).lower()

                                if (
                                    bid_no.lower()
                                    in combined
                                    or
                                    numeric_bid
                                    in combined
                                ):
                                    matching_urls.append(
                                        href
                                    )

                            real_url = (
                                matching_urls[0]
                                if matching_urls
                                else page.url
                            )

                            results.append(
                                TenderCandidate(
                                    portal=(
                                        "Government "
                                        "e-Marketplace (GeM)"
                                    ),
                                    state="Pan India",
                                    source_url=GEM_LISTING_URL,
                                    tender_id=bid_no,
                                    organization="NOT VERIFIED",
                                    title=(
                                        f"{term} | "
                                        f"{bid_no}"
                                    ),
                                    category=infer_category(
                                        term
                                    ),
                                    detail_url=real_url,
                                    discovered_doc_urls=(
                                        matching_urls[:5]
                                        if matching_urls
                                        else []
                                    ),
                                    keyword_hits=[
                                        term
                                    ],
                                )
                            )

                            new_count += 1

                        log.info(
                            "GeM | %s | page %d | %d new bids | total for search %d",
                            term,
                            page_number,
                            new_count,
                            len(
                                seen_bid_ids
                            ),
                        )

                        # Try common Next selectors.
                        next_locator = None

                        next_selectors = [
                            'a:has-text("Next")',
                            'button:has-text("Next")',
                            'li.next a',
                            'a[aria-label*="Next" i]',
                            'button[aria-label*="Next" i]',
                            'a[rel="next"]',
                        ]

                        for selector in next_selectors:
                            locator = page.locator(
                                selector
                            ).first

                            try:
                                if (
                                    locator.count()
                                    and
                                    locator.is_visible()
                                ):
                                    disabled = (
                                        locator.get_attribute(
                                            "disabled"
                                        )
                                        is not None
                                    )

                                    aria_disabled = (
                                        (
                                            locator.get_attribute(
                                                "aria-disabled"
                                            )
                                            or ""
                                        ).lower()
                                        ==
                                        "true"
                                    )

                                    classes = (
                                        locator.get_attribute(
                                            "class"
                                        )
                                        or ""
                                    ).lower()

                                    if (
                                        disabled
                                        or aria_disabled
                                        or "disabled" in classes
                                    ):
                                        continue

                                    next_locator = locator
                                    break

                            except Exception:
                                continue

                        if next_locator is None:
                            break

                        before_url = page.url
                        before_text = body_text[:3000]

                        try:
                            next_locator.click(
                                timeout=10000
                            )

                            page.wait_for_timeout(
                                2000
                            )

                            after_text = page.locator(
                                "body"
                            ).inner_text()[:3000]

                            if (
                                page.url == before_url
                                and
                                after_text == before_text
                            ):
                                break

                        except Exception:
                            break

                except Exception as error:
                    log.warning(
                        "GeM search failed for [%s]: %s",
                        term,
                        str(error)[:200],
                    )

            browser.close()

            if not search_ui_found:
                raise RuntimeError(
                    "GeM search UI was not found. "
                    "Portal layout may have changed."
                )

        unique = {}

        for tender in results:
            key = (
                tender.tender_id
                or
                tender.stable_key()
            )
            unique[key] = tender

        output = list(
            unique.values()
        )

        return (
            output,
            CrawlHealth(
                "Government e-Marketplace (GeM)",
                GEM_LISTING_URL,
                (
                    "SUCCESS"
                    if output
                    else "NO_RESULTS"
                ),
                len(output),
                (
                    ""
                    if output
                    else (
                        "GeM page/search UI loaded, but no matching "
                        "event/exhibition GTE bids were discovered."
                    )
                ),
            ),
        )

    except Exception as error:
        return (
            [],
            CrawlHealth(
                "Government e-Marketplace (GeM)",
                GEM_LISTING_URL,
                "PARSER_FAILED",
                0,
                str(error)[:250],
            ),
        )


# =============================================================================
# 12. DOCUMENT MANAGER
# =============================================================================

class DocumentManager:
    def __init__(self):
        self.http = session_with_headers()

    def discover_from_detail_page(
        self,
        candidate,
    ):
        urls = list(
            candidate.discovered_doc_urls
        )

        if not candidate.detail_url:
            return list(
                dict.fromkeys(
                    urls
                )
            )

        if is_downloadable_url(
            candidate.detail_url
        ):
            urls.append(
                candidate.detail_url
            )

            return list(
                dict.fromkeys(
                    urls
                )
            )

        try:
            response = self.http.get(
                candidate.detail_url,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )

            if response.status_code != 200:
                return list(
                    dict.fromkeys(
                        urls
                    )
                )

            content_type = (
                response.headers.get(
                    "content-type"
                )
                or ""
            ).lower()

            if "application/pdf" in content_type:
                urls.append(
                    response.url
                )

                return list(
                    dict.fromkeys(
                        urls
                    )
                )

            soup = BeautifulSoup(
                response.text,
                "html.parser",
            )

            relevant_labels = [
                "download",
                "tender document",
                "rfp",
                "atc",
                "corrigendum",
                "nit",
                "boq",
                "annexure",
                "eligibility",
                "pre-bid",
                "pre bid",
            ]

            for anchor in soup.find_all(
                "a",
                href=True,
            ):
                url = safe_urljoin(
                    response.url,
                    anchor.get(
                        "href"
                    ),
                )

                label = normalize_space(
                    anchor.get_text(
                        " ",
                        strip=True,
                    )
                ).lower()

                if (
                    is_downloadable_url(
                        url
                    )
                    or
                    any(
                        item in label
                        for item in relevant_labels
                    )
                ):
                    urls.append(url)

        except requests.RequestException:
            pass

        return list(
            dict.fromkeys(
                urls
            )
        )

    def download_and_extract(
        self,
        url,
        tender_key,
    ):
        try:
            response = self.http.get(
                url,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
                stream=True,
            )

            response.raise_for_status()

            content_type = (
                response.headers.get(
                    "content-type"
                )
                or ""
            ).lower()

            data = bytearray()

            for chunk in response.iter_content(
                65536
            ):
                if not chunk:
                    continue

                data.extend(chunk)

                if len(data) > MAX_DOC_BYTES:
                    raise ValueError(
                        "Document exceeded configured size limit."
                    )

            raw = bytes(data)

            sha256 = hashlib.sha256(
                raw
            ).hexdigest()

            extension = self._guess_ext(
                response.url,
                content_type,
                raw,
            )

            folder = (
                DOWNLOAD_DIR
                / tender_key
            )

            folder.mkdir(
                parents=True,
                exist_ok=True,
            )

            filepath = (
                folder
                / (
                    f"{sha256[:12]}"
                    f"{extension}"
                )
            )

            filepath.write_bytes(
                raw
            )

            text = self._extract_text(
                raw,
                extension,
                content_type,
            )

            return DownloadedDocument(
                url=response.url,
                local_path=str(
                    filepath
                ),
                mime_type=content_type,
                text=text[:MAX_DOC_TEXT_CHARS],
                sha256=sha256,
            )

        except Exception as error:
            return DownloadedDocument(
                url=url,
                local_path="",
                mime_type="",
                text="",
                sha256="",
                error=str(error)[:400],
            )

    def _guess_ext(
        self,
        url,
        content_type,
        raw,
    ):
        extension = Path(
            urllib.parse.urlsplit(
                url
            ).path
        ).suffix.lower()

        if extension in DOWNLOADABLE_EXTENSIONS:
            return extension

        if (
            raw.startswith(
                b"%PDF"
            )
            or
            "application/pdf" in content_type
        ):
            return ".pdf"

        if raw.startswith(
            b"PK\x03\x04"
        ):
            if (
                "spreadsheetml"
                in content_type
            ):
                return ".xlsx"

            if (
                "wordprocessingml"
                in content_type
            ):
                return ".docx"

            return ".zip"

        return ".bin"

    def _extract_text(
        self,
        raw,
        extension,
        content_type,
    ):
        if extension == ".pdf":
            return self._pdf_text(
                raw
            )

        if extension == ".docx":
            return self._docx_text(
                raw
            )

        if extension == ".xlsx":
            return self._excel_text(
                raw
            )

        if extension == ".zip":
            return self._zip_text(
                raw
            )

        # Legacy .doc/.xls are intentionally not silently parsed.
        # They will result in NO_DOCUMENT_TEXT and require manual review.
        if extension in (
            ".doc",
            ".xls",
        ):
            return ""

        if "text/" in content_type:
            return raw.decode(
                "utf-8",
                errors="ignore",
            )

        return ""

    def _pdf_text(
        self,
        raw,
    ):
        try:
            reader = PdfReader(
                io.BytesIO(
                    raw
                )
            )

            pages = []

            for page_number, page in enumerate(
                reader.pages
            ):
                text = (
                    page.extract_text()
                    or ""
                )

                if text:
                    pages.append(
                        f"\n--- PDF PAGE "
                        f"{page_number + 1} ---\n"
                        f"{text}"
                    )

            return "\n".join(
                pages
            )

        except Exception:
            return ""

    def _docx_text(
        self,
        raw,
    ):
        try:
            document = DocxDocument(
                io.BytesIO(
                    raw
                )
            )

            parts = []

            for paragraph in document.paragraphs:
                if paragraph.text.strip():
                    parts.append(
                        paragraph.text
                    )

            for table in document.tables:
                for row in table.rows:
                    parts.append(
                        " | ".join(
                            cell.text
                            for cell in row.cells
                        )
                    )

            return "\n".join(
                parts
            )

        except Exception:
            return ""

    def _excel_text(
        self,
        raw,
    ):
        try:
            excel = pd.ExcelFile(
                io.BytesIO(
                    raw
                )
            )

            parts = []

            for sheet_name in excel.sheet_names[:20]:
                dataframe = pd.read_excel(
                    io.BytesIO(
                        raw
                    ),
                    sheet_name=sheet_name,
                    header=None,
                )

                parts.append(
                    f"\n--- SHEET: "
                    f"{sheet_name} ---\n"
                )

                parts.append(
                    dataframe.astype(
                        str
                    ).to_csv(
                        index=False,
                        header=False,
                    )
                )

            return "".join(
                parts
            )

        except Exception:
            return ""

    def _zip_text(
        self,
        raw,
    ):
        results = []

        try:
            with zipfile.ZipFile(
                io.BytesIO(
                    raw
                )
            ) as archive:

                for entry in archive.infolist()[:75]:
                    extension = Path(
                        entry.filename
                    ).suffix.lower()

                    if extension not in (
                        ".pdf",
                        ".docx",
                        ".xlsx",
                        ".txt",
                    ):
                        continue

                    if entry.file_size > MAX_DOC_BYTES:
                        continue

                    data = archive.read(
                        entry
                    )

                    if extension == ".pdf":
                        text = self._pdf_text(
                            data
                        )

                    elif extension == ".docx":
                        text = self._docx_text(
                            data
                        )

                    elif extension == ".xlsx":
                        text = self._excel_text(
                            data
                        )

                    else:
                        text = data.decode(
                            "utf-8",
                            errors="ignore",
                        )

                    if text:
                        results.append(
                            f"\n--- ZIP FILE: "
                            f"{entry.filename} ---\n"
                            f"{text}"
                        )

        except Exception:
            pass

        return "\n".join(
            results
        )


# =============================================================================
# 13. OPENAI EXTRACTION
# =============================================================================

_OPENAI_CLIENT = None
_OPENAI_DISABLED_FOR_RUN = False
_OPENAI_DISABLE_REASON = ""


def get_openai_client():
    global _OPENAI_CLIENT

    if _OPENAI_CLIENT is None:
        _OPENAI_CLIENT = OpenAI(
            api_key=OPENAI_API_KEY,
            max_retries=0,
            timeout=60.0,
        )

    return _OPENAI_CLIENT


def _ai_unavailable_result(
    status,
    message,
):
    return {
        "extraction_status": status,
        "missing_or_unclear_fields": [
            message
        ],
        "evidence_quotes": [],
        "ai_model_used": "",
    }


def extract_eligibility(
    candidate,
    documents,
):
    global _OPENAI_DISABLED_FOR_RUN
    global _OPENAI_DISABLE_REASON

    usable_documents = [
        document
        for document in documents
        if (
            document.text
            and
            document.text.strip()
        )
    ]

    if not usable_documents:
        return _ai_unavailable_result(
            "NO_DOCUMENT_TEXT",
            "Tender document text could not be extracted.",
        )

    if not OPENAI_API_KEY:
        return _ai_unavailable_result(
            "OPENAI_API_KEY_MISSING",
            "OPENAI_API_KEY is not configured.",
        )

    if _OPENAI_DISABLED_FOR_RUN:
        return _ai_unavailable_result(
            "OPENAI_DISABLED_FOR_RUN",
            (
                _OPENAI_DISABLE_REASON
                or
                "OpenAI API was disabled for this workflow run."
            ),
        )

    combined_text = "\n\n".join(
        (
            f"SOURCE URL: {document.url}\n"
            f"LOCAL FILE: {document.local_path}\n"
            f"{document.text}"
        )
        for document in usable_documents
    )[:MAX_DOC_TEXT_CHARS]

    prompt = (
        "You are an Indian government procurement tender "
        "eligibility extraction engine.\n\n"

        "Your task is ONLY to extract factual eligibility "
        "requirements from the supplied tender documents.\n\n"

        "STRICT RULES:\n"
        "1. Extract ONLY information explicitly supported "
        "by the tender documents.\n"
        "2. NEVER assume standard GFR, GeM, MSME, EMD, "
        "turnover, work experience, consortium, exemption, "
        "80/50/40 or any standard tender rule.\n"
        "3. If something is not stated or unclear, use null.\n"
        "4. Do not convert an unstated requirement into zero.\n"
        "5. Convert explicitly stated monetary values into "
        "numeric INR amounts.\n"
        "6. If the tender asks for two or three similar works, "
        "return the amount required PER WORK.\n"
        "7. entity_type_restriction must explain whether "
        "proprietorship, partnership, LLP, private limited "
        "or public limited entities are permitted/restricted.\n"
        "8. experience_requirement_text should preserve the "
        "experience eligibility clause concisely.\n"
        "9. evidence_quotes should contain short exact snippets "
        "supporting important eligibility criteria.\n"
        "10. Add uncertain or missing critical criteria to "
        "missing_or_unclear_fields.\n\n"

        "DISCOVERED TENDER METADATA:\n"
        + json.dumps(
            asdict(candidate),
            ensure_ascii=False,
            indent=2,
        )
        + "\n\nTENDER DOCUMENT CONTENT:\n"
        + combined_text
    )

    try:
        client = get_openai_client()

        log.info(
            "Extracting eligibility using OpenAI model: %s",
            OPENAI_MODEL,
        )

        response = client.responses.parse(
            model=OPENAI_MODEL,
            input=[
                {
                    "role": "system",
                    "content": (
                        "Extract tender eligibility facts strictly "
                        "from supplied documents. "
                        "Do not infer missing requirements."
                    ),
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            text_format=EligibilityEvidence,
        )

        parsed = response.output_parsed

        if parsed is None:
            return {
                "extraction_status": "OPENAI_EXTRACTION_FAILED",
                "missing_or_unclear_fields": [
                    "OpenAI returned no structured output."
                ],
                "evidence_quotes": [],
                "ai_model_used": OPENAI_MODEL,
            }

        data = parsed.model_dump()

        data[
            "extraction_status"
        ] = "DOCUMENT_VERIFIED_EXTRACTION"

        data[
            "ai_model_used"
        ] = OPENAI_MODEL

        return data

    except Exception as error:
        message = str(error)
        lower_message = message.lower()

        if (
            "credit_balance_exhausted" in lower_message
            or
            "insufficient_quota" in lower_message
            or
            "no credits remaining" in lower_message
            or
            "billing" in lower_message
        ):
            _OPENAI_DISABLED_FOR_RUN = True
            _OPENAI_DISABLE_REASON = (
                "OpenAI API credits/quota unavailable for this run."
            )

            log.error(
                "OpenAI API quota/billing unavailable. "
                "Disabling AI extraction for remaining tenders."
            )

            return {
                "extraction_status": "OPENAI_QUOTA_EXHAUSTED",
                "missing_or_unclear_fields": [
                    _OPENAI_DISABLE_REASON
                ],
                "evidence_quotes": [],
                "ai_model_used": OPENAI_MODEL,
            }

        if (
            "invalid_api_key" in lower_message
            or
            "authentication" in lower_message
            or
            "incorrect api key" in lower_message
            or
            "401" in lower_message
        ):
            _OPENAI_DISABLED_FOR_RUN = True
            _OPENAI_DISABLE_REASON = (
                "OpenAI API authentication failed for this run."
            )

            log.error(
                "OpenAI authentication failed. "
                "Disabling AI extraction for remaining tenders."
            )

            return {
                "extraction_status": "OPENAI_AUTH_FAILED",
                "missing_or_unclear_fields": [
                    _OPENAI_DISABLE_REASON
                ],
                "evidence_quotes": [],
                "ai_model_used": OPENAI_MODEL,
            }

        if (
            "model_not_found" in lower_message
            or
            "does not exist" in lower_message
            or
            "not have access to model" in lower_message
            or
            "404" in lower_message
        ):
            _OPENAI_DISABLED_FOR_RUN = True
            _OPENAI_DISABLE_REASON = (
                f"Configured OpenAI model is unavailable: "
                f"{OPENAI_MODEL}"
            )

            log.error(
                "%s",
                _OPENAI_DISABLE_REASON,
            )

            return {
                "extraction_status": "OPENAI_MODEL_UNAVAILABLE",
                "missing_or_unclear_fields": [
                    _OPENAI_DISABLE_REASON
                ],
                "evidence_quotes": [],
                "ai_model_used": OPENAI_MODEL,
            }

        log.exception(
            "OpenAI eligibility extraction failed"
        )

        return {
            "extraction_status": "OPENAI_EXTRACTION_FAILED",
            "missing_or_unclear_fields": [
                message[:500]
            ],
            "evidence_quotes": [],
            "ai_model_used": OPENAI_MODEL,
        }


# =============================================================================
# 14. QUALIFICATION ENGINE
# =============================================================================

def evaluate_qualification(
    evidence,
):
    extraction_status = evidence.get(
        "extraction_status"
    )

    if (
        extraction_status
        !=
        "DOCUMENT_VERIFIED_EXTRACTION"
    ):
        return QualificationDecision(
            status="DOCUMENT REVIEW REQUIRED",
            reason=(
                "Verified structured eligibility evidence "
                "is not available."
            ),
            action_plan=(
                "Review the source tender documents manually."
            ),
        )

    fail_reasons = []

    entity = (
        evidence.get(
            "entity_type_restriction"
        )
        or ""
    ).lower()

    turnover = evidence.get(
        "average_turnover_required_inr"
    )

    single_work = evidence.get(
        "single_similar_work_required_inr"
    )

    two_work = evidence.get(
        "two_similar_works_each_required_inr"
    )

    three_work = evidence.get(
        "three_similar_works_each_required_inr"
    )

    net_worth = evidence.get(
        "min_net_worth_required_inr"
    )

    jv_allowed = evidence.get(
        "consortium_or_jv_allowed"
    )

    if entity:
        company_only_terms = [
            "private limited only",
            "public limited only",
            "companies only",
            "company only",
        ]

        if (
            any(
                term in entity
                for term in company_only_terms
            )
            and
            "propriet" not in entity
        ):
            fail_reasons.append(
                "Tender entity restriction does not "
                "permit a proprietorship."
            )

    if (
        turnover is not None
        and
        turnover
        >
        COMPANY_PROFILE[
            "avg_3yr_turnover_inr"
        ]
        and
        evidence.get(
            "msme_turnover_relaxation_explicit"
        )
        is not True
    ):
        fail_reasons.append(
            "Required average turnover "
            f"₹{turnover:,.0f} exceeds audited benchmark "
            f"₹{COMPANY_PROFILE['avg_3yr_turnover_inr']:,.0f}."
        )

    if (
        single_work is not None
        and
        single_work
        >
        COMPANY_PROFILE[
            "max_single_past_work_order_inr"
        ]
        and
        jv_allowed is not True
    ):
        fail_reasons.append(
            "Required single similar work "
            f"₹{single_work:,.0f} exceeds benchmark "
            f"₹{COMPANY_PROFILE['max_single_past_work_order_inr']:,.0f}."
        )

    if (
        two_work is not None
        and
        two_work
        >
        COMPANY_PROFILE[
            "max_two_works_threshold_inr"
        ]
        and
        jv_allowed is not True
    ):
        fail_reasons.append(
            "Required value for each of two similar works "
            f"₹{two_work:,.0f} exceeds benchmark "
            f"₹{COMPANY_PROFILE['max_two_works_threshold_inr']:,.0f}."
        )

    if (
        three_work is not None
        and
        three_work
        >
        COMPANY_PROFILE[
            "max_three_works_threshold_inr"
        ]
        and
        jv_allowed is not True
    ):
        fail_reasons.append(
            "Required value for each of three similar works "
            f"₹{three_work:,.0f} exceeds benchmark "
            f"₹{COMPANY_PROFILE['max_three_works_threshold_inr']:,.0f}."
        )

    if (
        net_worth is not None
        and
        net_worth
        >
        COMPANY_PROFILE[
            "audited_net_worth_inr"
        ]
    ):
        fail_reasons.append(
            "Required net worth "
            f"₹{net_worth:,.0f} exceeds audited benchmark "
            f"₹{COMPANY_PROFILE['audited_net_worth_inr']:,.0f}."
        )

    if evidence.get(
        "non_core_scope"
    ) is True:
        fail_reasons.append(
            "The principal scope is outside event / "
            "exhibition / experiential services."
        )

    if fail_reasons:
        return QualificationDecision(
            status="VERIFIED DISQUALIFIED",
            reason=" ".join(
                fail_reasons
            ),
            action_plan=(
                "Do not submit unless an amendment, relaxation "
                "or permitted JV structure resolves the blocker."
            ),
        )

    # Critical completeness check BEFORE manual-intervention classification.
    critical_missing = []

    if turnover is None:
        critical_missing.append(
            "Average turnover requirement not conclusively verified"
        )

    if (
        single_work is None
        and
        two_work is None
        and
        three_work is None
        and
        not evidence.get(
            "experience_requirement_text"
        )
    ):
        critical_missing.append(
            "Past/similar-work experience requirement "
            "not conclusively verified"
        )

    if evidence.get(
        "entity_type_restriction"
    ) is None:
        critical_missing.append(
            "Entity constitution eligibility "
            "not conclusively verified"
        )

    if evidence.get(
        "submission_deadline"
    ) is None:
        critical_missing.append(
            "Submission deadline not verified"
        )

    unclear_items = (
        evidence.get(
            "missing_or_unclear_fields"
        )
        or []
    )

    critical_terms = (
        "turnover",
        "experience",
        "similar work",
        "entity",
        "constitution",
        "eligibility",
        "qualification",
        "net worth",
    )

    for item in unclear_items:
        item_text = str(
            item
        )

        if any(
            term in item_text.lower()
            for term in critical_terms
        ):
            critical_missing.append(
                item_text
            )

    critical_missing = list(
        dict.fromkeys(
            critical_missing
        )
    )

    if critical_missing:
        return QualificationDecision(
            status="DOCUMENT REVIEW REQUIRED",
            reason=(
                "Critical eligibility information could not "
                "be fully verified: "
                + "; ".join(
                    critical_missing[:8]
                )
            ),
            action_plan=(
                "Review Eligibility / PQC / ATC / Experience "
                "clauses manually before treating this opportunity "
                "as qualified."
            ),
        )

    manual_reasons = []

    if evidence.get(
        "physical_submission_required"
    ) is True:
        manual_reasons.append(
            "Physical submission / DD / BG or offline "
            "document requirement"
        )

    if evidence.get(
        "qcbs_or_technical_pitch"
    ) is True:
        manual_reasons.append(
            "QCBS / technical presentation or creative pitch required"
        )

    if evidence.get(
        "named_celebrity_or_artist_mandate"
    ) is True:
        manual_reasons.append(
            "Named artist / celebrity mandate or authorization required"
        )

    if manual_reasons:
        return QualificationDecision(
            status="NEEDS MANUAL INTERVENTION",
            reason="; ".join(
                manual_reasons
            ),
            action_plan=(
                "Review commercial feasibility, submission logistics "
                "and supporting documentation."
            ),
        )

    return QualificationDecision(
        status="VERIFIED QUALIFIED",
        reason=(
            "No disqualifying condition was found after "
            "document-based eligibility extraction against "
            "the configured company benchmarks."
        ),
        action_plan=(
            "Proceed with technical/commercial bid preparation "
            "and verify the latest corrigenda before submission."
        ),
    )


# =============================================================================
# 15. GOOGLE SHEET HEADERS
# =============================================================================

ACTIVE_HEADERS = [
    "Date Found",
    "Tender Key",
    "Tender ID / Ref No",
    "Portal Name",
    "State",
    "Organization / Dept",
    "Tender Title & Scope",
    "Category",
    "Estimated Value (INR)",
    "EMD (INR)",
    "EMD/MSME Exemption Evidence",
    "Submission Deadline",
    "Min Turnover Req",
    "Past Experience Requirement",
    "Entity Restriction",
    "JV/Consortium Allowed",
    "Key Compliance",
    "RFP / Tender Doc URLs",
    "Downloaded Files",
    "Pre-Bid Date",
    "Portal / Detail Link",
    "Qualification Status",
    "Remarks & Action Plan",
    "Evidence Quotes",
    "Extraction Status",
    "AI Model Used",
]

PORTAL_HEADERS = [
    "Portal",
    "Category",
    "State",
    "URL",
    "Active",
    "Last Crawl",
    "Tender Count",
    "Crawler Status",
    "Crawler Message",
    "Crawl Priority",
]


# =============================================================================
# 16. GOOGLE SHEETS
# =============================================================================

def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = (
        service_account
        .Credentials
        .from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=scopes,
        )
    )

    return gspread.authorize(
        credentials
    )


def ensure_worksheet(
    spreadsheet,
    name,
    rows=5000,
    cols=35,
):
    try:
        return spreadsheet.worksheet(
            name
        )

    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(
            title=name,
            rows=rows,
            cols=cols,
        )


def existing_tender_keys(
    sheet,
):
    values = sheet.get_all_values()

    if not values:
        return set()

    headers = values[0]

    try:
        key_column = headers.index(
            "Tender Key"
        )
    except ValueError:
        return set()

    return {
        row[key_column]
        for row in values[1:]
        if (
            len(row) > key_column
            and
            row[key_column]
        )
    }


def clean_active_tenders_sheet(sheet):
    """
    Remove stale false-positive rows created by older crawler logic.

    This does NOT delete valid event tenders merely because AI extraction
    failed. It only removes rows whose tender title fails the same strict
    relevance test used for new candidates.
    """

    if not CLEAN_EXISTING_ACTIVE_TENDERS:
        return 0

    values = sheet.get_all_values()

    if len(values) <= 1:
        return 0

    headers = values[0]

    try:
        title_index = headers.index(
            "Tender Title & Scope"
        )
    except ValueError:
        log.warning(
            "Cannot clean Active_Tenders: title column not found."
        )
        return 0

    kept_rows = []
    removed = 0

    for row in values[1:]:
        padded = list(row) + [
            ""
        ] * max(
            0,
            len(headers) - len(row),
        )

        padded = padded[:len(headers)]

        title = padded[
            title_index
        ]

        if not normalize_space(
            title
        ):
            continue

        if is_relevant_event_title(
            title
        ):
            kept_rows.append(
                padded
            )
        else:
            removed += 1
            log.info(
                "Removing stale false-positive tracker row: %s",
                title[:160],
            )

    if removed:
        sheet.clear()

        output = [
            headers
        ] + kept_rows

        sheet.update(
            values=output,
            range_name=(
                f"A1:Z{len(output)}"
            ),
        )

        log.info(
            "Cleaned Active_Tenders: removed %d false-positive rows, "
            "kept %d rows.",
            removed,
            len(kept_rows),
        )

    return removed


def sync_portal_directory(
    sheet,
    portals,
    health,
):
    rows = [
        PORTAL_HEADERS
    ]

    for portal in portals:
        # Prefer exact portal-name health.
        status = health.get(
            portal.portal
        )

        # GeM master naming can differ; map by URL as fallback.
        if status is None:
            for candidate_health in health.values():
                if (
                    candidate_health.url
                    and
                    candidate_health.url.rstrip("/")
                    ==
                    portal.url.rstrip("/")
                ):
                    status = candidate_health
                    break

        rows.append([
            portal.portal,
            portal.category,
            portal.state,
            portal.url,
            portal.active,
            (
                status.checked_at
                if status
                else ""
            ),
            (
                status.discovered_count
                if status
                else 0
            ),
            (
                status.status
                if status
                else "NOT_RUN"
            ),
            (
                status.message
                if status
                else ""
            ),
            portal.priority,
        ])

    sheet.clear()

    sheet.update(
        values=rows,
        range_name=f"A1:J{len(rows)}",
    )


# =============================================================================
# 17. CANDIDATE FILTER
# =============================================================================

def filter_event_candidates(
    candidates,
):
    filtered = []
    rejected = []

    for candidate in candidates:
        if not valid_tender_title(
            candidate.title
        ):
            continue

        if not is_relevant_event_tender(
            candidate
        ):
            rejected.append(
                candidate
            )

            log.info(
                "Skipping non-event tender: %s",
                candidate.title[:150],
            )
            continue

        filtered.append(
            candidate
        )

    log.info(
        "Event relevance filter: %d kept / %d rejected",
        len(filtered),
        len(rejected),
    )

    return filtered


# =============================================================================
# 18. PROCESS TENDER
# =============================================================================

def process_candidate(
    candidate,
    document_manager,
):
    document_urls = (
        document_manager
        .discover_from_detail_page(
            candidate
        )
    )

    documents = []

    for document_url in document_urls[
        :MAX_DOCUMENTS_PER_TENDER
    ]:
        document = (
            document_manager
            .download_and_extract(
                document_url,
                candidate.stable_key(),
            )
        )

        documents.append(
            document
        )

    evidence = extract_eligibility(
        candidate,
        documents,
    )

    decision = evaluate_qualification(
        evidence
    )

    return (
        evidence,
        documents,
        decision,
    )


# =============================================================================
# 19. MAIN PIPELINE
# =============================================================================

def run_pipeline():
    started_at = time.time()

    log.info(
        "=" * 70
    )
    log.info(
        "%s started",
        APP_NAME,
    )
    log.info(
        "=" * 70
    )

    validate_environment()

    portals = load_master_portals()

    log.info(
        "Loaded %d active tender portals.",
        len(portals),
    )

    crawler = TenderCrawler()

    all_candidates = []
    health = {}

    non_gem_portals = [
        portal
        for portal in portals
        if "gem.gov.in" not in portal.url.lower()
    ]

    with ThreadPoolExecutor(
        max_workers=CRAWL_WORKERS
    ) as executor:

        futures = {
            executor.submit(
                crawler.crawl,
                portal,
            ): portal
            for portal in non_gem_portals
        }

        for future in as_completed(
            futures
        ):
            portal = futures[
                future
            ]

            try:
                items, status = future.result()

            except Exception as error:
                items = []
                status = CrawlHealth(
                    portal=portal.portal,
                    url=portal.url,
                    status="PARSER_FAILED",
                    discovered_count=0,
                    message=str(error)[:250],
                )

            health[
                portal.portal
            ] = status

            all_candidates.extend(
                items
            )

    gem_candidates, gem_health = crawl_gem()

    all_candidates.extend(
        gem_candidates
    )

    health[
        "Government e-Marketplace (GeM)"
    ] = gem_health

    unique_map = {}

    for candidate in all_candidates:
        if not valid_tender_title(
            candidate.title
        ):
            log.info(
                "Ignoring invalid tender title: %s",
                candidate.title,
            )
            continue

        key = candidate.stable_key()
        unique_map[key] = candidate

    candidates = list(
        unique_map.values()
    )

    log.info(
        "Raw valid tender candidates discovered: %d",
        len(candidates),
    )

    candidates = filter_event_candidates(
        candidates
    )

    log.info(
        "Relevant event/exhibition candidates: %d",
        len(candidates),
    )

    google_client = get_gspread_client()

    spreadsheet = google_client.open_by_key(
        SPREADSHEET_ID
    )

    active_sheet = ensure_worksheet(
        spreadsheet,
        "Active_Tenders",
        rows=10000,
        cols=35,
    )

    portal_sheet = ensure_worksheet(
        spreadsheet,
        "Portal_Directory",
        rows=max(
            500,
            len(portals) + 100,
        ),
        cols=15,
    )

    active_sheet.update(
        values=[
            ACTIVE_HEADERS
        ],
        range_name="A1:Z1",
    )

    clean_active_tenders_sheet(
        active_sheet
    )

    existing_keys = existing_tender_keys(
        active_sheet
    )

    document_manager = DocumentManager()

    new_rows = []

    for candidate in candidates:
        tender_key = candidate.stable_key()

        if tender_key in existing_keys:
            continue

        log.info(
            "Processing tender: %s | %s",
            (
                candidate.tender_id
                or
                tender_key
            ),
            candidate.title[:100],
        )

        evidence, documents, decision = process_candidate(
            candidate,
            document_manager,
        )

        downloaded_files = [
            document.local_path
            for document in documents
            if document.local_path
        ]

        document_urls = list(
            dict.fromkeys(
                candidate.discovered_doc_urls
                + [
                    document.url
                    for document in documents
                    if document.url
                ]
            )
        )

        experience_parts = []

        single = evidence.get(
            "single_similar_work_required_inr"
        )
        two = evidence.get(
            "two_similar_works_each_required_inr"
        )
        three = evidence.get(
            "three_similar_works_each_required_inr"
        )

        if single is not None:
            experience_parts.append(
                f"1 work >= ₹{single:,.0f}"
            )

        if two is not None:
            experience_parts.append(
                f"2 works each >= ₹{two:,.0f}"
            )

        if three is not None:
            experience_parts.append(
                f"3 works each >= ₹{three:,.0f}"
            )

        experience_text = (
            "; ".join(
                experience_parts
            )
            or
            evidence.get(
                "experience_requirement_text"
            )
            or
            "NOT VERIFIED / NOT STATED"
        )

        compliance_flags = []

        if evidence.get(
            "physical_submission_required"
        ) is True:
            compliance_flags.append(
                "Physical submission required"
            )

        if evidence.get(
            "qcbs_or_technical_pitch"
        ) is True:
            compliance_flags.append(
                "QCBS / technical pitch"
            )

        if evidence.get(
            "named_celebrity_or_artist_mandate"
        ) is True:
            compliance_flags.append(
                "Artist / celebrity mandate"
            )

        if evidence.get(
            "non_core_scope"
        ) is True:
            compliance_flags.append(
                "Non-core scope warning"
            )

        compliance_text = (
            "; ".join(
                compliance_flags
            )
            or
            "No special compliance flag extracted"
        )

        remarks = (
            decision.reason
            + " Action: "
            + decision.action_plan
        )

        new_rows.append([
            dt.datetime.now(
                tz=IST
            ).strftime(
                "%Y-%m-%d"
            ),

            tender_key,

            evidence.get(
                "tender_reference"
            )
            or
            candidate.tender_id
            or
            "NOT VERIFIED",

            candidate.portal,

            candidate.state,

            evidence.get(
                "organisation"
            )
            or
            candidate.organization
            or
            "NOT VERIFIED",

            evidence.get(
                "scope_summary"
            )
            or
            candidate.title,

            candidate.category,

            (
                evidence.get(
                    "estimated_value_inr"
                )
                if
                evidence.get(
                    "estimated_value_inr"
                )
                is not None
                else
                "NOT VERIFIED / NOT STATED"
            ),

            (
                evidence.get(
                    "emd_inr"
                )
                if
                evidence.get(
                    "emd_inr"
                )
                is not None
                else
                "NOT VERIFIED / NOT STATED"
            ),

            evidence.get(
                "emd_exemption_text"
            )
            or
            "NOT VERIFIED / NOT STATED",

            evidence.get(
                "submission_deadline"
            )
            or
            candidate.deadline_raw
            or
            "NOT VERIFIED",

            (
                evidence.get(
                    "average_turnover_required_inr"
                )
                if
                evidence.get(
                    "average_turnover_required_inr"
                )
                is not None
                else
                "NOT VERIFIED / NOT STATED"
            ),

            experience_text,

            evidence.get(
                "entity_type_restriction"
            )
            or
            "NOT VERIFIED / NOT STATED",

            (
                "Yes"
                if
                evidence.get(
                    "consortium_or_jv_allowed"
                )
                is True
                else
                "No"
                if
                evidence.get(
                    "consortium_or_jv_allowed"
                )
                is False
                else
                "NOT VERIFIED / NOT STATED"
            ),

            compliance_text,

            "\n".join(
                document_urls
            ),

            "\n".join(
                downloaded_files
            ),

            evidence.get(
                "pre_bid_date"
            )
            or
            "NOT VERIFIED / NOT STATED",

            candidate.detail_url
            or
            candidate.source_url,

            decision.status,

            remarks,

            "\n".join(
                (
                    evidence.get(
                        "evidence_quotes"
                    )
                    or []
                )[:10]
            ),

            evidence.get(
                "extraction_status",
                "UNKNOWN",
            ),

            evidence.get(
                "ai_model_used",
                "",
            ),
        ])

        existing_keys.add(
            tender_key
        )

    if new_rows:
        active_sheet.append_rows(
            new_rows,
            value_input_option="RAW",
        )

        log.info(
            "Successfully added %d new tender records.",
            len(new_rows),
        )
    else:
        log.info(
            "No new tender records to append."
        )

    sync_portal_directory(
        portal_sheet,
        portals,
        health,
    )

    elapsed = (
        time.time()
        - started_at
    )

    log.info(
        "=" * 70
    )

    log.info(
        "Tender Intelligence Agent completed "
        "in %.1f seconds.",
        elapsed,
    )

    log.info(
        "=" * 70
    )


# =============================================================================
# RUN
# =============================================================================

if __name__ == "__main__":
    run_pipeline()
