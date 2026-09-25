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

CODE_VERSION = "2026-09-25-SERVICE-FILTER-V6"

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
    "https://bidplus.gem.gov.in/all-bids",
)

GEM_GTE_URL = os.getenv(
    "GEM_GTE_URL",
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
        "100",
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


GENERIC_JS_FALLBACK = (
    os.getenv(
        "GENERIC_JS_FALLBACK",
        "true",
    ).strip().lower()
    in {"1", "true", "yes", "y"}
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


def contains_phrase(text, phrase):
    """
    Boundary-aware phrase matcher.

    Prevents false positives such as:
      "stall" matching "install"
      "event" matching "events" only when not intended
    while still supporting multi-word phrases.
    """
    haystack = normalize_space(
        text or ""
    ).lower()

    needle = normalize_space(
        phrase or ""
    ).lower()

    if not haystack or not needle:
        return False

    pattern = (
        r"(?<![a-z0-9])"
        + re.escape(needle)
        + r"(?![a-z0-9])"
    )

    return re.search(
        pattern,
        haystack,
        flags=re.I,
    ) is not None


def contains_any_phrase(text, phrases):
    return any(
        contains_phrase(
            text,
            phrase,
        )
        for phrase in phrases
    )


def keyword_hits(text):
    return [
        keyword
        for keyword in EVENT_KEYWORDS
        if contains_phrase(
            text,
            keyword,
        )
    ]


def infer_category(title):
    """
    Classify only the service family relevant to an event/experience agency.
    Event management is intentionally treated as a SERVICE category.
    """
    text = normalize_space(
        title or ""
    ).lower()

    if contains_any_phrase(
        text,
        [
            "event management",
            "event management agency",
            "event agency",
            "event organiser",
            "event organizer",
            "event production",
            "event execution",
            "event services",
            "event partner",
            "annual day",
            "foundation day",
            "award ceremony",
            "dealer meet",
            "launch event",
            "product launch",
            "brand activation",
            "experiential marketing",
        ],
    ):
        return "Services - Event Management"

    if contains_any_phrase(
        text,
        [
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
        ],
    ):
        return "Services - Exhibition / Stall"

    if contains_any_phrase(
        text,
        [
            "conference management",
            "organising conference",
            "organizing conference",
            "organisation of conference",
            "organization of conference",
            "conduct of conference",
            "organising convention",
            "organizing convention",
            "convention management",
            "summit management",
            "conclave management",
            "seminar management",
            "organising seminar",
            "organizing seminar",
        ],
    ):
        return "Services - Conference / Convention"

    if contains_any_phrase(
        text,
        [
            "advertising agency",
            "creative agency",
            "communication agency",
            "publicity campaign",
            "outreach campaign",
            "iec campaign",
            "media campaign",
            "social media agency",
            "social media and communication agency",
            "pr agency",
        ],
    ):
        return "Services - Creative / Advertising / Outreach"

    if contains_any_phrase(
        text,
        [
            "audio visual coverage",
            "audio visual production",
            "av production",
            "photography",
            "videography",
            "stage setup",
            "stage production",
            "sound and light",
            "event decoration",
            "tentage for event",
        ],
    ):
        return "Services - AV / Production"

    if (
        contains_phrase(
            text,
            "empanelment",
        )
        or
        contains_phrase(
            text,
            "empanel",
        )
    ):
        return "Services - Empanelment"

    if contains_any_phrase(
        text,
        [
            "mela management",
            "festival management",
            "cultural event",
            "cultural programme",
        ],
    ):
        return "Services - Festival / Mela / Cultural"

    return "Services - Other Event Related"


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
    Strict event/exhibition/creative relevance check.

    Uses boundary-aware phrase matching so "stall" does not match "install".
    """

    title = normalize_space(
        title or ""
    ).lower()

    if not title:
        return False

    if title in GENERIC_NAV_TITLES:
        return False

    if contains_any_phrase(
        title,
        FALSE_POSITIVE_PHRASES,
    ):
        return False

    if contains_any_phrase(
        title,
        NON_EVENT_EXCLUSIONS,
    ):
        return False

    if contains_any_phrase(
        title,
        EVENT_SERVICE_PHRASES,
    ):
        return True

    if (
        contains_phrase(
            title,
            "empanelment",
        )
        or
        contains_phrase(
            title,
            "empanel",
        )
    ):
        if contains_any_phrase(
            title,
            [
                "event",
                "exhibition",
                "advertising",
                "creative",
                "media",
                "publicity",
                "communication",
                "experiential",
            ],
        ):
            return True

    has_event_object = contains_any_phrase(
        title,
        EVENT_OBJECT_TERMS,
    )

    has_service_action = contains_any_phrase(
        title,
        SERVICE_ACTION_TERMS,
    )

    return (
        has_event_object
        and
        has_service_action
    )


# =============================================================================
# EVENT-SERVICE QUALITY FILTER
# =============================================================================

SERVICE_GOODS_EXCLUSIONS = [
    "supply of audio visual",
    "procurement of audio visual",
    "audio visual equipment",
    "audio visual equipments",
    "audio visual system",
    "av equipment",
    "video conferencing system",
    "conference system and accessories",
    "interactive display",
    "display and video conferencing",
    "supply of bags",
    "supply of airfryer",
    "supply of chairs",
    "supply of gifts",
    "supply of mementos",
    "supply of prizes",
    "procurement of prizes",
    "procurement of gifts",
    "procurement of banners",
    "procurement of shawls",
    "batteries",
    "ups",
    "diesel generator",
    "dg set",
    "fire alarm",
    "sprinkler",
    "fire hydrant",
]

SERVICE_RIGHTS_EXCLUSIONS = [
    "temporary allotment of open space",
    "allotment of open space",
    "allotment of space",
    "licensing of space",
    "licence of space",
    "license of space",
    "joyrides",
    "meenabazar",
    "parking rights",
    "kiosk allotment",
]

SERVICE_INFRA_EXCLUSIONS = [
    "conference hall",
    "revamping of conference hall",
    "interior works",
    "civil works",
    "construction of",
    "maintenance and repair",
    "annual maintenance",
    "day-to-day maintenance",
    "horticulture works",
    "road works",
]

NON_SPECIFIC_TENDER_TITLES = [
    "latest advertising agency tenders",
    "latest audio visual tenders",
    "latest audio visual accessory tenders",
    "latest audio visual equipment tenders",
    "latest audio visual instrument tenders",
    "audio visual equipments tenders from india",
    "advertising, film & media campaign govt. tenders",
    "advertising film and media campaign tenders",
    "event management tenders",
    "india expo centre & mart tenders",
    "india international convention and exhibition centre",
]

HIGH_CONFIDENCE_SERVICE_PHRASES = [
    "event management",
    "event management agency",
    "event agency",
    "event organiser",
    "event organizer",
    "event production",
    "event execution",
    "event services",
    "event partner",
    "empanelment of event",
    "empanelment of event management",
    "hiring of event",
    "selection of event management",
    "organising all the events",
    "organizing all the events",

    "exhibition stall",
    "exhibition pavilion",
    "stall design",
    "stall fabrication",
    "stall construction",
    "pavilion design",
    "pavilion fabrication",
    "pavilion construction",
    "exhibition management",
    "expo management",
    "trade fair management",

    "conference management",
    "organising conference",
    "organizing conference",
    "organisation of conference",
    "organization of conference",
    "organising convention",
    "organizing convention",
    "convention management",
    "summit management",
    "conclave management",
    "seminar management",
    "organising seminar",
    "organizing seminar",

    "advertising agency",
    "creative agency",
    "communication agency",
    "social media agency",
    "social media and communication agency",
    "publicity campaign",
    "outreach campaign",
    "iec campaign",
    "media campaign",
    "pr agency",

    "audio visual coverage",
    "audio visual production",
    "photography videography",
    "photography and videography",
    "stage production",
    "stage setup",
    "event decoration",
    "brand activation",
    "experiential marketing",

    "mela management",
    "festival management",
    "cultural event management",
]

def _extract_explicit_dates(text):
    """
    Extract common Indian tender dates from title/deadline text.
    Returns timezone-aware date objects where parsing is reliable.
    """
    text = normalize_space(
        text or ""
    )

    output = []

    numeric_patterns = [
        r"\b(\d{1,2})[/-](\d{1,2})[/-](20\d{2})\b",
        r"\b(20\d{2})[/-](\d{1,2})[/-](\d{1,2})\b",
    ]

    for match in re.finditer(
        numeric_patterns[0],
        text,
    ):
        try:
            output.append(
                dt.date(
                    int(match.group(3)),
                    int(match.group(2)),
                    int(match.group(1)),
                )
            )
        except ValueError:
            pass

    for match in re.finditer(
        numeric_patterns[1],
        text,
    ):
        try:
            output.append(
                dt.date(
                    int(match.group(1)),
                    int(match.group(2)),
                    int(match.group(3)),
                )
            )
        except ValueError:
            pass

    month_map = {
        "jan": 1, "january": 1,
        "feb": 2, "february": 2,
        "mar": 3, "march": 3,
        "apr": 4, "april": 4,
        "may": 5,
        "jun": 6, "june": 6,
        "jul": 7, "july": 7,
        "aug": 8, "august": 8,
        "sep": 9, "sept": 9, "september": 9,
        "oct": 10, "october": 10,
        "nov": 11, "november": 11,
        "dec": 12, "december": 12,
    }

    for match in re.finditer(
        r"\b(\d{1,2})(?:st|nd|rd|th)?\s+"
        r"(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
        r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
        r"[,\s]+(20\d{2})\b",
        text,
        flags=re.I,
    ):
        try:
            month = month_map[
                match.group(2).lower()
            ]
            output.append(
                dt.date(
                    int(match.group(3)),
                    month,
                    int(match.group(1)),
                )
            )
        except (ValueError, KeyError):
            pass

    return output


def is_expired_candidate(candidate):
    """
    Reject clearly expired/historical tenders.
    Unknown dates are not rejected.
    """
    today = dt.datetime.now(
        tz=IST
    ).date()

    title = normalize_space(
        candidate.title or ""
    )

    deadline_text = normalize_space(
        candidate.deadline_raw or ""
    )

    combined = (
        title
        + " "
        + deadline_text
    )

    dates = _extract_explicit_dates(
        combined
    )

    # If title explicitly says Last Date/closing/deadline and that date is old,
    # this is safe to reject.
    if re.search(
        r"(last\s*date|closing\s*date|submission\s*deadline|bid\s*end\s*date)",
        combined,
        flags=re.I,
    ):
        if dates and max(dates) < today:
            return True

    # Reject clearly historic archive entries from previous years when no
    # current-year reference is present.
    current_year = today.year
    years = {
        int(year)
        for year in re.findall(
            r"\b(20\d{2})\b",
            title,
        )
    }

    if (
        years
        and
        max(years) < current_year
    ):
        return True

    tender_id = normalize_space(
        candidate.tender_id or ""
    )

    id_years = {
        int(year)
        for year in re.findall(
            r"\b(20\d{2})\b",
            tender_id,
        )
    }

    if (
        id_years
        and
        max(id_years) < current_year - 1
    ):
        return True

    return False


def is_event_service_candidate(candidate):
    """
    Final business-fit filter for Soul Events and Consultancy.

    Keep procurement of EVENT/EXHIBITION/CREATIVE/AV SERVICES.
    Reject goods/equipment purchases, civil/interior works, space allotments,
    archive/category pages and clearly expired historical tenders.
    """
    title = normalize_space(
        candidate.title or ""
    ).lower()

    detail_url = (
        candidate.detail_url
        or candidate.source_url
        or ""
    ).lower()

    if not title:
        return False, "blank title"

    if any(
        phrase in title
        for phrase in NON_SPECIFIC_TENDER_TITLES
    ):
        return False, "category/listing page, not a specific tender"

    if any(
        token in detail_url
        for token in [
            "/archive",
            "/archived",
            "archive-tenders",
            "archived-tenders",
        ]
    ):
        return False, "archive/historical tender URL"

    if contains_any_phrase(
        title,
        SERVICE_RIGHTS_EXCLUSIONS,
    ):
        return False, "space/allotment/rights opportunity, not event-management service"

    if contains_any_phrase(
        title,
        SERVICE_INFRA_EXCLUSIONS,
    ):
        return False, "civil/interior/maintenance scope, not event-management service"

    high_confidence_service = contains_any_phrase(
        title,
        HIGH_CONFIDENCE_SERVICE_PHRASES,
    )

    if contains_any_phrase(
        title,
        SERVICE_GOODS_EXCLUSIONS,
    ) and not high_confidence_service:
        return False, "goods/equipment procurement, not service"

    # Generic supply/procurement/SITC scopes are rejected unless the same
    # title explicitly contains a high-confidence event-service phrase.
    if (
        contains_any_phrase(
            title,
            [
                "supply of",
                "procurement of",
                "purchase of",
                "sitc of",
                "supply installation testing commissioning",
                "supply, installation, testing and commissioning",
            ],
        )
        and
        not high_confidence_service
    ):
        return False, "supply/procurement scope without event-service mandate"

    if is_expired_candidate(
        candidate
    ):
        return False, "expired/historical tender"

    if not high_confidence_service:
        return False, "event word present but no clear service mandate"

    return True, "event/exhibition/creative service procurement"


PROCUREMENT_STRONG_TERMS = [
    "request for proposal",
    "request for quotation",
    "expression of interest",
    "notice inviting tender",
    "notice inviting bid",
    "tender",
    "bid",
    "rfp",
    "rfq",
    "eoi",
    "nit",
    "empanelment",
    "selection of agency",
    "selection of event",
    "appointment of agency",
    "hiring of agency",
    "engagement of agency",
]

NON_PROCUREMENT_CONTENT_TERMS = [
    "press release",
    "news update",
    "news updates",
    "media release",
    "speech",
    "statement by",
    "participated in",
    "participates in",
    "inaugurated",
    "celebrated",
    "celebrates",
    "gallery",
    "photo gallery",
    "annual report",
    "investor",
    "brochure",
    "success story",
]

NON_PROCUREMENT_URL_TERMS = [
    "/news/",
    "/news_",
    "/news-",
    "/news_marquee/",
    "/press-release",
    "/press_release",
    "/media/",
    "/gallery/",
    "/blog/",
    "/events/",
    "/speech",
    "/annual-report",
    "/investor",
    "/industry-projects/",
]


def procurement_intent_score(candidate):
    """
    Score whether a discovered item is actually a procurement opportunity,
    rather than a news/article/content page that merely mentions an event.

    A score >= 3 is considered procurement-intent-positive.
    """

    title = normalize_space(
        candidate.title or ""
    ).lower()

    tender_id = normalize_space(
        candidate.tender_id or ""
    )

    detail_url = (
        candidate.detail_url
        or candidate.source_url
        or ""
    ).lower()

    source_url = (
        candidate.source_url
        or ""
    ).lower()

    score = 0
    reasons = []

    # Strongest signal: a recognizable tender / bid reference.
    if tender_id and tender_id.upper() not in {
        "NOT VERIFIED",
        "NOT STATED",
    }:
        if re.search(
            r"(?:GEM/\d{4}/B/\d+|(?:RFP|RFQ|EOI|NIT|TENDER|BID)[\s:/_-]*[A-Za-z0-9._/-]{2,}|\d{4}_[A-Za-z0-9_-]+_\d+_\d+)",
            tender_id,
            flags=re.I,
        ):
            score += 4
            reasons.append("recognizable tender reference")
        elif len(tender_id) >= 5:
            score += 2
            reasons.append("non-empty tender reference")

    if any(
        term in title
        for term in PROCUREMENT_STRONG_TERMS
    ):
        score += 3
        reasons.append("procurement wording in title")

    if any(
        token in detail_url
        for token in [
            "showbiddocument",
            "/tender/",
            "/tenders/",
            "tenderid",
            "bidplus",
            "eprocure",
            "etender",
            "e-tender",
            "/rfp/",
            "/eoi/",
            "/nit/",
            "procurement",
        ]
    ):
        score += 2
        reasons.append("procurement-style detail URL")

    if any(
        token in source_url
        for token in [
            "eprocure",
            "etender",
            "tenders.gov",
            "gem.gov.in",
            "bidplus",
            "tendernews",
            "tendertiger",
            "indiantenders",
            "tenderdetail",
            "tendershark",
        ]
    ):
        score += 1
        reasons.append("known tender portal context")

    if is_relevant_event_title(
        candidate.title
    ):
        score += 1
        reasons.append("event-service scope")

    if any(
        term in title
        for term in NON_PROCUREMENT_CONTENT_TERMS
    ):
        score -= 5
        reasons.append("news/content wording")

    if any(
        term in detail_url
        for term in NON_PROCUREMENT_URL_TERMS
    ):
        score -= 5
        reasons.append("news/content URL")

    # Category/index pages are not specific tender opportunities.
    if any(
        term in detail_url
        for term in [
            "/industry-projects/",
            "/category/",
            "/categories/",
            "/search",
        ]
    ) and not tender_id:
        score -= 3
        reasons.append("category/search page without tender reference")

    return score, reasons


def is_procurement_candidate(candidate):
    score, _ = procurement_intent_score(
        candidate
    )
    return score >= 3


def _safe_number_from_text(value):
    """
    Parse a monetary value into INR when possible.
    Supports explicit rupees as well as lakh/crore wording.
    """
    if value is None:
        return None

    text_value = normalize_space(
        str(value)
    ).lower()

    if not text_value:
        return None

    match = re.search(
        r"(?:₹|rs\.?|inr)?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*(crore|cr|lakh|lac|lakhs|lacs)?",
        text_value,
        flags=re.I,
    )

    if not match:
        return None

    try:
        number = float(
            match.group(1).replace(
                ",",
                "",
            )
        )
    except ValueError:
        return None

    unit = (
        match.group(2)
        or ""
    ).lower()

    if unit in {
        "crore",
        "cr",
    }:
        number *= 10_000_000

    elif unit in {
        "lakh",
        "lac",
        "lakhs",
        "lacs",
    }:
        number *= 100_000

    return int(
        round(number)
    )


def _extract_labeled_value(text_value, labels, max_chars=120):
    if not text_value:
        return ""

    for label in labels:
        pattern = (
            rf"(?:{label})\s*"
            rf"(?:[:\-–]\s*)?"
            rf"([^\n\r]{{1,{max_chars}}})"
        )

        match = re.search(
            pattern,
            text_value,
            flags=re.I,
        )

        if match:
            return normalize_space(
                match.group(1)
            )

    return ""


def extract_deterministic_metadata(
    candidate,
    documents,
):
    """
    Extract basic tender metadata without AI.

    These fields should not depend on OpenAI availability:
    tender ref, organisation, deadlines, EMD, estimated value,
    turnover hint and pre-bid date.
    """

    chunks = [
        candidate.title or "",
        candidate.tender_id or "",
        candidate.deadline_raw or "",
        candidate.estimated_value_raw or "",
    ]

    for document in documents:
        if document.text:
            chunks.append(
                document.text[:250_000]
            )

    combined = "\n".join(
        chunks
    )

    refs = extract_ref_candidates(
        combined
    )

    tender_reference = (
        candidate.tender_id
        if (
            candidate.tender_id
            and
            candidate.tender_id.upper()
            not in {
                "NOT VERIFIED",
                "NOT STATED",
            }
        )
        else (
            refs[0]
            if refs
            else None
        )
    )

    deadline = _extract_labeled_value(
        combined,
        [
            r"bid\s*end\s*date(?:\s*/\s*time)?",
            r"bid\s*submission\s*end\s*date(?:\s*/\s*time)?",
            r"last\s*date(?:\s*and\s*time)?\s*of\s*submission",
            r"last\s*date\s*for\s*submission",
            r"submission\s*deadline",
            r"closing\s*date(?:\s*/\s*time)?",
            r"tender\s*closing\s*date",
        ],
        max_chars=80,
    ) or (
        candidate.deadline_raw
        if (
            candidate.deadline_raw
            and candidate.deadline_raw != "NOT VERIFIED"
        )
        else ""
    )

    pre_bid = _extract_labeled_value(
        combined,
        [
            r"pre[\s-]*bid\s*(?:meeting|conference)?\s*(?:date)?",
            r"pre[\s-]*bid\s*date",
        ],
        max_chars=80,
    )

    emd_text = _extract_labeled_value(
        combined,
        [
            r"emd",
            r"earnest\s*money\s*deposit",
            r"bid\s*security",
        ],
        max_chars=100,
    )

    estimated_text = _extract_labeled_value(
        combined,
        [
            r"estimated\s*(?:bid|tender|contract)?\s*value",
            r"estimated\s*cost",
            r"tender\s*value",
            r"bid\s*value",
            r"contract\s*value",
        ],
        max_chars=100,
    ) or (
        candidate.estimated_value_raw
        if (
            candidate.estimated_value_raw
            and candidate.estimated_value_raw != "NOT VERIFIED"
        )
        else ""
    )

    turnover_text = _extract_labeled_value(
        combined,
        [
            r"average\s*annual\s*turnover",
            r"minimum\s*annual\s*turnover",
            r"annual\s*turnover",
            r"minimum\s*turnover",
        ],
        max_chars=140,
    )

    organisation = _extract_labeled_value(
        combined,
        [
            r"organisation\s*name",
            r"organization\s*name",
            r"buyer\s*organisation",
            r"buyer\s*organization",
            r"department\s*name",
            r"ministry\s*/\s*state\s*name",
        ],
        max_chars=120,
    )

    if not organisation:
        organisation = (
            candidate.organization
            or ""
        )

    return {
        "tender_reference": tender_reference,
        "organisation": organisation or None,
        "submission_deadline": deadline or None,
        "pre_bid_date": pre_bid or None,
        "emd_inr": _safe_number_from_text(
            emd_text
        ),
        "estimated_value_inr": _safe_number_from_text(
            estimated_text
        ),
        "average_turnover_required_inr": _safe_number_from_text(
            turnover_text
        ),
        "deterministic_emd_text": emd_text or None,
        "deterministic_estimated_value_text": estimated_text or None,
        "deterministic_turnover_text": turnover_text or None,
    }


def merge_deterministic_metadata(
    candidate,
    documents,
    evidence,
):
    deterministic = extract_deterministic_metadata(
        candidate,
        documents,
    )

    merged = dict(
        evidence or {}
    )

    # Only fill blanks; never overwrite an AI-extracted explicit value.
    for field_name in [
        "tender_reference",
        "organisation",
        "submission_deadline",
        "pre_bid_date",
        "emd_inr",
        "estimated_value_inr",
        "average_turnover_required_inr",
    ]:
        if (
            merged.get(
                field_name
            )
            in (
                None,
                "",
                "NOT VERIFIED",
                "NOT VERIFIED / NOT STATED",
            )
        ):
            value = deterministic.get(
                field_name
            )

            if value not in (
                None,
                "",
            ):
                merged[
                    field_name
                ] = value

    merged[
        "deterministic_metadata"
    ] = deterministic

    return merged



def is_relevant_event_tender(candidate):
    # Relevance must be based on the actual tender title.
    # Do NOT use inferred category or full webpage text here.
    return is_relevant_event_title(
        candidate.title
    )



def extract_gem_bid_numbers(*chunks):
    """
    Extract public GeM bid numbers from rendered text or page HTML.
    """
    output = []

    for chunk in chunks:
        if not chunk:
            continue

        output.extend(
            re.findall(
                r"GEM/\d{4}/B/\d+",
                chunk,
                flags=re.I,
            )
        )

    return list(
        dict.fromkeys(
            item.upper()
            for item in output
        )
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
                    (
                        "SUCCESS"
                        if items
                        else "NO_MATCHES"
                    ),
                    len(items),
                    (
                        ""
                        if items
                        else "Portal was reached, but no matching tender candidates were discovered."
                    ),
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
        Crawl GePNIC/CPPP-style portals.

        Important:
        - Active Tender search pages often show a CAPTCHA before returning rows.
        - The public portal home page usually exposes a Latest Tenders table.
        - Therefore we crawl BOTH the active-tender endpoint and the public
          home page, then follow real pagination/detail links where available.
        """

        base = portal.url.rstrip("/")

        seed_urls = []

        # Exact URL from master file.
        seed_urls.append(base)

        # If master points directly to an /app endpoint, also try its root.
        if "/app" in base.lower():
            root_guess = base.split("/app", 1)[0]
            seed_urls.append(root_guess)

            separator = "&" if "?" in base else "?"
            seed_urls.append(
                base
                + separator
                + "page=FrontEndLatestActiveTenders"
                + "&service=page"
            )

        # Common GePNIC layouts.
        seed_urls.extend([
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
        ])

        seed_urls = list(
            dict.fromkeys(
                seed_urls
            )
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
            next_urls = []

            for row in soup.find_all(
                "tr"
            ):
                row_text = normalize_space(
                    row.get_text(
                        " ",
                        strip=True,
                    )
                )

                if not row_text:
                    continue

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

                anchor_items = []

                for anchor in row.find_all(
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

                    if (
                        href
                        and
                        not href.lower().startswith(
                            "javascript:"
                        )
                    ):
                        anchor_items.append(
                            (
                                label,
                                href,
                            )
                        )

                title_candidates = [
                    cell
                    for cell in cells
                    if (
                        valid_tender_title(
                            cell
                        )
                        and
                        keyword_hits(
                            cell
                        )
                    )
                ]

                title_candidates.extend([
                    label
                    for label, _ in anchor_items
                    if (
                        valid_tender_title(
                            label
                        )
                        and
                        keyword_hits(
                            label
                        )
                    )
                ])

                if title_candidates:
                    title = max(
                        title_candidates,
                        key=len,
                    )
                else:
                    # If event wording exists only in the combined row,
                    # retain the best non-generic cell.
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

                for label, href in anchor_items:
                    combined = (
                        label
                        + " "
                        + href
                    ).lower()

                    if is_downloadable_url(
                        href
                    ):
                        documents.append(
                            href
                        )

                    elif (
                        "component=%24directlink"
                        in href.lower()
                        or
                        any(
                            term
                            in combined
                            for term
                            in TENDERISH_TERMS
                        )
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

            # Follow pagination and direct tender-list pages.
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

                if (
                    not href
                    or
                    href.lower().startswith(
                        "javascript:"
                    )
                ):
                    continue

                label_lower = label.lower()
                href_lower = href.lower()

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
                    or
                    "frontendlatestactivetenders"
                    in href_lower
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
        seen_tender_keys = set()
        visited_pages = set()
        queue = list(
            seed_urls
        )

        page_no = 0

        while (
            queue
            and
            page_no < MAX_PAGES_PER_PORTAL
        ):
            current_url = queue.pop(0)

            if current_url in visited_pages:
                continue

            visited_pages.add(
                current_url
            )

            try:
                response = self.http.get(
                    current_url,
                    timeout=REQUEST_TIMEOUT,
                    allow_redirects=True,
                )

                if (
                    response.status_code != 200
                    or
                    len(response.text) < 300
                ):
                    continue

            except requests.RequestException:
                continue

            page_no += 1

            page_results, next_urls = parse_page(
                response.text,
                response.url,
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
                "%s | GePNIC page %d | %d new matching tenders | total %d | %s",
                portal.portal,
                page_no,
                new_count,
                len(results),
                response.url,
            )

            for next_url in next_urls:
                if (
                    next_url not in visited_pages
                    and
                    next_url not in queue
                ):
                    queue.append(
                        next_url
                    )

        if page_no == 0:
            raise RuntimeError(
                "No usable GePNIC page returned."
            )

        log.info(
            "%s | GePNIC crawl complete | pages=%d | tenders=%d",
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
        Crawl generic tender websites by traversing tender/procurement
        listing pages before treating links as individual tenders.

        Previous versions only inspected the landing page and often mistook
        a "Tenders" navigation link for a tender itself. This version:
        - recursively queues tender/procurement listing pages
        - parses event-related table rows/cards on every visited page
        - follows pagination
        - opens candidate detail pages only when the candidate itself is
          event-relevant.
        """

        first_response = self.http.get(
            portal.url,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
        )

        first_response.raise_for_status()

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

        same_host = urllib.parse.urlsplit(
            first_response.url
        ).netloc.lower()

        def enqueue(
            url,
        ):
            if not url:
                return

            parsed = urllib.parse.urlsplit(
                url
            )

            if (
                parsed.netloc
                and
                parsed.netloc.lower()
                != same_host
            ):
                return

            if (
                url not in visited
                and
                url not in queue
                and
                len(queue) < 500
            ):
                queue.append(
                    url
                )

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
                    response_url = page_url
                else:
                    response = self.http.get(
                        page_url,
                        timeout=REQUEST_TIMEOUT,
                        allow_redirects=True,
                    )

                    if response.status_code != 200:
                        continue

                    html = response.text
                    response_url = response.url

            except requests.RequestException:
                continue

            page_no += 1

            soup = BeautifulSoup(
                html,
                "html.parser",
            )

            new_count = 0

            # 1) Parse table rows as tender records.
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

                anchor_items = []

                for anchor in row.find_all(
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
                        response_url,
                        anchor.get(
                            "href"
                        ),
                    )

                    if href:
                        anchor_items.append(
                            (
                                label,
                                href,
                            )
                        )

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

                title_candidates = [
                    cell
                    for cell in cells
                    if (
                        valid_tender_title(
                            cell
                        )
                        and
                        keyword_hits(
                            cell
                        )
                    )
                ]

                title_candidates.extend([
                    label
                    for label, _ in anchor_items
                    if (
                        valid_tender_title(
                            label
                        )
                        and
                        keyword_hits(
                            label
                        )
                    )
                ])

                if not title_candidates:
                    continue

                title = max(
                    title_candidates,
                    key=len,
                )

                detail_url = response_url
                documents = []

                for label, href in anchor_items:
                    if is_downloadable_url(
                        href
                    ):
                        documents.append(
                            href
                        )
                    else:
                        detail_url = href

                refs = extract_ref_candidates(
                    row_text
                )

                candidate = TenderCandidate(
                    portal=portal.portal,
                    state=portal.state,
                    source_url=response_url,
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
                    detail_url=detail_url,
                    discovered_doc_urls=list(
                        dict.fromkeys(
                            documents
                        )
                    )[:MAX_DOCUMENTS_PER_TENDER],
                    keyword_hits=hits,
                )

                key = candidate.stable_key()

                if key not in seen_tender_keys:
                    seen_tender_keys.add(
                        key
                    )
                    results.append(
                        candidate
                    )
                    new_count += 1

            # 2) Traverse navigation/listing links and inspect likely
            #    individual tender links.
            candidate_links = []

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
                    response_url,
                    anchor.get(
                        "href"
                    ),
                )

                if (
                    not href
                    or
                    href.lower().startswith(
                        "javascript:"
                    )
                ):
                    continue

                combined = (
                    label
                    + " "
                    + href
                ).lower()

                label_lower = label.lower()

                is_pagination = (
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
                )

                is_listing = any(
                    term in combined
                    for term in [
                        "tender",
                        "procurement",
                        "eprocure",
                        "e-tender",
                        "etender",
                        "rfp",
                        "eoi",
                        "bid",
                        "notice inviting",
                    ]
                )

                if (
                    is_pagination
                    or
                    is_listing
                ):
                    enqueue(
                        href
                    )

                if (
                    keyword_hits(
                        label
                    )
                    and
                    valid_tender_title(
                        label
                    )
                ):
                    candidate_links.append(
                        (
                            label,
                            href,
                        )
                    )

            # 3) Open event-looking detail links and extract metadata/docs.
            for label, href in candidate_links[:MAX_GENERIC_LINKS]:
                if is_downloadable_url(
                    href
                ):
                    candidate = TenderCandidate(
                        portal=portal.portal,
                        state=portal.state,
                        source_url=response_url,
                        organization=portal.portal,
                        title=label[:500],
                        category=infer_category(
                            label
                        ),
                        detail_url=href,
                        discovered_doc_urls=[
                            href
                        ],
                        keyword_hits=keyword_hits(
                            label
                        ),
                    )

                    key = candidate.stable_key()

                    if key not in seen_tender_keys:
                        seen_tender_keys.add(
                            key
                        )
                        results.append(
                            candidate
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

                    title = label

                    for heading_name in [
                        "h1",
                        "h2",
                        "h3",
                    ]:
                        heading = child.find(
                            heading_name
                        )

                        if heading:
                            heading_title = normalize_space(
                                heading.get_text(
                                    " ",
                                    strip=True,
                                )
                            )

                            if (
                                valid_tender_title(
                                    heading_title
                                )
                                and
                                keyword_hits(
                                    heading_title
                                )
                            ):
                                title = heading_title
                                break

                    if not keyword_hits(
                        title
                    ):
                        continue

                    child_text = normalize_space(
                        child.get_text(
                            " ",
                            strip=True,
                        )
                    )

                    refs = extract_ref_candidates(
                        child_text
                    )

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

                    candidate = TenderCandidate(
                        portal=portal.portal,
                        state=portal.state,
                        source_url=response_url,
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
                        keyword_hits=keyword_hits(
                            title
                        ),
                    )

                    key = candidate.stable_key()

                    if key not in seen_tender_keys:
                        seen_tender_keys.add(
                            key
                        )
                        results.append(
                            candidate
                        )
                        new_count += 1

                except requests.RequestException:
                    continue

            log.info(
                "%s | generic page %d | %d new matching tenders | total %d | %s",
                portal.portal,
                page_no,
                new_count,
                len(results),
                response_url,
            )

        log.info(
            "%s | generic crawl complete | pages=%d | tenders=%d",
            portal.portal,
            page_no,
            len(results),
        )

        output = self._dedupe(
            results
        )

        if (
            not output
            and
            GENERIC_JS_FALLBACK
        ):
            js_output = self.crawl_generic_playwright(
                portal
            )

            if js_output:
                log.info(
                    "%s | JS fallback recovered %d event tenders.",
                    portal.portal,
                    len(js_output),
                )

                return js_output

        return output

    def crawl_generic_playwright(
        self,
        portal,
    ):
        """
        Best-effort fallback for JavaScript-rendered tender listing pages.
        It does not bypass CAPTCHA/login.
        """

        if (
            sync_playwright is None
            or
            not GENERIC_JS_FALLBACK
        ):
            return []

        results = []
        seen = set()

        try:
            with sync_playwright() as playwright:
                browser = playwright.chromium.launch(
                    headless=True,
                    args=[
                        "--no-sandbox",
                        "--disable-dev-shm-usage",
                    ],
                )

                page = browser.new_page(
                    viewport={
                        "width": 1440,
                        "height": 1000,
                    }
                )

                page.goto(
                    portal.url,
                    timeout=45000,
                    wait_until="domcontentloaded",
                )

                page.wait_for_timeout(
                    2500
                )

                for page_no in range(
                    1,
                    min(
                        MAX_GENERIC_PAGES,
                        25,
                    ) + 1,
                ):
                    body = page.locator(
                        "body"
                    ).inner_text()

                    lower_body = body.lower()

                    if (
                        "captcha" in lower_body
                        and
                        len(body) < 100000
                    ):
                        break

                    anchors = page.locator(
                        "a[href]"
                    )

                    new_count = 0

                    for index in range(
                        min(
                            anchors.count(),
                            2500,
                        )
                    ):
                        try:
                            anchor = anchors.nth(
                                index
                            )

                            label = normalize_space(
                                anchor.inner_text()
                            )

                            if not valid_tender_title(
                                label
                            ):
                                continue

                            if not is_relevant_event_title(
                                label
                            ):
                                continue

                            href = anchor.get_attribute(
                                "href"
                            )

                            if not href:
                                continue

                            detail_url = safe_urljoin(
                                page.url,
                                href,
                            )

                            tender = TenderCandidate(
                                portal=portal.portal,
                                state=portal.state,
                                source_url=page.url,
                                organization=portal.portal,
                                title=label[:500],
                                category=infer_category(
                                    label
                                ),
                                detail_url=detail_url,
                                discovered_doc_urls=(
                                    [detail_url]
                                    if is_downloadable_url(
                                        detail_url
                                    )
                                    else []
                                ),
                                keyword_hits=keyword_hits(
                                    label
                                ),
                            )

                            key = tender.stable_key()

                            if key in seen:
                                continue

                            seen.add(
                                key
                            )

                            results.append(
                                tender
                            )

                            new_count += 1

                        except Exception:
                            continue

                    log.info(
                        "%s | JS fallback page %d | %d new event tenders | total %d",
                        portal.portal,
                        page_no,
                        new_count,
                        len(results),
                    )

                    next_locator = None

                    for selector in [
                        'a:has-text("Next")',
                        'button:has-text("Next")',
                        'a[aria-label*="Next" i]',
                        'button[aria-label*="Next" i]',
                        'a[rel="next"]',
                        'li.next a',
                    ]:
                        try:
                            locator = page.locator(
                                selector
                            ).first

                            if not (
                                locator.count()
                                and
                                locator.is_visible()
                            ):
                                continue

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
                                aria_disabled
                                or
                                "disabled" in classes
                            ):
                                continue

                            next_locator = locator
                            break

                        except Exception:
                            continue

                    if next_locator is None:
                        break

                    before = hashlib.sha256(
                        body.encode(
                            "utf-8",
                            errors="ignore",
                        )
                    ).hexdigest()

                    try:
                        next_locator.click(
                            timeout=8000
                        )

                        page.wait_for_timeout(
                            1800
                        )

                        after_body = page.locator(
                            "body"
                        ).inner_text()

                        after = hashlib.sha256(
                            after_body.encode(
                                "utf-8",
                                errors="ignore",
                            )
                        ).hexdigest()

                        if after == before:
                            break

                    except Exception:
                        break

                browser.close()

        except Exception as error:
            log.info(
                "%s | JS fallback failed: %s",
                portal.portal,
                str(error)[:220],
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
    Resilient GeM crawler.

    Primary:
      https://bidplus.gem.gov.in/all-bids

    Fallback:
      https://bidplus-global.gem.gov.in/

    If the GitHub runner cannot reach the primary endpoint, the crawler
    automatically falls back to the GTE surface.
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
        "Event Agency",
        "Exhibition",
        "Exhibition Stall",
        "Conference Management",
        "Conclave",
        "Summit",
        "Mela",
        "Festival",
        "Creative Agency",
        "Advertising Agency",
        "Brand Activation",
        "Audio Visual",
        "Sound and Light",
        "Event Empanelment",
        "Outreach Campaign",
        "Experiential Marketing",
        "Foundation Day",
        "Roadshow",
        "Pavilion",
        "Stall Fabrication",
    ]

    def crawl_surface(
        target_url,
        surface_name,
    ):
        collected = []

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
                ),
                viewport={
                    "width": 1440,
                    "height": 1000,
                },
            )

            page = context.new_page()

            page.goto(
                target_url,
                timeout=60000,
                wait_until="domcontentloaded",
            )
            page.wait_for_timeout(4000)

            search_ui_found = False

            for term in search_terms:
                search_input = None

                for selector in [
                    'input[placeholder*="Enter Keyword" i]',
                    'input[placeholder*="Enter Keywords" i]',
                    'input[placeholder*="Keyword" i]',
                    'input[placeholder*="Search" i]',
                    'input[type="search"]',
                    'input#search_by',
                ]:
                    try:
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
                    except Exception:
                        continue

                if search_input is None:
                    log.warning(
                        "GeM %s search input not found for: %s",
                        surface_name,
                        term,
                    )
                    continue

                try:
                    search_input.fill("")
                    search_input.fill(term)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(3000)

                    # Some GeM versions require a Search button click.
                    for selector in [
                        'button:has-text("Search")',
                        'input[type="submit"][value*="Search" i]',
                        'button[aria-label*="Search" i]',
                    ]:
                        try:
                            button = page.locator(
                                selector
                            ).first

                            if (
                                button.count()
                                and
                                button.is_visible()
                            ):
                                button.click(
                                    timeout=5000
                                )
                                page.wait_for_timeout(
                                    2000
                                )
                                break
                        except Exception:
                            continue

                    seen_page_signatures = set()
                    seen_bid_ids = set()

                    for page_number in range(
                        1,
                        MAX_GEM_PAGES_PER_SEARCH + 1,
                    ):
                        body_text = page.locator(
                            "body"
                        ).inner_text()
                        html = page.content()

                        bid_numbers = extract_gem_bid_numbers(
                            body_text,
                            html,
                        )

                        anchors = page.locator(
                            "a[href]"
                        )
                        hrefs = []

                        for index in range(
                            min(
                                anchors.count(),
                                2500,
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

                        bid_numbers = extract_gem_bid_numbers(
                            "\n".join(bid_numbers),
                            "\n".join(
                                label
                                for label, _ in hrefs
                                if label
                            ),
                        )

                        signature = hashlib.sha256(
                            (
                                "|".join(bid_numbers)
                                + "|"
                                + page.url
                            ).encode(
                                "utf-8"
                            )
                        ).hexdigest()

                        if signature in seen_page_signatures:
                            break

                        seen_page_signatures.add(
                            signature
                        )

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
                            visible_title = ""

                            for label, href in hrefs:
                                combined = (
                                    f"{label} {href}"
                                ).lower()

                                if (
                                    bid_no.lower() in combined
                                    or
                                    numeric_bid in combined
                                ):
                                    matching_urls.append(
                                        href
                                    )

                                    if (
                                        label
                                        and
                                        len(label)
                                        >
                                        len(visible_title)
                                        and
                                        label.lower()
                                        !=
                                        bid_no.lower()
                                    ):
                                        visible_title = label

                            matching_urls = list(
                                dict.fromkeys(
                                    matching_urls
                                )
                            )

                            title = (
                                visible_title
                                if valid_tender_title(
                                    visible_title
                                )
                                else
                                f"{term} | {bid_no}"
                            )

                            real_url = (
                                matching_urls[0]
                                if matching_urls
                                else page.url
                            )

                            collected.append(
                                TenderCandidate(
                                    portal=(
                                        "Government "
                                        "e-Marketplace (GeM)"
                                    ),
                                    state="Pan India",
                                    source_url=target_url,
                                    tender_id=bid_no,
                                    organization="NOT VERIFIED",
                                    title=title[:500],
                                    category=infer_category(
                                        title
                                    ),
                                    detail_url=real_url,
                                    discovered_doc_urls=(
                                        matching_urls[:10]
                                    ),
                                    keyword_hits=keyword_hits(
                                        title
                                        + " "
                                        + term
                                    ),
                                )
                            )
                            new_count += 1

                        log.info(
                            "GeM %s | %s | page %d | %d new bids | total=%d",
                            surface_name,
                            term,
                            page_number,
                            new_count,
                            len(seen_bid_ids),
                        )

                        next_locator = None

                        for selector in [
                            'a:has-text("Next")',
                            'button:has-text("Next")',
                            'li.next a',
                            'a[aria-label*="Next" i]',
                            'button[aria-label*="Next" i]',
                            'a[rel="next"]',
                            'button[title*="Next" i]',
                            'a[title*="Next" i]',
                        ]:
                            try:
                                locator = page.locator(
                                    selector
                                ).first

                                if not (
                                    locator.count()
                                    and
                                    locator.is_visible()
                                ):
                                    continue

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
                                    or
                                    aria_disabled
                                    or
                                    "disabled" in classes
                                ):
                                    continue

                                next_locator = locator
                                break
                            except Exception:
                                continue

                        if next_locator is None:
                            break

                        before_signature = signature

                        try:
                            next_locator.click(
                                timeout=10000
                            )
                            page.wait_for_timeout(
                                2500
                            )

                            after_body = page.locator(
                                "body"
                            ).inner_text()
                            after_html = page.content()

                            after_bids = extract_gem_bid_numbers(
                                after_body,
                                after_html,
                            )

                            after_signature = hashlib.sha256(
                                (
                                    "|".join(after_bids)
                                    + "|"
                                    + page.url
                                ).encode(
                                    "utf-8"
                                )
                            ).hexdigest()

                            if (
                                after_signature
                                ==
                                before_signature
                            ):
                                break
                        except Exception:
                            break

                except Exception as error:
                    log.warning(
                        "GeM %s search failed for [%s]: %s",
                        surface_name,
                        term,
                        str(error)[:300],
                    )

            browser.close()

            if not search_ui_found:
                raise RuntimeError(
                    f"GeM {surface_name} search UI was not found."
                )

        unique = {}

        for tender in collected:
            key = (
                tender.tender_id
                or
                tender.stable_key()
            )

            existing = unique.get(key)

            if existing is None:
                unique[key] = tender
                continue

            existing_score = (
                len(existing.title or "")
                +
                (
                    100
                    if existing.discovered_doc_urls
                    else 0
                )
            )

            new_score = (
                len(tender.title or "")
                +
                (
                    100
                    if tender.discovered_doc_urls
                    else 0
                )
            )

            if new_score > existing_score:
                unique[key] = tender

        return list(
            unique.values()
        )

    primary_error = ""

    try:
        primary_results = crawl_surface(
            GEM_LISTING_URL,
            "PRIMARY",
        )

        if primary_results:
            return (
                primary_results,
                CrawlHealth(
                    "Government e-Marketplace (GeM)",
                    GEM_LISTING_URL,
                    "SUCCESS",
                    len(primary_results),
                    "Primary GeM all-bids surface used.",
                ),
            )

        log.warning(
            "Primary GeM surface returned no matching bids. Trying GTE fallback."
        )

    except Exception as error:
        primary_error = str(error)[:250]

        log.warning(
            "Primary GeM surface failed: %s",
            primary_error,
        )

    try:
        fallback_results = crawl_surface(
            GEM_GTE_URL,
            "GTE_FALLBACK",
        )

        if fallback_results:
            return (
                fallback_results,
                CrawlHealth(
                    "Government e-Marketplace (GeM)",
                    GEM_GTE_URL,
                    "FALLBACK_SUCCESS",
                    len(fallback_results),
                    (
                        "Primary GeM unavailable/empty; "
                        "GTE fallback used."
                    ),
                ),
            )

        return (
            [],
            CrawlHealth(
                "Government e-Marketplace (GeM)",
                GEM_GTE_URL,
                "NO_RESULTS",
                0,
                (
                    "Primary GeM unavailable/empty and "
                    "GTE fallback returned no matching bids."
                ),
            ),
        )

    except Exception as fallback_error:
        return (
            [],
            CrawlHealth(
                "Government e-Marketplace (GeM)",
                GEM_LISTING_URL,
                "PARSER_FAILED",
                0,
                (
                    "Primary error: "
                    f"{primary_error or 'no matching results'}; "
                    "Fallback error: "
                    f"{str(fallback_error)[:180]}"
                ),
            ),
        )


# =============================================================================
# 12. DOCUMENT MANAGER
# =============================================================================

class DocumentManager:
    def __init__(self):
        self.http = session_with_headers()

    def _document_relevance_score(
        self,
        candidate,
        url,
        label="",
    ):
        combined = normalize_space(
            f"{label} {url}"
        ).lower()

        score = 0

        tender_id = normalize_space(
            candidate.tender_id or ""
        ).lower()

        numeric_parts = re.findall(
            r"\d{5,}",
            tender_id
        )

        if tender_id and tender_id in combined:
            score += 6

        if any(
            part in combined
            for part in numeric_parts
        ):
            score += 4

        if "showbiddocument" in combined:
            score += 5

        if any(
            term in combined
            for term in [
                "tender document",
                "tender_document",
                "tender-document",
                "rfp",
                "request for proposal",
                "nit",
                "notice inviting tender",
                "eoi",
                "expression of interest",
                "rfq",
                "boq",
                "atc",
                "corrigendum",
                "addendum",
                "annexure",
                "eligibility",
                "bid document",
                "bid_document",
            ]
        ):
            score += 3

        if any(
            bad in combined
            for bad in [
                "annual-report",
                "annual_report",
                "investor",
                "brochure",
                "product",
                "contact",
                "vendor-registration",
                "vendor_registration",
                "dividend",
                "postal-ballot",
                "postal_ballot",
                "tds-guideline",
                "material-events",
                "gallery",
                "news",
            ]
        ):
            score -= 6

        return score

    def _filter_document_urls(
        self,
        candidate,
        items,
    ):
        """
        items can be URLs or (label, URL) tuples.
        Keep only documents plausibly tied to this tender.
        """
        scored = []

        for item in items:
            if isinstance(
                item,
                tuple,
            ):
                label, url = item
            else:
                label, url = "", item

            if not url:
                continue

            score = self._document_relevance_score(
                candidate,
                url,
                label,
            )

            # Direct candidate detail document gets a modest boost.
            if (
                candidate.detail_url
                and
                url == candidate.detail_url
                and
                is_downloadable_url(
                    url
                )
            ):
                score += 2

            if score >= 2:
                scored.append(
                    (
                        score,
                        url,
                    )
                )

        scored.sort(
            key=lambda item: item[0],
            reverse=True,
        )

        output = []

        for _, url in scored:
            if url not in output:
                output.append(
                    url
                )

        return output[:8]

    def discover_from_detail_page(
        self,
        candidate,
    ):
        discovered_items = [
            (
                "",
                url,
            )
            for url in candidate.discovered_doc_urls
            if url
        ]

        if not candidate.detail_url:
            return self._filter_document_urls(
                candidate,
                discovered_items,
            )

        if is_downloadable_url(
            candidate.detail_url
        ):
            discovered_items.append(
                (
                    candidate.title,
                    candidate.detail_url,
                )
            )

            return self._filter_document_urls(
                candidate,
                discovered_items,
            )

        try:
            response = self.http.get(
                candidate.detail_url,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )

            if response.status_code != 200:
                return self._filter_document_urls(
                    candidate,
                    discovered_items,
                )

            content_type = (
                response.headers.get(
                    "content-type"
                )
                or ""
            ).lower()

            if "application/pdf" in content_type:
                discovered_items.append(
                    (
                        candidate.title,
                        response.url,
                    )
                )

                return self._filter_document_urls(
                    candidate,
                    discovered_items,
                )

            soup = BeautifulSoup(
                response.text,
                "html.parser",
            )

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
                )

                if not url:
                    continue

                if (
                    is_downloadable_url(
                        url
                    )
                    or
                    any(
                        term in (
                            label
                            + " "
                            + url
                        ).lower()
                        for term in [
                            "tender",
                            "rfp",
                            "nit",
                            "eoi",
                            "rfq",
                            "boq",
                            "atc",
                            "corrigendum",
                            "addendum",
                            "annexure",
                            "bid document",
                            "eligibility",
                        ]
                    )
                ):
                    discovered_items.append(
                        (
                            label,
                            url,
                        )
                    )

        except requests.RequestException:
            pass

        return self._filter_document_urls(
            candidate,
            discovered_items,
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
    Remove stale false positives and malformed rows from older crawler logic.
    Keeps only rows that still look like a specific event-related procurement.
    """

    if not CLEAN_EXISTING_ACTIVE_TENDERS:
        return 0

    values = sheet.get_all_values()

    if len(values) <= 1:
        return 0

    headers = values[0]

    def idx(name):
        try:
            return headers.index(
                name
            )
        except ValueError:
            return None

    title_index = idx(
        "Tender Title & Scope"
    )
    tender_id_index = idx(
        "Tender ID / Ref No"
    )
    portal_index = idx(
        "Portal Name"
    )
    state_index = idx(
        "State"
    )
    org_index = idx(
        "Organization / Dept"
    )
    detail_index = idx(
        "Portal / Detail Link"
    )
    docs_index = idx(
        "RFP / Tender Doc URLs"
    )

    if title_index is None:
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

        detail_url = (
            padded[detail_index]
            if detail_index is not None
            else ""
        )

        candidate = TenderCandidate(
            portal=(
                padded[portal_index]
                if portal_index is not None
                else ""
            ),
            state=(
                padded[state_index]
                if state_index is not None
                else "Pan India"
            ),
            source_url=detail_url,
            tender_id=(
                padded[tender_id_index]
                if tender_id_index is not None
                else ""
            ),
            organization=(
                padded[org_index]
                if org_index is not None
                else ""
            ),
            title=title,
            detail_url=detail_url,
            discovered_doc_urls=(
                [
                    item
                    for item in (
                        padded[docs_index].splitlines()
                        if docs_index is not None
                        else []
                    )
                    if item
                ]
            ),
        )

        service_ok, _service_reason = is_event_service_candidate(
            candidate
        )

        keep = (
            is_relevant_event_title(
                title
            )
            and
            service_ok
            and
            is_procurement_candidate(
                candidate
            )
        )

        # Remove obviously corrupted historic rows such as "35" in URL/status cells.
        if keep:
            invalid_scalar = False

            for field_name in [
                "RFP / Tender Doc URLs",
                "Portal / Detail Link",
                "AI Model Used",
                "Extraction Status",
            ]:
                field_index = idx(
                    field_name
                )

                if field_index is None:
                    continue

                value = normalize_space(
                    padded[
                        field_index
                    ]
                )

                if value.isdigit():
                    invalid_scalar = True
                    break

            if invalid_scalar:
                keep = False

        if keep:
            kept_rows.append(
                padded
            )
        else:
            removed += 1

            log.info(
                "Removing stale/invalid tracker row: %s",
                title[:160],
            )

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
        "Cleaned Active_Tenders: removed %d stale/invalid rows, kept %d rows.",
        removed,
        len(kept_rows),
    )

    return removed



VALID_EXTRACTION_STATUSES = {
    "OK",
    "SUCCESS",
    "NO_DOCUMENT_TEXT",
    "OPENAI_API_KEY_MISSING",
    "OPENAI_DISABLED_FOR_RUN",
    "OPENAI_QUOTA_EXHAUSTED",
    "OPENAI_AUTH_ERROR",
    "OPENAI_MODEL_ERROR",
    "OPENAI_ERROR",
    "PARSE_FAILED",
    "UNKNOWN",
}


def sanitize_active_row(row):
    if len(row) != len(
        ACTIVE_HEADERS
    ):
        raise ValueError(
            f"Active_Tenders row has {len(row)} columns; "
            f"expected {len(ACTIVE_HEADERS)}."
        )

    clean = []

    for value in row:
        if value is None:
            clean.append(
                ""
            )
        elif isinstance(
            value,
            (
                int,
                float,
            ),
        ):
            clean.append(
                value
            )
        else:
            clean.append(
                str(value).strip()
            )

    mapping = dict(
        zip(
            ACTIVE_HEADERS,
            clean,
        )
    )

    title = normalize_space(
        mapping[
            "Tender Title & Scope"
        ]
    )

    if not title:
        raise ValueError(
            "Tender title is blank."
        )

    detail_link = normalize_space(
        mapping[
            "Portal / Detail Link"
        ]
    )

    if (
        detail_link
        and
        detail_link.isdigit()
    ):
        raise ValueError(
            "Portal/detail link is numeric corruption."
        )

    doc_urls = normalize_space(
        mapping[
            "RFP / Tender Doc URLs"
        ]
    )

    if (
        doc_urls
        and
        doc_urls.isdigit()
    ):
        raise ValueError(
            "Tender document URLs field is numeric corruption."
        )

    ai_model = normalize_space(
        mapping[
            "AI Model Used"
        ]
    )

    if ai_model.isdigit():
        clean[
            ACTIVE_HEADERS.index(
                "AI Model Used"
            )
        ] = ""


    evidence_quotes_value = normalize_space(
        mapping[
            "Evidence Quotes"
        ]
    )

    if re.fullmatch(
        r"\d+(?:\.0+)?",
        evidence_quotes_value,
    ):
        clean[
            ACTIVE_HEADERS.index(
                "Evidence Quotes"
            )
        ] = ""

    extraction_status = normalize_space(
        mapping[
            "Extraction Status"
        ]
    )

    if (
        extraction_status
        and
        extraction_status not in VALID_EXTRACTION_STATUSES
    ):
        clean[
            ACTIVE_HEADERS.index(
                "Extraction Status"
            )
        ] = "UNKNOWN"

    return clean


def write_active_rows_exact(
    sheet,
    rows,
):
    if not rows:
        return 0

    validated = []

    for row in rows:
        try:
            validated.append(
                sanitize_active_row(
                    row
                )
            )
        except Exception as error:
            log.error(
                "Rejecting malformed Active_Tenders row: %s",
                error,
            )

    if not validated:
        return 0

    existing = sheet.get_all_values()

    start_row = max(
        2,
        len(existing) + 1,
    )

    end_row = (
        start_row
        + len(validated)
        - 1
    )

    sheet.update(
        values=validated,
        range_name=(
            f"A{start_row}:Z{end_row}"
        ),
        value_input_option="RAW",
    )

    return len(
        validated
    )



def sync_portal_directory(
    sheet,
    portals,
    health,
):
    rows = [
        PORTAL_HEADERS
    ]

    gem_health = health.get(
        "Government e-Marketplace (GeM)"
    )

    for portal in portals:
        status = health.get(
            portal.portal
        )

        # GeM master sheet may contain separate standard and BidPlus/GTE rows.
        # Reuse the actual GeM run status rather than leaving stale NOT_RUN junk.
        if (
            status is None
            and
            gem_health is not None
            and
            (
                "gem.gov.in"
                in portal.url.lower()
                or
                "bidplus"
                in portal.url.lower()
            )
        ):
            status = gem_health

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

        checked_at = (
            str(
                status.checked_at
                or ""
            ).strip()
            if status
            else ""
        )

        if (
            checked_at
            and
            not re.match(
                r"^\d{4}-\d{2}-\d{2} ",
                checked_at,
            )
        ):
            checked_at = ""

        crawler_status = (
            str(
                status.status
                or "NOT_RUN"
            ).strip()
            if status
            else "NOT_RUN"
        )

        if crawler_status.isdigit():
            crawler_status = "UNKNOWN"

        message = (
            str(
                status.message
                or ""
            ).strip()
            if status
            else ""
        )

        if re.fullmatch(
            r"\d+(?:\.0+)?",
            message,
        ):
            message = ""

        try:
            discovered_count = (
                int(
                    status.discovered_count
                    or 0
                )
                if status
                else 0
            )
        except Exception:
            discovered_count = 0

        if (
            crawler_status == "SUCCESS"
            and
            discovered_count == 0
            and
            re.fullmatch(
                r"\d+(?:\.0+)?",
                message or "",
            )
        ):
            message = ""

        rows.append([
            portal.portal,
            portal.category,
            portal.state,
            portal.url,
            portal.active,
            checked_at,
            discovered_count,
            crawler_status,
            message,
            portal.priority,
        ])

    sheet.clear()

    sheet.update(
        values=rows,
        range_name=f"A1:J{len(rows)}",
        value_input_option="RAW",
    )

    log.info(
        "Portal_Directory rewritten with %d portal rows.",
        len(rows) - 1,
    )


# =============================================================================
# 17. CANDIDATE FILTER
# =============================================================================

def filter_event_candidates(
    candidates,
):
    filtered = []
    rejected = []

    # Cross-source dedupe: prefer one specific opportunity instead of the same
    # tender repeated from official + aggregator/category pages.
    seen_business_keys = set()

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

        service_ok, service_reason = is_event_service_candidate(
            candidate
        )

        if not service_ok:
            rejected.append(
                candidate
            )

            log.info(
                "Skipping non-service/random tender: %s | %s",
                candidate.title[:150],
                service_reason,
            )
            continue

        procurement_score, procurement_reasons = (
            procurement_intent_score(
                candidate
            )
        )

        if procurement_score < 3:
            rejected.append(
                candidate
            )

            log.info(
                "Skipping non-procurement content: %s | score=%d | %s",
                candidate.title[:150],
                procurement_score,
                ", ".join(
                    procurement_reasons
                )[:250],
            )
            continue

        # Reclassify after the strict service filter.
        candidate.category = infer_category(
            candidate.title
        )

        ref = normalize_space(
            candidate.tender_id or ""
        ).lower()

        if (
            ref
            and
            ref not in {
                "not verified",
                "not stated",
            }
        ):
            business_key = (
                "REF:"
                + re.sub(
                    r"\s+",
                    "",
                    ref,
                )
            )
        else:
            normalized_title = re.sub(
                r"[^a-z0-9]+",
                " ",
                normalize_space(
                    candidate.title
                ).lower(),
            ).strip()

            business_key = (
                "TITLE:"
                + normalized_title[:220]
            )

        if business_key in seen_business_keys:
            log.info(
                "Skipping duplicate tender across sources: %s",
                candidate.title[:150],
            )
            continue

        seen_business_keys.add(
            business_key
        )

        filtered.append(
            candidate
        )

    log.info(
        "Strict SERVICE filter: %d kept / %d rejected",
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

    evidence = merge_deterministic_metadata(
        candidate,
        documents,
        evidence,
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

def validate_deployed_build():
    """
    Fail immediately if GitHub is running an older/incomplete file.
    """
    required_functions = [
        "procurement_intent_score",
        "extract_deterministic_metadata",
        "merge_deterministic_metadata",
        "sanitize_active_row",
        "write_active_rows_exact",
        "clean_active_tenders_sheet",
        "sync_portal_directory",
        "crawl_gem",
        "is_event_service_candidate",
        "is_expired_candidate",
    ]

    missing = [
        name
        for name in required_functions
        if name not in globals()
    ]

    log.info(
        "Crawler code version: %s",
        CODE_VERSION,
    )
    log.info(
        "Strict procurement gate enabled: %s",
        "procurement_intent_score" in globals(),
    )
    log.info(
        "Active tender cleanup enabled: %s",
        CLEAN_EXISTING_ACTIVE_TENDERS,
    )
    log.info(
        "Generic JS fallback enabled: %s",
        GENERIC_JS_FALLBACK,
    )
    log.info(
        "GeM primary URL: %s",
        GEM_LISTING_URL,
    )
    log.info(
        "GeM fallback URL: %s",
        GEM_GTE_URL,
    )

    if missing:
        raise RuntimeError(
            "Outdated/incomplete tender_agent.py deployed. "
            f"Missing functions: {missing}"
        )



def run_pipeline():
    validate_deployed_build()

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

        # Last safety layer before downloading documents or writing to Sheets.
        if not is_relevant_event_title(
            candidate.title
        ):
            log.warning(
                "FINAL WRITE BLOCKED - non-event title: %s",
                candidate.title[:200],
            )
            continue

        service_ok, service_reason = is_event_service_candidate(
            candidate
        )

        if not service_ok:
            log.warning(
                "FINAL WRITE BLOCKED - non-service/random scope: %s | %s",
                candidate.title[:200],
                service_reason,
            )
            continue

        candidate.category = infer_category(
            candidate.title
        )

        procurement_score, procurement_reasons = procurement_intent_score(
            candidate
        )

        if procurement_score < 3:
            log.warning(
                "FINAL WRITE BLOCKED - no procurement intent: %s | score=%d | %s",
                candidate.title[:200],
                procurement_score,
                ", ".join(procurement_reasons)[:250],
            )
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
        written = write_active_rows_exact(
            active_sheet,
            new_rows,
        )

        log.info(
            "Successfully added %d validated new tender records.",
            written,
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
