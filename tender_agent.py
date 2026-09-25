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
from google import genai
from google.genai import types
from pydantic import BaseModel
from pypdf import PdfReader
from docx import Document as DocxDocument

try:
    from playwright.sync_api import sync_playwright
except Exception:
    sync_playwright = None


# =============================================================================
# 1. CONFIG
# =============================================================================

APP_NAME = "Pan-India Tender Intelligence Agent"

IST = dt.timezone(
    dt.timedelta(hours=5, minutes=30)
)

SERVICE_ACCOUNT_FILE = os.getenv(
    "SERVICE_ACCOUNT_FILE",
    "service_account.json"
)

SPREADSHEET_ID = os.getenv(
    "SPREADSHEET_ID",
    ""
)

if not SPREADSHEET_ID:
    raise RuntimeError(
        "SPREADSHEET_ID environment variable is missing."
    )

EXCEL_MASTER_FILE = os.getenv(
    "EXCEL_MASTER_FILE",
    "Pan_India_Tender_URL_Master.xlsx"
)

GEMINI_API_KEY = os.getenv(
    "GEMINI_API_KEY",
    ""
)

GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-2.5-flash-lite"
)

GEM_LISTING_URL = os.getenv(
    "GEM_LISTING_URL",
    "https://bidplus-global.gem.gov.in/"
)

DOWNLOAD_DIR = Path(
    os.getenv(
        "DOWNLOAD_DIR",
        "tender_downloads"
    )
)

DOWNLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True
)

REQUEST_TIMEOUT = int(
    os.getenv(
        "REQUEST_TIMEOUT",
        "18"
    )
)

CRAWL_WORKERS = int(
    os.getenv(
        "CRAWL_WORKERS",
        "12"
    )
)

MAX_DOC_BYTES = int(
    os.getenv(
        "MAX_DOC_BYTES",
        str(30 * 1024 * 1024)
    )
)

MAX_DOC_TEXT_CHARS = int(
    os.getenv(
        "MAX_DOC_TEXT_CHARS",
        "120000"
    )
)

MAX_GENERIC_LINKS = int(
    os.getenv(
        "MAX_GENERIC_LINKS",
        "25"
    )
)

logging.basicConfig(
    level=os.getenv(
        "LOG_LEVEL",
        "INFO"
    ).upper(),
    format="%(asctime)s | %(levelname)s | %(message)s",
)

log = logging.getLogger(
    "tender-agent"
)


# =============================================================================
# 2. COMPANY PROFILE
# =============================================================================

COMPANY_PROFILE = {
    "agency_name": "Soul Events and Consultancy",
    "constitution": "Proprietorship",

    "avg_3yr_turnover_inr": 38046000,
    "fy_2025_26_turnover_inr": 50411575,
    "audited_net_worth_inr": 4680004,

    "max_single_past_work_order_inr": 7898550,
    "max_two_works_threshold_inr": 7223000,
    "max_three_works_threshold_inr": 4628443,
}


# =============================================================================
# 3. SEARCH KEYWORDS
# =============================================================================

EVENT_KEYWORDS = [
    "Event Management",
    "Event Agency",
    "Event Management Agency",

    "Exhibition",
    "Exhibition Stall",
    "Exhibition Pavilion",

    "Design and Fabrication",
    "Fabrication",

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

    "Publicity",
    "Outreach Campaign",
    "IEC",

    "Media Campaign",
    "Advertising Agency",
    "Creative Agency",

    "AV Production",
    "Audio Visual",
    "LED",
    "Sound and Light",

    "Stage",
    "Tentage",
    "Decoration",
    "Venue Management",

    "Hospitality",
    "Manpower",

    "Printing",
    "Branding",
    "Signage",

    "Digital Marketing",
    "Social Media",

    "PR Agency",
    "Communication Agency",

    "Empanelment of Event Agency",
    "Empanelment of Advertising Agency",

    "EOI Event",
    "RFP Event",

    "Tourism Event",
    "Sports Event",

    "Government Function",
    "Launch Event",
    "Inauguration",
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


# =============================================================================
# 4. DATA MODELS
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
        default_factory=lambda:
        dt.datetime.now(
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
# 5. GEMINI STRUCTURED RESPONSE
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

    evidence_quotes: list[str] = []

    missing_or_unclear_fields: list[str] = []


# =============================================================================
# 6. HELPERS
# =============================================================================

def normalize_space(value):

    return re.sub(
        r"\s+",
        " ",
        value or ""
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

    t = (
        title or ""
    ).lower()

    if (
        "exhibition" in t
        or "stall" in t
        or "pavilion" in t
    ):
        return "Exhibition / Stall"

    if "empanel" in t:
        return "Empanelment"

    if any(
        x in t
        for x in [
            "conference",
            "summit",
            "conclave",
            "seminar"
        ]
    ):
        return "Conference / Conclave"

    if any(
        x in t
        for x in [
            "advertising",
            "media campaign",
            "creative agency"
        ]
    ):
        return "Advertising / Creative / Outreach"

    if any(
        x in t
        for x in [
            "festival",
            "mela",
            "cultural"
        ]
    ):
        return "Festival / Mela / Cultural"

    return "Event / Experiential"


def safe_urljoin(
    base,
    href
):

    return urllib.parse.urljoin(
        base,
        href or ""
    )


def is_downloadable_url(url):

    path = urllib.parse.urlsplit(
        url
    ).path.lower()

    return path.endswith(
        DOWNLOADABLE_EXTENSIONS
    )


def extract_ref_candidates(text):

    patterns = [
        r"GEM/\d{4}/B/\d+",

        r"\b\d{4}_[A-Za-z0-9_-]+_\d+_\d+\b",

        r"\b(?:RFP|EOI|NIT|TENDER|BID)"
        r"[\s:/-]*"
        r"[A-Za-z0-9._/-]{3,}\b",
    ]

    output = []

    for pattern in patterns:

        output.extend(
            re.findall(
                pattern,
                text or "",
                flags=re.I
            )
        )

    return list(
        dict.fromkeys(
            normalize_space(x)
            for x in output
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

        "Accept-Language":
            "en-IN,en;q=0.9",
    })

    return session


# =============================================================================
# 7. MASTER PORTAL LOADER
# =============================================================================

def load_master_portals():

    if not os.path.exists(
        EXCEL_MASTER_FILE
    ):

        raise FileNotFoundError(
            f"Master Excel not found: "
            f"{EXCEL_MASTER_FILE}"
        )

    df = pd.read_excel(
        EXCEL_MASTER_FILE,
        sheet_name="Tender URL Master",
        skiprows=2
    )

    portals = []

    for _, row in df.iterrows():

        if str(
            row.get(
                "Active",
                ""
            )
        ).strip().lower() != "yes":

            continue

        url = str(
            row.get(
                "URL",
                ""
            )
        ).strip()

        if not url.startswith(
            (
                "http://",
                "https://"
            )
        ):
            continue

        portals.append(
            Portal(
                portal=str(
                    row.get(
                        "Portal / Organisation",
                        ""
                    )
                ).strip(),

                category=str(
                    row.get(
                        "Category",
                        ""
                    )
                ).strip(),

                state=str(
                    row.get(
                        "State/Region",
                        ""
                    )
                ).strip(),

                url=url,

                priority=str(
                    row.get(
                        "Priority",
                        "P2"
                    )
                ).strip().upper()
                or "P2",

                source_type=str(
                    row.get(
                        "Source Type",
                        ""
                    )
                ).strip(),
            )
        )

    return portals


# =============================================================================
# 8. PORTAL CRAWLER
# =============================================================================

class TenderCrawler:

    def __init__(self):

        self.http = (
            session_with_headers()
        )


    def crawl(
        self,
        portal
    ):

        try:

            if (
                "gem.gov.in"
                in portal.url.lower()
            ):

                return [], CrawlHealth(
                    portal.portal,
                    portal.url,
                    "SKIPPED_DEDICATED_ADAPTER"
                )


            if self._looks_gepnic(
                portal
            ):

                items = (
                    self.crawl_gepnic(
                        portal
                    )
                )

            else:

                items = (
                    self.crawl_generic(
                        portal
                    )
                )


            return (
                items,
                CrawlHealth(
                    portal.portal,
                    portal.url,
                    "SUCCESS",
                    len(items)
                )
            )


        except requests.exceptions.Timeout:

            return [], CrawlHealth(
                portal.portal,
                portal.url,
                "TIMEOUT",
                0,
                "HTTP request timed out"
            )


        except requests.exceptions.HTTPError as error:

            code = getattr(
                error.response,
                "status_code",
                None
            )

            status = (
                "LOGIN_OR_BLOCKED"
                if code in (
                    401,
                    403
                )
                else
                "HTTP_ERROR"
            )

            return [], CrawlHealth(
                portal.portal,
                portal.url,
                status,
                0,
                str(error)[:250]
            )


        except RuntimeError as error:

            message = str(
                error
            )

            status = (
                "CAPTCHA_OR_LOGIN"
                if (
                    "CAPTCHA" in message
                    or "LOGIN" in message
                )
                else
                "PARSER_FAILED"
            )

            return [], CrawlHealth(
                portal.portal,
                portal.url,
                status,
                0,
                message[:250]
            )


        except Exception as error:

            return [], CrawlHealth(
                portal.portal,
                portal.url,
                "PARSER_FAILED",
                0,
                str(error)[:250]
            )


    def _looks_gepnic(
        self,
        portal
    ):

        url = (
            portal.url.lower()
        )

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


    # =========================================================================
    # GePNIC
    # =========================================================================

    def crawl_gepnic(
        self,
        portal
    ):

        base = (
            portal.url.rstrip("/")
        )

        candidate_urls = [
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
        ]

        html = ""
        final_url = ""

        for url in dict.fromkeys(
            candidate_urls
        ):

            try:

                response = (
                    self.http.get(
                        url,
                        timeout=REQUEST_TIMEOUT,
                        allow_redirects=True
                    )
                )

                if (
                    response.status_code
                    == 200
                    and len(
                        response.text
                    ) > 500
                ):

                    html = (
                        response.text
                    )

                    final_url = (
                        response.url
                    )

                    break

            except requests.RequestException:

                continue


        if not html:

            raise RuntimeError(
                "No usable GePNIC page returned"
            )


        if (
            "captcha"
            in html.lower()
            and
            "latest active tenders"
            not in html.lower()
        ):

            raise RuntimeError(
                "CAPTCHA_OR_LOGIN_REQUIRED"
            )


        soup = BeautifulSoup(
            html,
            "html.parser"
        )


        results = []


        for row in soup.find_all(
            "tr"
        ):

            text = normalize_space(
                row.get_text(
                    " ",
                    strip=True
                )
            )

            hits = keyword_hits(
                text
            )

            if not hits:
                continue


            cells = [
                normalize_space(
                    td.get_text(
                        " ",
                        strip=True
                    )
                )
                for td
                in row.find_all(
                    "td"
                )
            ]


            if len(
                cells
            ) < 2:
                continue


            refs = (
                extract_ref_candidates(
                    text
                )
            )


            tender_id = (
                refs[0]
                if refs
                else ""
            )


            title = max(
                cells,
                key=len
            )


            detail_url = (
                final_url
            )


            documents = []


            for anchor in row.find_all(
                "a",
                href=True
            ):

                href = safe_urljoin(
                    final_url,
                    anchor.get(
                        "href"
                    )
                )

                if (
                    not href
                    or href.lower().startswith(
                        "javascript:"
                    )
                ):
                    continue


                if is_downloadable_url(
                    href
                ):

                    documents.append(
                        href
                    )


                elif any(
                    keyword in (
                        anchor.get_text(
                            " ",
                            strip=True
                        )
                        + " "
                        + href
                    ).lower()

                    for keyword
                    in TENDERISH_TERMS
                ):

                    detail_url = (
                        href
                    )


            results.append(
                TenderCandidate(
                    portal=portal.portal,
                    state=portal.state,
                    source_url=final_url,

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


        return self._dedupe(
            results
        )


    # =========================================================================
    # GENERIC WEBSITE
    # =========================================================================

    def crawl_generic(
        self,
        portal
    ):

        response = (
            self.http.get(
                portal.url,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True
            )
        )

        response.raise_for_status()


        if (
            "captcha"
            in response.text.lower()
            and len(
                response.text
            ) < 50000
        ):

            raise RuntimeError(
                "CAPTCHA_OR_LOGIN_REQUIRED"
            )


        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )


        links = []


        for anchor in soup.find_all(
            "a",
            href=True
        ):

            label = normalize_space(
                anchor.get_text(
                    " ",
                    strip=True
                )
            )


            href = safe_urljoin(
                response.url,
                anchor[
                    "href"
                ]
            )


            combined = (
                label
                + " "
                + href
            ).lower()


            if any(
                term
                in combined

                for term
                in TENDERISH_TERMS
            ):

                links.append(
                    (
                        label,
                        href
                    )
                )


        results = []


        for (
            label,
            href
        ) in list(
            dict.fromkeys(
                links
            )
        )[:MAX_GENERIC_LINKS]:


            if is_downloadable_url(
                href
            ):

                if keyword_hits(
                    label
                ):

                    results.append(
                        TenderCandidate(
                            portal=portal.portal,
                            state=portal.state,

                            source_url=response.url,

                            organization=portal.portal,

                            title=label or href,

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
                    )

                continue


            try:

                child_response = (
                    self.http.get(
                        href,
                        timeout=REQUEST_TIMEOUT,
                        allow_redirects=True
                    )
                )


                if (
                    child_response.status_code
                    != 200
                ):

                    continue


                child = BeautifulSoup(
                    child_response.text,
                    "html.parser"
                )


                text = normalize_space(
                    child.get_text(
                        " ",
                        strip=True
                    )
                )


                hits = keyword_hits(
                    text
                    + " "
                    + label
                )


                if not hits:
                    continue


                documents = []


                for child_anchor in (
                    child.find_all(
                        "a",
                        href=True
                    )
                ):

                    child_url = (
                        safe_urljoin(
                            child_response.url,
                            child_anchor[
                                "href"
                            ]
                        )
                    )


                    if is_downloadable_url(
                        child_url
                    ):

                        documents.append(
                            child_url
                        )


                refs = (
                    extract_ref_candidates(
                        text
                    )
                )


                title = (
                    label
                    or text[:400]
                )


                results.append(
                    TenderCandidate(
                        portal=portal.portal,
                        state=portal.state,

                        source_url=response.url,

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
                        )[:20],

                        keyword_hits=hits,
                    )
                )


            except requests.RequestException:

                continue


        return self._dedupe(
            results
        )


    def _dedupe(
        self,
        items
    ):

        seen = set()
        output = []

        for tender in items:

            key = (
                tender.stable_key()
            )

            if key not in seen:

                seen.add(
                    key
                )

                output.append(
                    tender
                )

        return output


# =============================================================================
# 9. GEM CRAWLER
# =============================================================================

def crawl_gem():

    if sync_playwright is None:

        return [], CrawlHealth(
            "Government e-Marketplace (GeM)",
            GEM_LISTING_URL,
            "DEPENDENCY_MISSING",
            0,
            "playwright is not installed"
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
    ]


    results = []


    try:

        with sync_playwright() as playwright:

            browser = (
                playwright.chromium.launch(
                    headless=True,

                    args=[
                        "--no-sandbox",
                        "--disable-dev-shm-usage",
                    ]
                )
            )


            page = (
                browser.new_page()
            )


            page.goto(
                GEM_LISTING_URL,
                timeout=45000,
                wait_until="domcontentloaded"
            )


            page.wait_for_timeout(
                2500
            )


            search_ui_found = False


            for term in search_terms:

                search_input = None


                for selector in [
                    'input[placeholder*="Keyword" i]',
                    'input[placeholder*="Search" i]',
                    'input[type="search"]',
                    'input#search_by',
                ]:

                    locator = (
                        page.locator(
                            selector
                        ).first
                    )


                    if (
                        locator.count()
                        and locator.is_visible()
                    ):

                        search_input = (
                            locator
                        )

                        search_ui_found = True

                        break


                if search_input is None:

                    continue


                try:

                    search_input.fill(
                        ""
                    )

                    search_input.fill(
                        term
                    )


                    page.keyboard.press(
                        "Enter"
                    )


                    page.wait_for_timeout(
                        2000
                    )


                    body_text = (
                        page.locator(
                            "body"
                        ).inner_text()
                    )


                    for bid_no in re.findall(
                        r"GEM/\d{4}/B/\d+",
                        body_text
                    ):

                        bid_number = (
                            bid_no.split(
                                "/"
                            )[-1]
                        )


                        doc_url = (
                            "https://"
                            "bidplus-global.gem.gov.in/"
                            "showbidDocument/"
                            f"{bid_number}"
                        )


                        results.append(
                            TenderCandidate(
                                portal=(
                                    "Government "
                                    "e-Marketplace (GeM)"
                                ),

                                state="Pan India",

                                source_url=(
                                    GEM_LISTING_URL
                                ),

                                tender_id=bid_no,

                                organization=(
                                    "NOT VERIFIED"
                                ),

                                title=(
                                    f"{term} - "
                                    f"{bid_no}"
                                ),

                                category=infer_category(
                                    term
                                ),

                                detail_url=doc_url,

                                discovered_doc_urls=[
                                    doc_url
                                ],

                                keyword_hits=[
                                    term
                                ],
                            )
                        )


                except Exception as error:

                    log.warning(
                        "GeM search failed "
                        "for %s: %s",

                        term,
                        error
                    )


            browser.close()


            # IMPORTANT:
            # If we could not even find the search box,
            # do NOT report SUCCESS with zero tenders.

            if not search_ui_found:

                raise RuntimeError(
                    "GeM search UI not found. "
                    "Portal structure may have changed."
                )


        deduplicated = {
            item.tender_id
            or item.stable_key():
                item

            for item
            in results
        }


        output = list(
            deduplicated.values()
        )


        return (
            output,

            CrawlHealth(
                "Government e-Marketplace (GeM)",
                GEM_LISTING_URL,
                "SUCCESS",
                len(output)
            )
        )


    except Exception as error:

        return [], CrawlHealth(
            "Government e-Marketplace (GeM)",
            GEM_LISTING_URL,
            "PARSER_FAILED",
            0,
            str(error)[:250]
        )


# =============================================================================
# 10. DOCUMENT MANAGER
# =============================================================================

class DocumentManager:

    def __init__(
        self
    ):

        self.http = (
            session_with_headers()
        )


    def discover_from_detail_page(
        self,
        candidate
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

            response = (
                self.http.get(
                    candidate.detail_url,
                    timeout=REQUEST_TIMEOUT,
                    allow_redirects=True
                )
            )


            if (
                response.status_code
                != 200
            ):

                return list(
                    dict.fromkeys(
                        urls
                    )
                )


            soup = BeautifulSoup(
                response.text,
                "html.parser"
            )


            for anchor in soup.find_all(
                "a",
                href=True
            ):

                url = safe_urljoin(
                    response.url,
                    anchor[
                        "href"
                    ]
                )


                label = normalize_space(
                    anchor.get_text(
                        " ",
                        strip=True
                    )
                ).lower()


                if (
                    is_downloadable_url(
                        url
                    )
                    or any(
                        item
                        in label

                        for item in [
                            "download",
                            "tender document",
                            "rfp",
                            "atc",
                            "corrigendum",
                            "nit",
                            "boq",
                            "annexure",
                            "pre-bid",
                            "pre bid",
                        ]
                    )
                ):

                    urls.append(
                        url
                    )


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
        tender_key
    ):

        try:

            response = (
                self.http.get(
                    url,
                    timeout=REQUEST_TIMEOUT,
                    allow_redirects=True,
                    stream=True
                )
            )


            response.raise_for_status()


            content_type = (
                response.headers.get(
                    "content-type"
                )
                or ""
            ).lower()


            data = bytearray()


            for chunk in (
                response.iter_content(
                    65536
                )
            ):

                if not chunk:
                    continue


                data.extend(
                    chunk
                )


                if (
                    len(data)
                    >
                    MAX_DOC_BYTES
                ):

                    raise ValueError(
                        "Document exceeds "
                        "configured size limit"
                    )


            raw = bytes(
                data
            )


            sha256 = hashlib.sha256(
                raw
            ).hexdigest()


            extension = (
                self._guess_ext(
                    response.url,
                    content_type,
                    raw
                )
            )


            folder = (
                DOWNLOAD_DIR
                / tender_key
            )


            folder.mkdir(
                parents=True,
                exist_ok=True
            )


            path = (
                folder
                / (
                    f"{sha256[:12]}"
                    f"{extension}"
                )
            )


            path.write_bytes(
                raw
            )


            text = (
                self._extract_text(
                    raw,
                    extension,
                    content_type
                )
            )


            return DownloadedDocument(
                url=response.url,
                local_path=str(
                    path
                ),
                mime_type=content_type,
                text=text[
                    :MAX_DOC_TEXT_CHARS
                ],
                sha256=sha256
            )


        except Exception as error:

            return DownloadedDocument(
                url=url,
                local_path="",
                mime_type="",
                text="",
                sha256="",
                error=str(
                    error
                )[:400]
            )


    def _guess_ext(
        self,
        url,
        content_type,
        raw
    ):

        extension = Path(
            urllib.parse.urlsplit(
                url
            ).path
        ).suffix.lower()


        if (
            extension
            in DOWNLOADABLE_EXTENSIONS
        ):
            return extension


        if (
            raw.startswith(
                b"%PDF"
            )
            or
            "application/pdf"
            in content_type
        ):
            return ".pdf"


        if raw.startswith(
            b"PK\x03\x04"
        ):
            return ".zip"


        if (
            "wordprocessingml"
            in content_type
        ):
            return ".docx"


        if (
            "spreadsheetml"
            in content_type
        ):
            return ".xlsx"


        return ".bin"


    def _extract_text(
        self,
        raw,
        extension,
        content_type
    ):

        if extension == ".pdf":
            return self._pdf_text(
                raw
            )


        if extension == ".docx":
            return self._docx_text(
                raw
            )


        if extension == ".zip":
            return self._zip_text(
                raw
            )


        if extension == ".xlsx":
            return self._excel_text(
                raw
            )


        if (
            "text/"
            in content_type
        ):

            return raw.decode(
                "utf-8",
                errors="ignore"
            )


        return ""


    def _pdf_text(
        self,
        raw
    ):

        try:

            reader = PdfReader(
                io.BytesIO(
                    raw
                )
            )


            return "\n".join(
                page.extract_text()
                or ""

                for page
                in reader.pages
            )


        except Exception:

            return ""


    def _docx_text(
        self,
        raw
    ):

        try:

            document = DocxDocument(
                io.BytesIO(
                    raw
                )
            )


            return "\n".join(
                paragraph.text
                for paragraph
                in document.paragraphs
            )


        except Exception:

            return ""


    def _zip_text(
        self,
        raw
    ):

        chunks = []


        try:

            with zipfile.ZipFile(
                io.BytesIO(
                    raw
                )
            ) as archive:


                for info in (
                    archive.infolist()[:50]
                ):

                    extension = Path(
                        info.filename
                    ).suffix.lower()


                    if extension not in (
                        ".pdf",
                        ".docx",
                        ".txt",
                        ".xlsx"
                    ):
                        continue


                    data = archive.read(
                        info
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
                            errors="ignore"
                        )


                    if text:

                        chunks.append(
                            "\n--- FILE: "
                            f"{info.filename}"
                            " ---\n"
                            f"{text}"
                        )


        except Exception:

            pass


        return "\n".join(
            chunks
        )


    def _excel_text(
        self,
        raw
    ):

        try:

            excel = pd.ExcelFile(
                io.BytesIO(
                    raw
                )
            )


            parts = []


            for sheet in (
                excel.sheet_names[:15]
            ):

                dataframe = pd.read_excel(
                    io.BytesIO(
                        raw
                    ),
                    sheet_name=sheet,
                    header=None
                )


                parts.append(
                    "\n--- SHEET: "
                    f"{sheet}"
                    " ---\n"
                )


                parts.append(
                    dataframe
                    .astype(
                        str
                    )
                    .to_csv(
                        index=False,
                        header=False
                    )
                )


            return "".join(
                parts
            )


        except Exception:

            return ""


# =============================================================================
# 11. GEMINI ELIGIBILITY EXTRACTION
# =============================================================================

def extract_eligibility(
    candidate,
    documents
):

    usable_documents = [
        document
        for document
        in documents
        if document.text.strip()
    ]


    combined_text = "\n\n".join(
        (
            f"SOURCE URL: "
            f"{document.url}\n"

            f"LOCAL FILE: "
            f"{document.local_path}\n"

            f"{document.text}"
        )

        for document
        in usable_documents
    )[:MAX_DOC_TEXT_CHARS]


    if not combined_text:

        return {
            "extraction_status":
                "NO_DOCUMENT_TEXT",

            "missing_or_unclear_fields": [
                "Tender document text "
                "could not be extracted"
            ],

            "evidence_quotes":
                [],
        }


    if not GEMINI_API_KEY:

        return {
            "extraction_status":
                "DOCUMENT_TEXT_AVAILABLE_"
                "AI_NOT_CONFIGURED",

            "missing_or_unclear_fields": [
                "GEMINI_API_KEY "
                "is not configured"
            ],

            "evidence_quotes":
                [],
        }


    client = genai.Client(
        api_key=GEMINI_API_KEY
    )


    prompt = (
        "You are extracting procurement eligibility facts "
        "from government tender documents.\n\n"

        "STRICT RULES:\n"

        "- Extract ONLY facts explicitly supported by the document text.\n"

        "- Never assume MSME exemption, EMD exemption, turnover, "
        "experience rules, deadlines, 80/50/40 rules, or standard terms.\n"

        "- If a criterion is absent or unclear, return null.\n"

        "- INR amounts must be numeric rupees.\n"

        "- evidence_quotes must contain short exact source snippets.\n"

        "- Identify entity-type restrictions.\n"

        "- Identify JV / consortium permissions.\n"

        "- Identify physical submission requirements.\n"

        "- Identify QCBS or technical presentation requirements.\n"

        "- Identify celebrity / artist mandate requirements.\n"

        "- Add all important unclear or missing eligibility items "
        "to missing_or_unclear_fields.\n\n"

        "DISCOVERED TENDER:\n"

        + json.dumps(
            asdict(
                candidate
            ),
            ensure_ascii=False,
            indent=2
        )

        + "\n\nDOCUMENT TEXT:\n"

        + combined_text
    )


    try:

        response = (
            client.models.generate_content(
                model=GEMINI_MODEL,
                contents=prompt,

                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=EligibilityEvidence,
                    temperature=0,
                ),
            )
        )


        data = json.loads(
            response.text
        )


        data[
            "extraction_status"
        ] = (
            "DOCUMENT_VERIFIED_EXTRACTION"
        )


        return data


    except Exception as error:

        return {
            "extraction_status":
                "AI_EXTRACTION_FAILED",

            "missing_or_unclear_fields": [
                str(
                    error
                )[:350]
            ],

            "evidence_quotes":
                [],
        }


# =============================================================================
# 12. QUALIFICATION ENGINE
# =============================================================================

def evaluate_qualification(
    evidence
):

    # -------------------------------------------------------------------------
    # No verified extraction
    # -------------------------------------------------------------------------

    if (
        evidence.get(
            "extraction_status"
        )
        !=
        "DOCUMENT_VERIFIED_EXTRACTION"
    ):

        return QualificationDecision(
            "DOCUMENT REVIEW REQUIRED",

            "Verified structured eligibility evidence "
            "is not available.",

            "Open the source tender documents and review "
            "the eligibility clauses manually."
        )


    fail_reasons = []


    entity = (
        evidence.get(
            "entity_type_restriction"
        )
        or ""
    ).lower()


    turnover = (
        evidence.get(
            "average_turnover_required_inr"
        )
    )


    single_work = (
        evidence.get(
            "single_similar_work_required_inr"
        )
    )


    two_work = (
        evidence.get(
            "two_similar_works_each_required_inr"
        )
    )


    three_work = (
        evidence.get(
            "three_similar_works_each_required_inr"
        )
    )


    net_worth = (
        evidence.get(
            "min_net_worth_required_inr"
        )
    )


    jv_allowed = (
        evidence.get(
            "consortium_or_jv_allowed"
        )
    )


    # -------------------------------------------------------------------------
    # Entity restriction
    # -------------------------------------------------------------------------

    if (
        entity
        and (
            "private limited"
            in entity

            or "public limited"
            in entity

            or "company only"
            in entity
        )
        and
        "propriet"
        not in entity
    ):

        fail_reasons.append(
            "Entity type restriction "
            "conflicts with proprietorship."
        )


    # -------------------------------------------------------------------------
    # Turnover
    # -------------------------------------------------------------------------

    if (
        turnover is not None

        and turnover
        >
        COMPANY_PROFILE[
            "avg_3yr_turnover_inr"
        ]

        and evidence.get(
            "msme_turnover_relaxation_explicit"
        )
        is not True
    ):

        fail_reasons.append(
            "Required turnover "
            f"₹{turnover:,.0f} "
            "exceeds benchmark "
            f"₹"
            f"{COMPANY_PROFILE['avg_3yr_turnover_inr']:,.0f}."
        )


    # -------------------------------------------------------------------------
    # Single similar work
    # -------------------------------------------------------------------------

    if (
        single_work
        is not None

        and single_work
        >
        COMPANY_PROFILE[
            "max_single_past_work_order_inr"
        ]

        and jv_allowed
        is not True
    ):

        fail_reasons.append(
            "Single similar work requirement "
            f"₹{single_work:,.0f} "
            "exceeds benchmark "
            f"₹"
            f"{COMPANY_PROFILE['max_single_past_work_order_inr']:,.0f}."
        )


    # -------------------------------------------------------------------------
    # Two-work threshold
    # -------------------------------------------------------------------------

    if (
        two_work
        is not None

        and two_work
        >
        COMPANY_PROFILE[
            "max_two_works_threshold_inr"
        ]

        and jv_allowed
        is not True
    ):

        fail_reasons.append(
            "Two-work threshold "
            f"₹{two_work:,.0f} "
            "exceeds benchmark "
            f"₹"
            f"{COMPANY_PROFILE['max_two_works_threshold_inr']:,.0f}."
        )


    # -------------------------------------------------------------------------
    # Three-work threshold
    # -------------------------------------------------------------------------

    if (
        three_work
        is not None

        and three_work
        >
        COMPANY_PROFILE[
            "max_three_works_threshold_inr"
        ]

        and jv_allowed
        is not True
    ):

        fail_reasons.append(
            "Three-work threshold "
            f"₹{three_work:,.0f} "
            "exceeds benchmark "
            f"₹"
            f"{COMPANY_PROFILE['max_three_works_threshold_inr']:,.0f}."
        )


    # -------------------------------------------------------------------------
    # Net worth
    # -------------------------------------------------------------------------

    if (
        net_worth
        is not None

        and net_worth
        >
        COMPANY_PROFILE[
            "audited_net_worth_inr"
        ]
    ):

        fail_reasons.append(
            "Required net worth "
            f"₹{net_worth:,.0f} "
            "exceeds audited benchmark "
            f"₹"
            f"{COMPANY_PROFILE['audited_net_worth_inr']:,.0f}."
        )


    # -------------------------------------------------------------------------
    # Non-core scope
    # -------------------------------------------------------------------------

    if (
        evidence.get(
            "non_core_scope"
        )
        is True
    ):

        fail_reasons.append(
            "Tender scope is outside "
            "the configured event / "
            "experiential core scope."
        )


    # -------------------------------------------------------------------------
    # Hard disqualification
    # -------------------------------------------------------------------------

    if fail_reasons:

        return QualificationDecision(
            "VERIFIED DISQUALIFIED",

            " ".join(
                fail_reasons
            ),

            "Do not bid unless clarification, amendment "
            "or permitted JV resolves the blocker."
        )


    # -------------------------------------------------------------------------
    # Manual intervention
    # -------------------------------------------------------------------------

    manual_reasons = []


    if (
        evidence.get(
            "physical_submission_required"
        )
        is True
    ):

        manual_reasons.append(
            "Physical submission requirement"
        )


    if (
        evidence.get(
            "qcbs_or_technical_pitch"
        )
        is True
    ):

        manual_reasons.append(
            "QCBS / technical presentation"
        )


    if (
        evidence.get(
            "named_celebrity_or_artist_mandate"
        )
        is True
    ):

        manual_reasons.append(
            "Artist / celebrity mandate"
        )


    if manual_reasons:

        return QualificationDecision(
            "NEEDS MANUAL INTERVENTION",

            "; ".join(
                manual_reasons
            ),

            "Review submission logistics and "
            "commercial / technical feasibility."
        )


    # -------------------------------------------------------------------------
    # CRITICAL COMPLETENESS CHECK
    # -------------------------------------------------------------------------

    critical_missing = []


    if (
        evidence.get(
            "average_turnover_required_inr"
        )
        is None
    ):

        critical_missing.append(
            "Turnover requirement"
        )


    if (
        evidence.get(
            "single_similar_work_required_inr"
        )
        is None

        and evidence.get(
            "two_similar_works_each_required_inr"
        )
        is None

        and evidence.get(
            "three_similar_works_each_required_inr"
        )
        is None
    ):

        critical_missing.append(
            "Past experience requirement"
        )


    if (
        evidence.get(
            "entity_type_restriction"
        )
        is None
    ):

        critical_missing.append(
            "Entity-type eligibility"
        )


    # Use Gemini's own uncertainty list too

    unclear = (
        evidence.get(
            "missing_or_unclear_fields"
        )
        or []
    )


    critical_keywords = (
        "turnover",
        "experience",
        "similar work",
        "entity",
        "constitution",
        "net worth",
        "qualification",
        "eligibility",
    )


    for item in unclear:

        if any(
            keyword in str(
                item
            ).lower()

            for keyword
            in critical_keywords
        ):

            critical_missing.append(
                str(
                    item
                )
            )


    critical_missing = list(
        dict.fromkeys(
            critical_missing
        )
    )


    if critical_missing:

        return QualificationDecision(
            "DOCUMENT REVIEW REQUIRED",

            "Critical eligibility fields "
            "could not be fully verified: "
            + "; ".join(
                critical_missing[:8]
            ),

            "Review the eligibility/PQC/ATC clauses "
            "before treating this tender as qualified."
        )


    # -------------------------------------------------------------------------
    # Verified qualified
    # -------------------------------------------------------------------------

    return QualificationDecision(
        "VERIFIED QUALIFIED",

        "No disqualifying condition was found "
        "in the verified eligibility evidence.",

        "Proceed to technical / commercial bid preparation "
        "and verify the latest corrigenda before submission."
    )


# =============================================================================
# 13. GOOGLE SHEET COLUMNS
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
# 14. GOOGLE SHEET CONNECTION
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
            scopes=scopes
        )
    )


    return gspread.authorize(
        credentials
    )


def ensure_worksheet(
    spreadsheet,
    name,
    rows=2000,
    cols=30
):

    try:

        return spreadsheet.worksheet(
            name
        )

    except gspread.WorksheetNotFound:

        return spreadsheet.add_worksheet(
            title=name,
            rows=rows,
            cols=cols
        )


def existing_tender_keys(
    sheet
):

    values = (
        sheet.get_all_values()
    )


    if not values:
        return set()


    header = values[
        0
    ]


    try:

        index = header.index(
            "Tender Key"
        )

    except ValueError:

        return set()


    return {
        row[index]
        for row
        in values[1:]
        if (
            len(row) > index
            and row[index]
        )
    }


# =============================================================================
# 15. PORTAL STATUS SYNC
# =============================================================================

def sync_portal_directory(
    sheet,
    portals,
    health
):

    rows = [
        PORTAL_HEADERS
    ]


    for portal in portals:

        status = health.get(
            portal.portal
        )


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
        range_name=
            f"A1:J{len(rows)}"
    )


# =============================================================================
# 16. PROCESS TENDER
# =============================================================================

def process_candidate(
    candidate,
    document_manager
):

    document_urls = (
        document_manager
        .discover_from_detail_page(
            candidate
        )
    )


    documents = [
        document_manager
        .download_and_extract(
            url,
            candidate.stable_key()
        )

        for url
        in document_urls[:20]
    ]


    evidence = (
        extract_eligibility(
            candidate,
            documents
        )
    )


    decision = (
        evaluate_qualification(
            evidence
        )
    )


    return (
        evidence,
        documents,
        decision
    )


# =============================================================================
# 17. MAIN PIPELINE
# =============================================================================

def run_pipeline():

    started = (
        time.time()
    )


    log.info(
        "%s started",
        APP_NAME
    )


    # -------------------------------------------------------------------------
    # Load portals
    # -------------------------------------------------------------------------

    portals = (
        load_master_portals()
    )


    log.info(
        "Loaded %d active portals.",
        len(
            portals
        )
    )


    crawler = (
        TenderCrawler()
    )


    all_candidates = []

    health = {}


    # -------------------------------------------------------------------------
    # Generic / GePNIC portals
    # -------------------------------------------------------------------------

    non_gem_portals = [
        portal
        for portal
        in portals

        if (
            "gem.gov.in"
            not in portal.url.lower()
        )
    ]


    with ThreadPoolExecutor(
        max_workers=
            CRAWL_WORKERS
    ) as executor:


        futures = {
            executor.submit(
                crawler.crawl,
                portal
            ):
                portal

            for portal
            in non_gem_portals
        }


        for future in (
            as_completed(
                futures
            )
        ):

            portal = (
                futures[
                    future
                ]
            )


            try:

                items, status = (
                    future.result()
                )


            except Exception as error:

                items = []

                status = CrawlHealth(
                    portal.portal,
                    portal.url,
                    "PARSER_FAILED",
                    0,
                    str(
                        error
                    )[:250]
                )


            health[
                portal.portal
            ] = status


            all_candidates.extend(
                items
            )


    # -------------------------------------------------------------------------
    # GeM
    # -------------------------------------------------------------------------

    gem_items, gem_health = (
        crawl_gem()
    )


    all_candidates.extend(
        gem_items
    )


    health[
        "Government e-Marketplace (GeM)"
    ] = gem_health


    # -------------------------------------------------------------------------
    # Deduplicate
    # -------------------------------------------------------------------------

    candidate_map = {
        candidate.stable_key():
            candidate

        for candidate
        in all_candidates
    }


    candidates = list(
        candidate_map.values()
    )


    log.info(
        "Unique tender candidates: %d",
        len(
            candidates
        )
    )


    # -------------------------------------------------------------------------
    # Google Sheets
    # -------------------------------------------------------------------------

    google_client = (
        get_gspread_client()
    )


    spreadsheet = (
        google_client.open_by_key(
            SPREADSHEET_ID
        )
    )


    active_sheet = (
        ensure_worksheet(
            spreadsheet,
            "Active_Tenders",
            rows=5000,
            cols=30
        )
    )


    portal_sheet = (
        ensure_worksheet(
            spreadsheet,
            "Portal_Directory",
            rows=max(
                500,
                len(
                    portals
                ) + 50
            ),
            cols=12
        )
    )


    active_sheet.update(
        values=[
            ACTIVE_HEADERS
        ],
        range_name=
            "A1:Y1"
    )


    existing = (
        existing_tender_keys(
            active_sheet
        )
    )


    document_manager = (
        DocumentManager()
    )


    new_rows = []


    # =========================================================================
    # PROCESS NEW TENDERS
    # =========================================================================

    for candidate in candidates:

        key = (
            candidate.stable_key()
        )


        if key in existing:
            continue


        log.info(
            "Processing tender: %s",
            candidate.tender_id
            or key
        )


        evidence, documents, decision = (
            process_candidate(
                candidate,
                document_manager
            )
        )


        downloaded_files = [
            document.local_path
            for document
            in documents
            if document.local_path
        ]


        document_urls = list(
            dict.fromkeys(
                candidate.discovered_doc_urls
                + [
                    document.url
                    for document
                    in documents
                    if document.url
                ]
            )
        )


        # ---------------------------------------------------------------------
        # Experience text
        # ---------------------------------------------------------------------

        experience_parts = []


        if (
            evidence.get(
                "single_similar_work_required_inr"
            )
            is not None
        ):

            experience_parts.append(
                "1 work >= "
                f"₹"
                f"{evidence['single_similar_work_required_inr']:,.0f}"
            )


        if (
            evidence.get(
                "two_similar_works_each_required_inr"
            )
            is not None
        ):

            experience_parts.append(
                "2 works each >= "
                f"₹"
                f"{evidence['two_similar_works_each_required_inr']:,.0f}"
            )


        if (
            evidence.get(
                "three_similar_works_each_required_inr"
            )
            is not None
        ):

            experience_parts.append(
                "3 works each >= "
                f"₹"
                f"{evidence['three_similar_works_each_required_inr']:,.0f}"
            )


        experience = (
            "; ".join(
                experience_parts
            )
            or
            "NOT VERIFIED / NOT STATED"
        )


        # ---------------------------------------------------------------------
        # Compliance flags
        # ---------------------------------------------------------------------

        key_compliance = []


        if (
            evidence.get(
                "physical_submission_required"
            )
            is True
        ):

            key_compliance.append(
                "Physical submission required"
            )


        if (
            evidence.get(
                "qcbs_or_technical_pitch"
            )
            is True
        ):

            key_compliance.append(
                "QCBS / technical pitch"
            )


        if (
            evidence.get(
                "named_celebrity_or_artist_mandate"
            )
            is True
        ):

            key_compliance.append(
                "Artist / celebrity mandate"
            )


        remarks = (
            f"{decision.reason} "
            f"Action: "
            f"{decision.action_plan}"
        )


        # ---------------------------------------------------------------------
        # Sheet row
        # ---------------------------------------------------------------------

        new_rows.append([

            dt.datetime.now(
                tz=IST
            ).strftime(
                "%Y-%m-%d"
            ),

            key,

            evidence.get(
                "tender_reference"
            )
            or candidate.tender_id
            or "NOT VERIFIED",

            candidate.portal,

            candidate.state,

            evidence.get(
                "organisation"
            )
            or candidate.organization
            or "NOT VERIFIED",

            evidence.get(
                "scope_summary"
            )
            or candidate.title,

            candidate.category,

            (
                evidence.get(
                    "estimated_value_inr"
                )
                if evidence.get(
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
                if evidence.get(
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
            or candidate.deadline_raw
            or "NOT VERIFIED",

            (
                evidence.get(
                    "average_turnover_required_inr"
                )
                if evidence.get(
                    "average_turnover_required_inr"
                )
                is not None
                else
                "NOT VERIFIED / NOT STATED"
            ),

            experience,

            evidence.get(
                "entity_type_restriction"
            )
            or
            "NOT VERIFIED / NOT STATED",

            (
                "Yes"
                if evidence.get(
                    "consortium_or_jv_allowed"
                )
                is True

                else
                "No"
                if evidence.get(
                    "consortium_or_jv_allowed"
                )
                is False

                else
                "NOT VERIFIED / NOT STATED"
            ),

            (
                "; ".join(
                    key_compliance
                )
                or
                "No special flag extracted"
            ),

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
            or candidate.source_url,

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
                "UNKNOWN"
            ),
        ])


    # -------------------------------------------------------------------------
    # Write new tenders
    # -------------------------------------------------------------------------

    if new_rows:

        active_sheet.append_rows(
            new_rows,
            value_input_option="RAW"
        )


        log.info(
            "Added %d new tenders.",
            len(
                new_rows
            )
        )


    else:

        log.info(
            "No new tenders to add."
        )


    # -------------------------------------------------------------------------
    # Update portal health
    # -------------------------------------------------------------------------

    sync_portal_directory(
        portal_sheet,
        portals,
        health
    )


    elapsed = (
        time.time()
        - started
    )


    log.info(
        "Completed in %.1f seconds.",
        elapsed
    )


# =============================================================================
# RUN
# =============================================================================

if __name__ == "__main__":

    run_pipeline()
