
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

APP_NAME = "Pan-India Tender Intelligence Agent"
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE", "service_account.json")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1YgsKZNiamaZ-E2ZZhIY2MLm8diC5okmeUMGFQdgQzHQ")
EXCEL_MASTER_FILE = os.getenv("EXCEL_MASTER_FILE", "Pan_India_Tender_URL_Master.xlsx")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash-lite")
GEM_LISTING_URL = os.getenv("GEM_LISTING_URL", "https://bidplus-global.gem.gov.in/")

DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "tender_downloads"))
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "18"))
CRAWL_WORKERS = int(os.getenv("CRAWL_WORKERS", "12"))
MAX_DOC_BYTES = int(os.getenv("MAX_DOC_BYTES", str(30 * 1024 * 1024)))
MAX_DOC_TEXT_CHARS = int(os.getenv("MAX_DOC_TEXT_CHARS", "120000"))
MAX_GENERIC_LINKS = int(os.getenv("MAX_GENERIC_LINKS", "25"))

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s | %(levelname)s | %(message)s",
)
log = logging.getLogger("tender-agent")

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

EVENT_KEYWORDS = [
    "Event Management", "Event Agency", "Event Management Agency",
    "Exhibition", "Exhibition Stall", "Exhibition Pavilion",
    "Design and Fabrication", "Fabrication", "Expo", "Trade Fair", "Mela",
    "Festival", "Conference", "Seminar", "Convention", "Summit", "Conclave",
    "Roadshow", "Dealer Meet", "Annual Day", "Foundation Day", "Award Ceremony",
    "Cultural Programme", "Corporate Event", "Activation", "Brand Activation",
    "Experiential Marketing", "Publicity", "Outreach Campaign", "IEC",
    "Media Campaign", "Advertising Agency", "Creative Agency", "AV Production",
    "Audio Visual", "LED", "Sound and Light", "Stage", "Tentage", "Decoration",
    "Venue Management", "Hospitality", "Manpower", "Printing", "Branding",
    "Signage", "Digital Marketing", "Social Media", "PR Agency",
    "Communication Agency", "Empanelment of Event Agency",
    "Empanelment of Advertising Agency", "EOI Event", "RFP Event",
    "Tourism Event", "Sports Event", "Government Function",
    "Launch Event", "Inauguration",
]

TENDERISH_TERMS = (
    "tender", "bid", "rfp", "eoi", "procurement", "quotation",
    "corrigendum", "nit", "empanelment", "notice inviting"
)

DOWNLOADABLE_EXTENSIONS = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip")

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
        default_factory=lambda: dt.datetime.now(tz=IST).strftime("%Y-%m-%d %H:%M:%S IST")
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
    discovered_doc_urls: list[str] = field(default_factory=list)
    keyword_hits: list[str] = field(default_factory=list)

    def stable_key(self):
        raw = "|".join([
            (self.tender_id or "").strip().lower(),
            (self.organization or "").strip().lower(),
            (self.title or "").strip().lower(),
            (self.deadline_raw or "").strip().lower(),
            (self.detail_url or self.source_url or "").strip().lower(),
        ])
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]

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

def normalize_space(value):
    return re.sub(r"\s+", " ", value or "").strip()

def keyword_hits(text):
    t = (text or "").lower()
    return [kw for kw in EVENT_KEYWORDS if kw.lower() in t]

def infer_category(title):
    t = (title or "").lower()
    if "exhibition" in t or "stall" in t or "pavilion" in t:
        return "Exhibition / Stall"
    if "empanel" in t:
        return "Empanelment"
    if any(x in t for x in ["conference", "summit", "conclave", "seminar"]):
        return "Conference / Conclave"
    if any(x in t for x in ["advertising", "media campaign", "creative agency"]):
        return "Advertising / Creative / Outreach"
    if any(x in t for x in ["festival", "mela", "cultural"]):
        return "Festival / Mela / Cultural"
    return "Event / Experiential"

def safe_urljoin(base, href):
    return urllib.parse.urljoin(base, href or "")

def is_downloadable_url(url):
    path = urllib.parse.urlsplit(url).path.lower()
    return path.endswith(DOWNLOADABLE_EXTENSIONS)

def extract_ref_candidates(text):
    patterns = [
        r"GEM/\d{4}/B/\d+",
        r"\b\d{4}_[A-Za-z0-9_-]+_\d+_\d+\b",
        r"\b(?:RFP|EOI|NIT|TENDER|BID)[\s:/-]*[A-Za-z0-9._/-]{3,}\b",
    ]
    out = []
    for p in patterns:
        out.extend(re.findall(p, text or "", flags=re.I))
    return list(dict.fromkeys(normalize_space(x) for x in out))

def session_with_headers():
    s = requests.Session()
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/128.0 Safari/537.36"
        ),
        "Accept-Language": "en-IN,en;q=0.9",
    })
    return s

def load_master_portals():
    if not os.path.exists(EXCEL_MASTER_FILE):
        raise FileNotFoundError(f"Master Excel not found: {EXCEL_MASTER_FILE}")

    df = pd.read_excel(EXCEL_MASTER_FILE, sheet_name="Tender URL Master", skiprows=2)
    portals = []

    for _, r in df.iterrows():
        if str(r.get("Active", "")).strip().lower() != "yes":
            continue

        url = str(r.get("URL", "")).strip()
        if not url.startswith(("http://", "https://")):
            continue

        portals.append(
            Portal(
                portal=str(r.get("Portal / Organisation", "")).strip(),
                category=str(r.get("Category", "")).strip(),
                state=str(r.get("State/Region", "")).strip(),
                url=url,
                priority=str(r.get("Priority", "P2")).strip().upper() or "P2",
                source_type=str(r.get("Source Type", "")).strip(),
            )
        )

    return portals

class TenderCrawler:
    def __init__(self):
        self.http = session_with_headers()

    def crawl(self, portal):
        try:
            if "gem.gov.in" in portal.url.lower():
                return [], CrawlHealth(portal.portal, portal.url, "SKIPPED_DEDICATED_ADAPTER")

            if self._looks_gepnic(portal):
                items = self.crawl_gepnic(portal)
            else:
                items = self.crawl_generic(portal)

            return items, CrawlHealth(portal.portal, portal.url, "SUCCESS", len(items))

        except requests.exceptions.Timeout:
            return [], CrawlHealth(portal.portal, portal.url, "TIMEOUT", 0, "HTTP request timed out")
        except requests.exceptions.HTTPError as e:
            code = getattr(e.response, "status_code", None)
            status = "LOGIN_OR_BLOCKED" if code in (401, 403) else "HTTP_ERROR"
            return [], CrawlHealth(portal.portal, portal.url, status, 0, str(e)[:250])
        except RuntimeError as e:
            msg = str(e)
            status = "CAPTCHA_OR_LOGIN" if "CAPTCHA" in msg or "LOGIN" in msg else "PARSER_FAILED"
            return [], CrawlHealth(portal.portal, portal.url, status, 0, msg[:250])
        except Exception as e:
            return [], CrawlHealth(portal.portal, portal.url, "PARSER_FAILED", 0, str(e)[:250])

    def _looks_gepnic(self, portal):
        u = portal.url.lower()
        return any(
            x in u for x in [
                "eprocure.gov.in", "etenders.gov.in", "tenders.gov.in",
                "tenders.nic.in", "etender", "eproc", ".nic.in",
            ]
        )

    def crawl_gepnic(self, portal):
        base = portal.url.rstrip("/")
        candidate_urls = [
            f"{base}/nicgep/app?page=FrontEndLatestActiveTenders&service=page",
            f"{base}/eprocure/app?page=FrontEndLatestActiveTenders&service=page",
            f"{base}/app?page=FrontEndLatestActiveTenders&service=page",
            base,
        ]

        html = ""
        final_url = ""

        for u in dict.fromkeys(candidate_urls):
            try:
                r = self.http.get(u, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                if r.status_code == 200 and len(r.text) > 500:
                    html = r.text
                    final_url = r.url
                    break
            except requests.RequestException:
                continue

        if not html:
            raise RuntimeError("No usable GePNIC page returned")

        if "captcha" in html.lower() and "latest active tenders" not in html.lower():
            raise RuntimeError("CAPTCHA_OR_LOGIN_REQUIRED")

        soup = BeautifulSoup(html, "html.parser")
        out = []

        for row in soup.find_all("tr"):
            text = normalize_space(row.get_text(" ", strip=True))
            hits = keyword_hits(text)
            if not hits:
                continue

            cells = [normalize_space(td.get_text(" ", strip=True)) for td in row.find_all("td")]
            if len(cells) < 2:
                continue

            refs = extract_ref_candidates(text)
            tender_id = refs[0] if refs else ""
            title = max(cells, key=len)
            detail_url = final_url
            docs = []

            for a in row.find_all("a", href=True):
                href = safe_urljoin(final_url, a.get("href"))
                if not href or href.lower().startswith("javascript:"):
                    continue
                if is_downloadable_url(href):
                    docs.append(href)
                elif any(k in (a.get_text(" ", strip=True) + " " + href).lower() for k in TENDERISH_TERMS):
                    detail_url = href

            out.append(
                TenderCandidate(
                    portal=portal.portal,
                    state=portal.state,
                    source_url=final_url,
                    tender_id=tender_id,
                    organization=portal.portal,
                    title=title[:500],
                    category=infer_category(title),
                    detail_url=detail_url,
                    discovered_doc_urls=list(dict.fromkeys(docs)),
                    keyword_hits=hits,
                )
            )

        return self._dedupe(out)

    def crawl_generic(self, portal):
        r = self.http.get(portal.url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        r.raise_for_status()
        if "captcha" in r.text.lower() and len(r.text) < 50000:
            raise RuntimeError("CAPTCHA_OR_LOGIN_REQUIRED")
        soup = BeautifulSoup(r.text, "html.parser")

        links = []
        for a in soup.find_all("a", href=True):
            label = normalize_space(a.get_text(" ", strip=True))
            href = safe_urljoin(r.url, a["href"])
            combined = (label + " " + href).lower()
            if any(t in combined for t in TENDERISH_TERMS):
                links.append((label, href))

        out = []

        for label, href in list(dict.fromkeys(links))[:MAX_GENERIC_LINKS]:
            if is_downloadable_url(href):
                if keyword_hits(label):
                    out.append(
                        TenderCandidate(
                            portal=portal.portal,
                            state=portal.state,
                            source_url=r.url,
                            organization=portal.portal,
                            title=label or href,
                            category=infer_category(label),
                            detail_url=href,
                            discovered_doc_urls=[href],
                            keyword_hits=keyword_hits(label),
                        )
                    )
                continue

            try:
                rr = self.http.get(href, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                if rr.status_code != 200:
                    continue

                child = BeautifulSoup(rr.text, "html.parser")
                text = normalize_space(child.get_text(" ", strip=True))
                hits = keyword_hits(text + " " + label)

                if not hits:
                    continue

                docs = []
                for a2 in child.find_all("a", href=True):
                    u2 = safe_urljoin(rr.url, a2["href"])
                    if is_downloadable_url(u2):
                        docs.append(u2)

                refs = extract_ref_candidates(text)

                out.append(
                    TenderCandidate(
                        portal=portal.portal,
                        state=portal.state,
                        source_url=r.url,
                        tender_id=refs[0] if refs else "",
                        organization=portal.portal,
                        title=(label or text[:400])[:500],
                        category=infer_category(label or text[:400]),
                        detail_url=rr.url,
                        discovered_doc_urls=list(dict.fromkeys(docs))[:20],
                        keyword_hits=hits,
                    )
                )

            except requests.RequestException:
                continue

        return self._dedupe(out)

    def _dedupe(self, items):
        seen = set()
        out = []
        for t in items:
            key = t.stable_key()
            if key not in seen:
                seen.add(key)
                out.append(t)
        return out

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
        "Event Management", "Exhibition", "Conference", "Conclave",
        "Summit", "Mela", "Festival", "Creative Agency",
        "Advertising Agency", "Brand Activation", "Audio Visual",
        "Sound and Light", "Empanelment Event", "Outreach Campaign",
    ]

    results = []

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage"]
            )
            page = browser.new_page()
            page.goto(GEM_LISTING_URL, timeout=45000, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)

            for term in search_terms:
                try:
                    search_input = None
                    for sel in [
                        'input[placeholder*="Keyword" i]',
                        'input[placeholder*="Search" i]',
                        'input[type="search"]',
                        'input#search_by',
                    ]:
                        loc = page.locator(sel).first
                        if loc.count() and loc.is_visible():
                            search_input = loc
                            break

                    if search_input is None:
                        continue

                    search_input.fill("")
                    search_input.fill(term)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(2000)

                    body_text = page.locator("body").inner_text()

                    for bid_no in re.findall(r"GEM/\d{4}/B/\d+", body_text):
                        bid_num = bid_no.split("/")[-1]
                        doc_url = f"https://bidplus-global.gem.gov.in/showbidDocument/{bid_num}"

                        results.append(
                            TenderCandidate(
                                portal="Government e-Marketplace (GeM)",
                                state="Pan India",
                                source_url=GEM_LISTING_URL,
                                tender_id=bid_no,
                                organization="NOT VERIFIED",
                                title=f"{term} - {bid_no}",
                                category=infer_category(term),
                                detail_url=doc_url,
                                discovered_doc_urls=[doc_url],
                                keyword_hits=[term],
                            )
                        )

                except Exception as e:
                    log.warning("GeM search failed for %s: %s", term, e)

            browser.close()

        dedupe = {x.tender_id or x.stable_key(): x for x in results}
        out = list(dedupe.values())

        return out, CrawlHealth(
            "Government e-Marketplace (GeM)",
            GEM_LISTING_URL,
            "SUCCESS",
            len(out)
        )

    except Exception as e:
        return [], CrawlHealth(
            "Government e-Marketplace (GeM)",
            GEM_LISTING_URL,
            "PARSER_FAILED",
            0,
            str(e)[:250]
        )

class DocumentManager:
    def __init__(self):
        self.http = session_with_headers()

    def discover_from_detail_page(self, candidate):
        urls = list(candidate.discovered_doc_urls)

        if not candidate.detail_url:
            return list(dict.fromkeys(urls))

        if is_downloadable_url(candidate.detail_url):
            urls.append(candidate.detail_url)
            return list(dict.fromkeys(urls))

        try:
            r = self.http.get(candidate.detail_url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            if r.status_code != 200:
                return list(dict.fromkeys(urls))

            soup = BeautifulSoup(r.text, "html.parser")

            for a in soup.find_all("a", href=True):
                u = safe_urljoin(r.url, a["href"])
                label = normalize_space(a.get_text(" ", strip=True)).lower()

                if is_downloadable_url(u) or any(
                    x in label
                    for x in [
                        "download", "tender document", "rfp", "atc",
                        "corrigendum", "nit", "boq", "annexure",
                        "pre-bid", "pre bid"
                    ]
                ):
                    urls.append(u)

        except requests.RequestException:
            pass

        return list(dict.fromkeys(urls))

    def download_and_extract(self, url, tender_key):
        try:
            r = self.http.get(
                url,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
                stream=True
            )
            r.raise_for_status()

            ctype = (r.headers.get("content-type") or "").lower()
            data = bytearray()

            for chunk in r.iter_content(65536):
                if not chunk:
                    continue
                data.extend(chunk)
                if len(data) > MAX_DOC_BYTES:
                    raise ValueError("Document exceeds configured size limit")

            raw = bytes(data)
            sha = hashlib.sha256(raw).hexdigest()
            ext = self._guess_ext(r.url, ctype, raw)

            folder = DOWNLOAD_DIR / tender_key
            folder.mkdir(parents=True, exist_ok=True)

            path = folder / f"{sha[:12]}{ext}"
            path.write_bytes(raw)

            text = self._extract_text(raw, ext, ctype)

            return DownloadedDocument(
                url=r.url,
                local_path=str(path),
                mime_type=ctype,
                text=text[:MAX_DOC_TEXT_CHARS],
                sha256=sha
            )

        except Exception as e:
            return DownloadedDocument(
                url=url,
                local_path="",
                mime_type="",
                text="",
                sha256="",
                error=str(e)[:400]
            )

    def _guess_ext(self, url, ctype, raw):
        ext = Path(urllib.parse.urlsplit(url).path).suffix.lower()
        if ext in DOWNLOADABLE_EXTENSIONS:
            return ext
        if raw.startswith(b"%PDF") or "application/pdf" in ctype:
            return ".pdf"
        if raw.startswith(b"PK\x03\x04"):
            return ".zip"
        if "wordprocessingml" in ctype:
            return ".docx"
        if "spreadsheetml" in ctype:
            return ".xlsx"
        return ".bin"

    def _extract_text(self, raw, ext, ctype):
        if ext == ".pdf":
            return self._pdf_text(raw)
        if ext == ".docx":
            return self._docx_text(raw)
        if ext == ".zip":
            return self._zip_text(raw)
        if ext == ".xlsx":
            return self._excel_text(raw)
        if "text/" in ctype:
            return raw.decode("utf-8", errors="ignore")
        return ""

    def _pdf_text(self, raw):
        try:
            reader = PdfReader(io.BytesIO(raw))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception:
            return ""

    def _docx_text(self, raw):
        try:
            doc = DocxDocument(io.BytesIO(raw))
            return "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            return ""

    def _zip_text(self, raw):
        chunks = []

        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as zf:
                for info in zf.infolist()[:50]:
                    ext = Path(info.filename).suffix.lower()
                    if ext not in (".pdf", ".docx", ".txt", ".xlsx"):
                        continue

                    data = zf.read(info)

                    if ext == ".pdf":
                        text = self._pdf_text(data)
                    elif ext == ".docx":
                        text = self._docx_text(data)
                    elif ext == ".xlsx":
                        text = self._excel_text(data)
                    else:
                        text = data.decode("utf-8", errors="ignore")

                    if text:
                        chunks.append(f"\n--- FILE: {info.filename} ---\n{text}")

        except Exception:
            pass

        return "\n".join(chunks)

    def _excel_text(self, raw):
        try:
            xls = pd.ExcelFile(io.BytesIO(raw))
            parts = []

            for sheet in xls.sheet_names[:15]:
                df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, header=None)
                parts.append(f"\n--- SHEET: {sheet} ---\n")
                parts.append(df.astype(str).to_csv(index=False, header=False))

            return "".join(parts)

        except Exception:
            return ""

def extract_eligibility(candidate, documents):
    usable = [d for d in documents if d.text.strip()]
    combined = "\n\n".join(
        f"SOURCE URL: {d.url}\nLOCAL FILE: {d.local_path}\n{d.text}"
        for d in usable
    )[:MAX_DOC_TEXT_CHARS]

    if not combined:
        return {
            "extraction_status": "NO_DOCUMENT_TEXT",
            "missing_or_unclear_fields": [
                "Tender document text could not be extracted"
            ],
            "evidence_quotes": [],
        }

    if not GEMINI_API_KEY:
        return {
            "extraction_status": "DOCUMENT_TEXT_AVAILABLE_AI_NOT_CONFIGURED",
            "missing_or_unclear_fields": [
                "GEMINI_API_KEY is not configured"
            ],
            "evidence_quotes": [],
        }

    client = genai.Client(api_key=GEMINI_API_KEY)

    prompt = (
        "You are extracting procurement eligibility facts from government tender documents.\n\n"
        "STRICT RULES:\n"
        "- Extract ONLY facts explicitly supported by the document text.\n"
        "- Never assume MSME exemption, EMD exemption, turnover, experience rules, deadline, or standard terms.\n"
        "- If absent or unclear, return null.\n"
        "- INR amounts must be numeric rupees.\n"
        "- evidence_quotes must contain short exact source snippets.\n"
        "- Identify entity-type restrictions, JV/consortium rules, physical submission, QCBS/pitch, and artist mandate.\n\n"
        "DISCOVERED TENDER:\n"
        + json.dumps(asdict(candidate), ensure_ascii=False, indent=2)
        + "\n\nDOCUMENT TEXT:\n"
        + combined
    )

    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=EligibilityEvidence,
                temperature=0,
            ),
        )

        data = json.loads(response.text)
        data["extraction_status"] = "DOCUMENT_VERIFIED_EXTRACTION"
        return data

    except Exception as e:
        return {
            "extraction_status": "AI_EXTRACTION_FAILED",
            "missing_or_unclear_fields": [str(e)[:350]],
            "evidence_quotes": [],
        }

def evaluate_qualification(e):
    if e.get("extraction_status") != "DOCUMENT_VERIFIED_EXTRACTION":
        return QualificationDecision(
            "DOCUMENT REVIEW REQUIRED",
            "Verified structured eligibility evidence is not available.",
            "Open the source tender documents and review the eligibility clauses manually."
        )

    fail = []
    entity = (e.get("entity_type_restriction") or "").lower()
    turnover = e.get("average_turnover_required_inr")
    single = e.get("single_similar_work_required_inr")
    two_each = e.get("two_similar_works_each_required_inr")
    three_each = e.get("three_similar_works_each_required_inr")
    networth = e.get("min_net_worth_required_inr")
    jv_allowed = e.get("consortium_or_jv_allowed")

    if (
        entity
        and (
            "private limited" in entity
            or "public limited" in entity
            or "company only" in entity
        )
        and "propriet" not in entity
    ):
        fail.append("Entity type restriction conflicts with proprietorship.")

    if (
        turnover is not None
        and turnover > COMPANY_PROFILE["avg_3yr_turnover_inr"]
        and e.get("msme_turnover_relaxation_explicit") is not True
    ):
        fail.append(
            f"Required turnover ₹{turnover:,.0f} exceeds benchmark "
            f"₹{COMPANY_PROFILE['avg_3yr_turnover_inr']:,.0f}."
        )

    if (
        single is not None
        and single > COMPANY_PROFILE["max_single_past_work_order_inr"]
        and jv_allowed is not True
    ):
        fail.append(
            f"Single similar work requirement ₹{single:,.0f} exceeds benchmark "
            f"₹{COMPANY_PROFILE['max_single_past_work_order_inr']:,.0f}."
        )

    if (
        two_each is not None
        and two_each > COMPANY_PROFILE["max_two_works_threshold_inr"]
        and jv_allowed is not True
    ):
        fail.append(
            f"Two-work threshold ₹{two_each:,.0f} exceeds benchmark "
            f"₹{COMPANY_PROFILE['max_two_works_threshold_inr']:,.0f}."
        )

    if (
        three_each is not None
        and three_each > COMPANY_PROFILE["max_three_works_threshold_inr"]
        and jv_allowed is not True
    ):
        fail.append(
            f"Three-work threshold ₹{three_each:,.0f} exceeds benchmark "
            f"₹{COMPANY_PROFILE['max_three_works_threshold_inr']:,.0f}."
        )

    if (
        networth is not None
        and networth > COMPANY_PROFILE["audited_net_worth_inr"]
    ):
        fail.append(
            f"Required net worth ₹{networth:,.0f} exceeds audited benchmark "
            f"₹{COMPANY_PROFILE['audited_net_worth_inr']:,.0f}."
        )

    if e.get("non_core_scope") is True:
        fail.append("Tender scope is outside configured event/experiential core scope.")

    if fail:
        return QualificationDecision(
            "VERIFIED DISQUALIFIED",
            " ".join(fail),
            "Do not bid unless clarification, amendment or permitted JV resolves the blocker."
        )

    manual = []

    if e.get("physical_submission_required") is True:
        manual.append("Physical submission requirement")

    if e.get("qcbs_or_technical_pitch") is True:
        manual.append("QCBS / technical presentation")

    if e.get("named_celebrity_or_artist_mandate") is True:
        manual.append("Artist / celebrity mandate")

    if manual:
        return QualificationDecision(
            "NEEDS MANUAL INTERVENTION",
            "; ".join(manual),
            "Review submission logistics and commercial/technical feasibility."
        )

    return QualificationDecision(
        "VERIFIED QUALIFIED",
        "No disqualifying condition was found in the extracted eligibility evidence.",
        "Proceed to technical/commercial bid preparation and verify latest corrigenda."
    )

ACTIVE_HEADERS = [
    "Date Found", "Tender Key", "Tender ID / Ref No", "Portal Name",
    "State", "Organization / Dept", "Tender Title & Scope", "Category",
    "Estimated Value (INR)", "EMD (INR)", "EMD/MSME Exemption Evidence",
    "Submission Deadline", "Min Turnover Req", "Past Experience Requirement",
    "Entity Restriction", "JV/Consortium Allowed", "Key Compliance",
    "RFP / Tender Doc URLs", "Downloaded Files", "Pre-Bid Date",
    "Portal / Detail Link", "Qualification Status", "Remarks & Action Plan",
    "Evidence Quotes", "Extraction Status",
]

PORTAL_HEADERS = [
    "Portal", "Category", "State", "URL", "Active", "Last Crawl",
    "Tender Count", "Crawler Status", "Crawler Message", "Crawl Priority",
]

def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scopes
    )

    return gspread.authorize(creds)

def ensure_worksheet(spreadsheet, name, rows=2000, cols=30):
    try:
        return spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=name, rows=rows, cols=cols)

def existing_tender_keys(sheet):
    values = sheet.get_all_values()

    if not values:
        return set()

    header = values[0]

    try:
        idx = header.index("Tender Key")
    except ValueError:
        return set()

    return {
        r[idx]
        for r in values[1:]
        if len(r) > idx and r[idx]
    }

def sync_portal_directory(sheet, portals, health):
    rows = [PORTAL_HEADERS]

    for p in portals:
        h = health.get(p.portal)

        rows.append([
            p.portal,
            p.category,
            p.state,
            p.url,
            p.active,
            h.checked_at if h else "",
            h.discovered_count if h else 0,
            h.status if h else "NOT_RUN",
            h.message if h else "",
            p.priority,
        ])

    sheet.clear()
    sheet.update(values=rows, range_name=f"A1:J{len(rows)}")

def process_candidate(candidate, docman):
    doc_urls = docman.discover_from_detail_page(candidate)
    docs = [
        docman.download_and_extract(url, candidate.stable_key())
        for url in doc_urls[:20]
    ]

    evidence = extract_eligibility(candidate, docs)
    decision = evaluate_qualification(evidence)

    return evidence, docs, decision

def run_pipeline():
    started = time.time()
    log.info("%s started", APP_NAME)

    portals = load_master_portals()
    crawler = TenderCrawler()
    all_candidates = []
    health = {}

    non_gem = [p for p in portals if "gem.gov.in" not in p.url.lower()]

    with ThreadPoolExecutor(max_workers=CRAWL_WORKERS) as executor:
        futures = {executor.submit(crawler.crawl, p): p for p in non_gem}

        for future in as_completed(futures):
            p = futures[future]

            try:
                items, h = future.result()
            except Exception as e:
                items = []
                h = CrawlHealth(
                    p.portal,
                    p.url,
                    "PARSER_FAILED",
                    0,
                    str(e)[:250]
                )

            health[p.portal] = h
            all_candidates.extend(items)

    gem_items, gem_health = crawl_gem()
    all_candidates.extend(gem_items)
    health["Government e-Marketplace (GeM)"] = gem_health

    candidates = list({
        c.stable_key(): c
        for c in all_candidates
    }.values())

    gc = get_gspread_client()
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    active_sheet = ensure_worksheet(
        spreadsheet,
        "Active_Tenders",
        rows=5000,
        cols=30
    )

    portal_sheet = ensure_worksheet(
        spreadsheet,
        "Portal_Directory",
        rows=max(500, len(portals) + 50),
        cols=12
    )

    active_sheet.update(values=[ACTIVE_HEADERS], range_name="A1:Y1")
    existing = existing_tender_keys(active_sheet)

    docman = DocumentManager()
    new_rows = []

    for c in candidates:
        key = c.stable_key()

        if key in existing:
            continue

        evidence, docs, decision = process_candidate(c, docman)

        downloaded_files = [d.local_path for d in docs if d.local_path]
        doc_urls = list(dict.fromkeys(
            c.discovered_doc_urls + [d.url for d in docs if d.url]
        ))

        exp_parts = []

        if evidence.get("single_similar_work_required_inr") is not None:
            exp_parts.append(
                f"1 work >= ₹{evidence['single_similar_work_required_inr']:,.0f}"
            )

        if evidence.get("two_similar_works_each_required_inr") is not None:
            exp_parts.append(
                f"2 works each >= ₹{evidence['two_similar_works_each_required_inr']:,.0f}"
            )

        if evidence.get("three_similar_works_each_required_inr") is not None:
            exp_parts.append(
                f"3 works each >= ₹{evidence['three_similar_works_each_required_inr']:,.0f}"
            )

        experience = "; ".join(exp_parts) or "NOT VERIFIED / NOT STATED"

        key_compliance = []

        if evidence.get("physical_submission_required") is True:
            key_compliance.append("Physical submission required")

        if evidence.get("qcbs_or_technical_pitch") is True:
            key_compliance.append("QCBS / technical pitch")

        if evidence.get("named_celebrity_or_artist_mandate") is True:
            key_compliance.append("Artist / celebrity mandate")

        remarks = f"{decision.reason} Action: {decision.action_plan}"

        new_rows.append([
            dt.datetime.now(tz=IST).strftime("%Y-%m-%d"),
            key,
            evidence.get("tender_reference") or c.tender_id or "NOT VERIFIED",
            c.portal,
            c.state,
            evidence.get("organisation") or c.organization or "NOT VERIFIED",
            evidence.get("scope_summary") or c.title,
            c.category,
            evidence.get("estimated_value_inr")
                if evidence.get("estimated_value_inr") is not None
                else "NOT VERIFIED / NOT STATED",
            evidence.get("emd_inr")
                if evidence.get("emd_inr") is not None
                else "NOT VERIFIED / NOT STATED",
            evidence.get("emd_exemption_text") or "NOT VERIFIED / NOT STATED",
            evidence.get("submission_deadline") or c.deadline_raw or "NOT VERIFIED",
            evidence.get("average_turnover_required_inr")
                if evidence.get("average_turnover_required_inr") is not None
                else "NOT VERIFIED / NOT STATED",
            experience,
            evidence.get("entity_type_restriction") or "NOT VERIFIED / NOT STATED",
            "Yes" if evidence.get("consortium_or_jv_allowed") is True
                else "No" if evidence.get("consortium_or_jv_allowed") is False
                else "NOT VERIFIED / NOT STATED",
            "; ".join(key_compliance) or "No special flag extracted",
            "\n".join(doc_urls),
            "\n".join(downloaded_files),
            evidence.get("pre_bid_date") or "NOT VERIFIED / NOT STATED",
            c.detail_url or c.source_url,
            decision.status,
            remarks,
            "\n".join((evidence.get("evidence_quotes") or [])[:10]),
            evidence.get("extraction_status", "UNKNOWN"),
        ])

    if new_rows:
        active_sheet.append_rows(new_rows, value_input_option="RAW")

    sync_portal_directory(portal_sheet, portals, health)

    log.info("Completed in %.1f sec", time.time() - started)

if __name__ == "__main__":
    run_pipeline()
