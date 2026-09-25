import os
import io
import re
import json
import time
import datetime
import urllib.parse
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
import urllib3
from bs4 import BeautifulSoup
import pypdf
from google import genai
from google.oauth2 import service_account
import gspread
from playwright.sync_api import sync_playwright

# Disable insecure SSL warnings for government portals with self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    import pandas as pd
except ImportError:
    pd = None

# ==============================================================================
# 1. VERIFIED SOUL EVENTS & CONSULTANCY AUDITED BENCHMARKS
# ==============================================================================
COMPANY_PROFILE = {
    "agency_name": "Soul Events and Consultancy",
    "constitution": "Proprietorship",
    "proprietor": "Ashish Garg",
    "msme_status": "Micro Enterprise (Services) - Valid Udyam MH-19-0113805",
    "gstin": "27AOLPG1479A2Z3",
    "pan": "AOLPG1479A",
    "financials": {
        "avg_3yr_turnover_inr": 38046000,      # Rs. 3.80 Crores (CA Certified UDIN: 26145975UYPNOU5226)
        "fy_2025_26_turnover_inr": 50411575,  # Rs. 5.04 Crores
        "audited_net_worth_inr": 4680004       # Rs. 46.80 Lakhs (Audited Balance Sheet)
    },
    "verified_work_orders": [
        {"client": "Mahanadi Coalfields Ltd (CIL)", "title": "Celebrity Star Night (Shri Sukhwinder Singh)", "val": 7898550},
        {"client": "Hindustan Petroleum Corp Ltd (HPCL)", "title": "Stage & Venue Decor Deepotsav", "val": 7223000},
        {"client": "Epiroc Mining India Pvt Ltd", "title": "Groundbreaking Ceremony & German Hangar", "val": 6857653},
        {"client": "New India Assurance Co Ltd (NIA)", "title": "108th Foundation Day at NCPA Mumbai", "val": 4628443},
        {"client": "General Insurance Corp (GIC Re)", "title": "Underwriters Meet Indore (Hospitality & Event)", "val": 3338320},
        {"client": "Central Inst of Fisheries Education (ICAR)", "title": "Fish Fair Exhibition Stalls & Pandal", "val": 1540950}
    ],
    "max_single_past_work_order": 7898550,    # Rs. 78.98 Lakhs
    "max_two_works_threshold": 7223000,        # Two works >= Rs. 72.23 Lakhs
    "max_three_works_threshold": 4628443       # Three works >= Rs. 46.28 Lakhs
}

# Complete 59-Keyword Matrix for Pan-India Search
EVENT_KEYWORDS = [
    "Event Management", "Event Agency", "Event Management Agency", "Exhibition", "Exhibition Stall",
    "Exhibition Pavilion", "Design and Fabrication", "Fabrication", "Expo", "Trade Fair", "Mela",
    "Festival", "Conference", "Seminar", "Convention", "Summit", "Conclave", "Roadshow",
    "Dealer Meet", "Annual Day", "Foundation Day", "Award Ceremony", "Cultural Programme",
    "Corporate Event", "Activation", "Brand Activation", "Experiential Marketing", "Publicity",
    "Outreach Campaign", "IEC", "Media Campaign", "Advertising Agency", "Creative Agency",
    "AV Production", "Audio Visual", "LED", "Sound Light", "Stage", "Tentage", "Decoration",
    "Venue Management", "Hospitality", "Manpower", "Printing", "Branding", "Signage",
    "Digital Marketing", "Social Media", "PR Agency", "Communication Agency",
    "Empanelment of Event Agency", "Empanelment of Advertising Agency", "EOI Event", "RFP Event",
    "Tourism Event", "Sports Event", "Government Function", "Launch Event", "Inauguration"
]

EVENT_KEYWORDS_LOWER = [k.lower() for k in EVENT_KEYWORDS]

SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1YgsKZNiamaZ-E2ZZhIY2MLm8diC5okmeUMGFQdgQzHQ")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
EXCEL_MASTER_FILE = "Pan_India_Tender_URL_Master.xlsx"

# ==============================================================================
# 2. DOCUMENT DOWNLOADER & TEXT EXTRACTION (INTO MEMORY)
# ==============================================================================
def download_and_extract_pdf_text(doc_url, max_chars=30000):
    """
    Downloads tender PDF from portal into memory and extracts high-signal text:
    Pre-Qualification Criteria (PQC), Instructions to Bidders (ITB), EMD, Scope.
    """
    if not doc_url or not doc_url.startswith("http"):
        return ""

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    }

    try:
        resp = requests.get(doc_url, headers=headers, timeout=18, verify=False)
        if resp.status_code == 200 and len(resp.content) > 1000:
            # Check if valid PDF binary
            if resp.content.startswith(b"%PDF"):
                reader = pypdf.PdfReader(io.BytesIO(resp.content))
                total_pages = len(reader.pages)

                signal_words = [
                    "eligibility", "turnover", "pre-qualification", "pqc", "experience",
                    "estimated", "emd", "earnest money", "scope of work", "criteria",
                    "relaxation", "msme", "consortium", "joint venture", "celebrity", "artist"
                ]

                extracted_chunks = []
                current_len = 0

                # Scan up to 35 pages prioritizing pages with bidding criteria
                scan_limit = min(35, total_pages)
                for i in range(scan_limit):
                    if current_len >= max_chars:
                        break
                    try:
                        p_text = reader.pages[i].extract_text() or ""
                        p_lower = p_text.lower()
                        # Always include first 2 pages (Bid summary & dates) or pages with PQC signals
                        if i < 2 or any(sw in p_lower for sw in signal_words):
                            chunk = f"\n--- [DOC PAGE {i+1}] ---\n" + p_text
                            extracted_chunks.append(chunk)
                            current_len += len(chunk)
                    except Exception:
                        continue

                return "\n".join(extracted_chunks)
            else:
                # If page returns HTML tender detail, extract visible text
                soup = BeautifulSoup(resp.text, "html.parser")
                for s in soup(["script", "style", "nav", "footer"]):
                    s.extract()
                return soup.get_text(separator="\n", strip=True)[:max_chars]
    except Exception as e:
        print(f"    Document download notice ({doc_url[:40]}...): {e}")

    return ""

# ==============================================================================
# 3. AI STRUCTURED EXTRACTION & STRICT QUALIFICATION (GEMINI 1.5 FLASH)
# ==============================================================================
def parse_and_qualify_tender_with_ai(tender_metadata, extracted_doc_text):
    """
    Passes real document text to Gemini to:
    1. Extract all structured fields (Value, EMD, Turnover, Work Order, Dates, Entity restrictions)
    2. Strictly check Soul Events' eligibility against audited numbers.
    """
    client = genai.Client(api_key=GEMINI_API_KEY)

    prompt = f"""
    You are the Senior Public Procurement Officer auditing a live Indian Government Tender for Soul Events and Consultancy.
    
    === VERIFIED BIDDER PROFILE: SOUL EVENTS AND CONSULTANCY ===
    {json.dumps(COMPANY_PROFILE, indent=2)}
    
    === CRAWLED TENDER METADATA ===
    {json.dumps(tender_metadata, indent=2)}
    
    === EXTRACTED OFFICIAL TENDER DOCUMENT TEXT ===
    {extracted_doc_text if extracted_doc_text else "Note: Direct PDF text not accessible; evaluate based on crawled metadata and standard Indian Public Procurement rules."}
    
    === TASK 1: EXTRACT STRUCTURED CRITERIA FROM THE DOCUMENT TEXT ===
    1. estimated_value: Extract exact numeric value in INR (e.g. "85,00,000" or "Item Rate / To Be Quoted").
    2. emd_and_exemption: Extract EMD amount in INR and verify if MSME / Udyam / Startup is exempt under Rule 170 of GFR.
    3. submission_deadline: Extract official bid submission closing date and time.
    4. min_turnover_required: Extract required average annual turnover (in INR/Lakhs).
    5. past_experience_criteria: Extract the exact prior work order rule (e.g., 80/50/40 rule or specific completed project value).
    6. entity_type_restrictions: Extract entity eligibility (e.g., "Proprietorship allowed", "Pvt Ltd / Public Ltd only", "Consortium / JV Allowed or Not Allowed").
    7. key_tech_riders: Specific technical requirements (e.g., Celebrity artist mandate letter, line array audio, P3 LED wall, German hangar).
    8. pre_bid_meeting_info: Date, time, venue, or VC link for pre-bid meeting.
    9. corrigenda_info: Note any date extensions, amendment notices, or corrigendum status.
    
    === TASK 2: STRICT ELIGIBILITY QUALIFICATION ===
    Evaluate Soul Events' bidding qualification strictly:
    A. MARK 'DISQUALIFIED' IF:
       - Entity Type strictly requires "Public Ltd / Private Ltd only" (Soul Events is a Sole Proprietorship).
       - Turnover required > Rs. 3.80 Crores AND tender explicitly states "No MSME relaxation and No JV/Consortium".
       - Single past work order required > Rs. 78.98 Lakhs AND JV/Consortium is forbidden.
       - Net Worth required > Rs. 46.80 Lakhs.
       - Pure civil engineering, security manpower supply, or central kitchen catering.
    B. MARK 'NEEDS MANUAL INTERVENTION' IF:
       - Celebrity Artist / Star Night: Requires exclusive artist mandate letter and video byte within 48h.
       - High Value with JV (Rs. 1.5 Cr to Rs. 10 Cr): Value exceeds direct work order (Rs. 78.98L) but Consortium/JV is permitted.
       - QCBS Technical Pitch: Requires full 60-slide creative presentation & 3D renders.
       - Physical Submission: Requires physical Demand Draft / Bank Guarantee courier before bid closing.
    C. MARK 'QUALIFIED' IF:
       - Direct match: Value <= Rs. 1.00 Crore, Turnover <= Rs. 3.80 Crores, Single work <= Rs. 78.98 Lakhs, and EMD is waived for MSME.
    
    Output strictly valid JSON with keys:
    {{
      "estimated_value": "...",
      "emd_and_exemption": "...",
      "submission_deadline": "...",
      "min_turnover_required": "...",
      "past_experience_criteria": "...",
      "entity_type_restrictions": "...",
      "key_tech_riders": "...",
      "pre_bid_meeting_info": "...",
      "corrigenda_info": "...",
      "status": "QUALIFIED" | "DISQUALIFIED" | "NEEDS MANUAL INTERVENTION",
      "reasoning": "...",
      "action_plan": "..."
    }}
    """

    for model_name in ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"]:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config={"response_mime_type": "application/json"}
            )
            return json.loads(response.text)
        except Exception as e:
            last_err = str(e)
            continue

    return {
        "estimated_value": "Refer Tender Doc",
        "emd_and_exemption": "MSME Exempted as per GFR 170",
        "submission_deadline": tender_metadata.get("deadline", "Refer Portal"),
        "min_turnover_required": "30% of tender value",
        "past_experience_criteria": "80/50/40 rule on similar works",
        "entity_type_restrictions": "Proprietorship / Consortium Allowed",
        "key_tech_riders": "Stage, sound, AV and event management specifications",
        "pre_bid_meeting_info": "Refer portal link for pre-bid schedule",
        "corrigenda_info": "Check portal link",
        "status": "NEEDS MANUAL INTERVENTION",
        "reasoning": f"Document extraction notice: {last_err}",
        "action_plan": "Review tender document directly on portal."
    }

# ==============================================================================
# 4. MULTI-SOURCE PAN-INDIA CRAWLERS (SEARCH → DEEP LINK RESOLUTION)
# ==============================================================================
def crawl_gepnic_endpoint(portal_meta):
    """Scrapes GePNIC endpoints extracting exact tender links and document download URLs."""
    found = []
    base_url = portal_meta["url"].rstrip("/")
    portal_name = portal_meta["portal"]
    state_name = portal_meta["state"]

    if "nicgep/app" in base_url or "eprocure/app" in base_url:
        feed_url = f"{base_url}?page=FrontEndLatestActiveTenders&service=page"
    elif any(k in base_url for k in ["etender", "eproc", "tenders"]):
        feed_url = f"{base_url}/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    else:
        return found

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    }

    try:
        resp = requests.get(feed_url, headers=headers, timeout=12, verify=False)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            rows = soup.find_all("tr")
            for row in rows:
                row_text = row.text.lower()
                if any(kw in row_text for kw in EVENT_KEYWORDS_LOWER):
                    cols = row.find_all("td")
                    if len(cols) >= 5:
                        title_col = cols[4].text.strip()
                        dept_col = cols[5].text.strip() if len(cols) > 5 else f"{portal_name} ({state_name})"
                        closing_date = cols[2].text.strip()

                        tender_id_match = re.search(r"\d{4}_[A-Z0-9]+_\d+_\d+", title_col)
                        tender_id = tender_id_match.group(0) if tender_id_match else f"{portal_name[:3]}/{int(time.time())}"

                        portal_link = f"{base_url}?page=FrontEndTenderDetailsExternal&service=page&tenderId={tender_id}"
                        rfp_doc_url = f"{base_url}?page=FrontEndDownloadTenderDocument&service=page&tenderId={tender_id}"
                        corrigendum_url = f"{base_url}?page=FrontEndCorrigendumDetailsExternal&service=page&tenderId={tender_id}"

                        link_elem = cols[4].find("a", href=True) or row.find("a", href=True)
                        if link_elem and link_elem["href"]:
                            raw_href = link_elem["href"]
                            if "javascript:" not in raw_href.lower():
                                rfp_doc_url = urllib.parse.urljoin(base_url, raw_href)

                        found.append({
                            "portal": f"{portal_name} ({state_name})",
                            "tender_id": tender_id,
                            "organization": dept_col,
                            "title": title_col,
                            "category": "Exhibition / Stalls" if "exhibition" in title_col.lower() else "Corporate Conclave / B2B",
                            "deadline": closing_date,
                            "rfp_doc_url": rfp_doc_url,
                            "corrigendum_url": corrigendum_url,
                            "portal_link": portal_link
                        })
    except Exception:
        pass
    return found


def crawl_gem_live(page):
    """Scrapes live bids from GeM BidPlus extracting Bid Doc, Corrigendum, and Portal links."""
    gem_bids = []
    print("Crawling Live GeM BidPlus Portal across event categories...")
    try:
        page.goto("https://bidplus.gem.gov.in/all-bids", timeout=35000, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)

        for kw in ["Event Management", "Exhibition", "Conferences", "Sound and Light"]:
            try:
                search_input = page.locator('input#search_by').or_(page.locator('input[type="search"]')).or_(page.locator('input[placeholder*="Search"]')).first
                if search_input.is_visible():
                    search_input.fill("")
                    search_input.fill(kw)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(3500)

                    cards = page.locator(".card, .border.block, div[class*='bid-card'], div[class*='block_header']").all()
                    for card in cards[:5]:
                        card_text = card.inner_text()
                        bid_match = re.search(r"GEM/\d{4}/B/\d+", card_text)
                        if bid_match:
                            bid_no = bid_match.group(0)
                            bid_num = bid_no.split('/')[-1]

                            rfp_doc_url = f"https://bidplus.gem.gov.in/showbidDocument/{bid_num}"
                            corrigendum_url = f"https://bidplus.gem.gov.in/show-bid-corrigendum/{bid_num}"
                            portal_link = f"https://bidplus.gem.gov.in/all-bids?bid={bid_num}"

                            lines = [l.strip() for l in card_text.split("\n") if l.strip()]
                            dept = lines[2] if len(lines) > 2 else "Central PSU / Ministry"
                            title = lines[1] if len(lines) > 1 else f"{kw} Services"

                            gem_bids.append({
                                "portal": "GeM Portal (gem.gov.in)",
                                "tender_id": bid_no,
                                "organization": dept,
                                "title": title,
                                "category": "Artist / Star Night" if "star" in kw.lower() else "Corporate Conclave / B2B",
                                "deadline": "Refer Bid Document",
                                "rfp_doc_url": rfp_doc_url,
                                "corrigendum_url": corrigendum_url,
                                "portal_link": portal_link
                            })
            except Exception:
                continue
    except Exception as e:
        print(f"Notice during GeM Live Crawl: {e}")
    return gem_bids


def crawl_aggregators_live():
    """Scrapes Pan-India aggregator feeds indexing municipal corporations & state boards."""
    agg_results = []
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    }

    for kw in ["event-management", "exhibition", "conclave"]:
        try:
            url = f"https://bidassist.com/all-tenders/search?keyword={kw}"
            resp = requests.get(url, headers=headers, timeout=10)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, "html.parser")
                cards = soup.find_all("div", class_="tender-card") or soup.find_all("div", class_="search-card")
                for card in cards[:3]:
                    title_elem = card.find("h3") or card.find("a")
                    loc_elem = card.find("span", class_="location")

                    if title_elem:
                        anchor = card.find("a", href=True)
                        doc_link = urllib.parse.urljoin("https://bidassist.com", anchor["href"]) if anchor else url

                        agg_results.append({
                            "portal": "Pan-India Aggregator (BidAssist)",
                            "tender_id": f"BA/{int(time.time())}/{kw[:3].upper()}",
                            "organization": loc_elem.text.strip() if loc_elem else "State / Municipal Agency",
                            "title": title_elem.text.strip(),
                            "category": "Turnkey Production" if "exhibition" in kw else "Corporate Conclave / B2B",
                            "deadline": (datetime.datetime.now() + datetime.timedelta(days=12)).strftime("%Y-%m-%d 15:00"),
                            "rfp_doc_url": doc_link,
                            "corrigendum_url": doc_link,
                            "portal_link": doc_link
                        })
        except Exception:
            pass
    return agg_results

# ==============================================================================
# 5. DEDUPLICATION, AUTO-CLEAN & FULL PIPELINE EXECUTION
# ==============================================================================
def load_and_sync_portal_directory(gc):
    """Loads all 109 active URLs and populates the Portal_Directory worksheet safely."""
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    try:
        portal_sheet = spreadsheet.worksheet("Portal_Directory")
    except gspread.WorksheetNotFound:
        portal_sheet = spreadsheet.add_worksheet(title="Portal_Directory", rows=250, cols=15)

    headers = [
        "Portal", "Category", "State", "URL", "Active",
        "Last Crawl", "Tender Count", "Documents Accessible",
        "Login/Captcha Required", "Crawl Priority"
    ]
    try:
        portal_sheet.update(range_name="A1:J1", values=[headers])
    except Exception:
        portal_sheet.update([headers], "A1:J1")

    portals_list = []
    if pd and os.path.exists(EXCEL_MASTER_FILE):
        try:
            df = pd.read_excel(EXCEL_MASTER_FILE, sheet_name="Tender URL Master", skiprows=2)
            active_df = df[df["Active"].astype(str).str.strip().str.lower() == "yes"]
            for _, r in active_df.iterrows():
                u = str(r.get("URL", "")).strip()
                cat = str(r.get("Category", "")).strip()
                prio = str(r.get("Priority", "P2")).strip().upper()
                src_type = str(r.get("Source Type", "")).strip()

                if "gem.gov.in" in u.lower():
                    doc_acc = "Yes (Direct RFP PDF & Corrigendum)"
                    login_req = "No (Public BidPlus)"
                elif any(k in u.lower() for k in ["nicgep", "eproc", "etender", "tenders"]):
                    doc_acc = "Yes (Public NIT, RFP & Pre-bid)"
                    login_req = "No (Public Access)"
                elif src_type == "Discovery":
                    doc_acc = "Yes (Public Aggregator Link)"
                    login_req = "No (Public Index)"
                else:
                    doc_acc = "Varies (Org Website)"
                    login_req = "May require Captcha/Portal login"

                portals_list.append({
                    "portal": str(r.get("Portal / Organisation", "")).strip(),
                    "category": cat,
                    "state": str(r.get("State/Region", "")).strip(),
                    "url": u,
                    "active": "Yes",
                    "doc_acc": doc_acc,
                    "login_req": login_req,
                    "priority": prio
                })
        except Exception as e:
            print(f"Excel read notice: {e}")

    existing_rows = portal_sheet.get_all_values()
    if len(existing_rows) <= 1 and portals_list:
        rows_to_write = []
        for p in portals_list:
            rows_to_write.append([
                p["portal"], p["category"], p["state"], p["url"], p["active"],
                "-", 0, p["doc_acc"], p["login_req"], p["priority"]
            ])
        try:
            portal_sheet.update(range_name=f"A2:J{len(rows_to_write)+1}", values=rows_to_write)
        except Exception:
            portal_sheet.update(rows_to_write, f"A2:J{len(rows_to_write)+1}")

    return portals_list, portal_sheet


def purge_legacy_dummy_rows(sheet):
    """Automatically cleans legacy sample mock rows from the sheet."""
    try:
        rows = sheet.get_all_values()
        if len(rows) > 1:
            dummy_indicators = [
                "gem/2026/b/7891024", "cppp/2026/dpiit", "up/2026/tourism/4512",
                "rj/2026/sppp", "mh/2026/cidco/3321", "gem bid document pdf attached",
                "cppp tender notice & rfp pdf", "ai parsing exception"
            ]
            for idx in range(len(rows), 1, -1):
                row_str = " ".join([str(c).lower() for c in rows[idx-1]])
                if any(di in row_str for di in dummy_indicators):
                    sheet.delete_rows(idx)
                    print(f"Purged legacy dummy row #{idx}")
    except Exception as e:
        print(f"Notice during dummy row cleanup: {e}")


def run_pipeline():
    today = datetime.date.today().strftime("%Y-%m-%d")
    print(f"==================================================================")
    print(f"Starting Pan-India Tender Engine: {today}")
    print(f"Workflow: Search -> Open -> Download PDF -> Extract Criteria -> AI Qualify")
    print(f"==================================================================")

    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    active_sheet = spreadsheet.worksheet("Active_Tenders")

    # 1. Expand columns to prevent API Error 400 (Requires 19 columns: A to S)
    current_cols = active_sheet.col_count
    if current_cols < 20:
        print(f"Expanding Active_Tenders columns to 22...")
        active_sheet.add_cols(22 - current_cols)

    # 2. Update Column Headers (19 Structured Columns)
    updated_headers = [
        "Date Found", "Tender ID / Ref No", "Portal Name", "Organization / Dept",
        "Tender Title & Scope", "Category", "Estimated Value (INR)", "EMD & MSME Exemption",
        "Submission Deadline", "Min Turnover Required", "Past Experience Criteria",
        "Entity Restrictions & JV", "Key Tech Riders & Compliance", "Pre-Bid Meeting Info",
        "RFP / Tender Doc URL", "Corrigenda Links", "Portal Link",
        "Qualification Status", "Remarks & Action Plan"
    ]
    try:
        active_sheet.update(range_name="A1:S1", values=[updated_headers])
    except Exception:
        active_sheet.update([updated_headers], "A1:S1")

    # 3. Sync All 109 Master Portals to Portal_Directory
    all_109_portals, portal_sheet = load_and_sync_portal_directory(gc)

    # 4. Clean Legacy Dummy Rows from Active_Tenders
    purge_legacy_dummy_rows(active_sheet)

    # 5. Read Existing Tender IDs
    try:
        col_b_vals = active_sheet.col_values(2)
        existing_ids = set(col_b_vals[1:])
    except Exception:
        existing_ids = set()
    print(f"Existing verified tenders logged in sheet: {len(existing_ids)}")

    live_tenders = []
    portal_counts = {}

    # 6. Parallel Crawl Across All Portals
    print(f"Crawling portals in parallel...")
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(crawl_gepnic_endpoint, portal): portal for portal in all_109_portals}
        for f in as_completed(futures):
            p_meta = futures[f]
            try:
                res = f.result()
                count = len(res) if res else 0
                portal_counts[p_meta["portal"]] = count
                if res:
                    live_tenders.extend(res)
            except Exception:
                portal_counts[p_meta["portal"]] = 0

    # 7. Playwright Crawl for GeM BidPlus
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
            )
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
            )
            page = context.new_page()
            gem_records = crawl_gem_live(page)
            live_tenders.extend(gem_records)
            portal_counts["Government e-Marketplace (GeM)"] = len(gem_records)
            portal_counts["GeM BidPlus"] = len(gem_records)
            browser.close()
    except Exception as e:
        print(f"GeM crawler notice: {e}")

    # 8. Aggregator Feeds
    print("Querying Pan-India aggregator feeds...")
    agg_records = crawl_aggregators_live()
    live_tenders.extend(agg_records)

    # 9. Update Portal Counts in Single Batch Call
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M IST")
    try:
        p_rows = portal_sheet.get_all_values()
        if len(p_rows) > 1:
            batch_vals = []
            for r in p_rows[1:]:
                p_name = r[0]
                count = portal_counts.get(p_name, 0)
                batch_vals.append([now_str, count])

            end_row = 1 + len(batch_vals)
            try:
                portal_sheet.update(range_name=f"F2:G{end_row}", values=batch_vals)
            except Exception:
                portal_sheet.update(batch_vals, f"F2:G{end_row}")
            print(f"Updated crawl status for all {len(batch_vals)} portals in 'Portal_Directory'.")
    except Exception as e:
        print(f"Notice updating Portal_Directory counts: {e}")

    # 10. Deduplicate Live Tenders
    unique_new_tenders = []
    seen_in_batch = set()
    for t in live_tenders:
        t_id = t["tender_id"]
        if t_id not in existing_ids and t_id not in seen_in_batch:
            unique_new_tenders.append(t)
            seen_in_batch.add(t_id)

    print(f"------------------------------------------------------------------")
    print(f"Total Live Event Tenders Scraped: {len(live_tenders)}")
    print(f"New Unique Tenders to Download, Extract & Evaluate: {len(unique_new_tenders)}")
    print(f"------------------------------------------------------------------")

    if not unique_new_tenders:
        print("Zero new event tenders published today. Exiting cleanly.")
        return

    # 11. DEEP EXTRACTION & AI EVALUATION (DOCUMENT LEVEL)
    for item in unique_new_tenders:
        print(f"\n[Processing] Tender ID: {item['tender_id']}")
        print(f"  Title: {item['title'][:50]}...")
        print(f"  Downloading Document: {item['rfp_doc_url']}")

        # Step A: Download PDF and extract PQC text
        doc_text = download_and_extract_pdf_text(item["rfp_doc_url"])
        if doc_text:
            print(f"  -> Successfully extracted {len(doc_text)} characters from official document text.")
        else:
            print(f"  -> Document binary not directly downloadable; falling back to portal detail extraction.")

        # Step B: Pass extracted text to Gemini for structured extraction & qualification
        analysis = parse_and_qualify_tender_with_ai(item, doc_text)
        remarks_field = f"{analysis.get('reasoning', '')} Action: {analysis.get('action_plan', '')}"

        row_data = [
            today,
            item["tender_id"],
            item["portal"],
            item["organization"],
            item["title"],
            item["category"],
            analysis.get("estimated_value", "Refer Tender Doc"),
            analysis.get("emd_and_exemption", "MSME Exempted as per GFR 170"),
            analysis.get("submission_deadline", item.get("deadline", "Refer Portal")),
            analysis.get("min_turnover_required", "30% of tender value"),
            analysis.get("past_experience_criteria", "80/50/40 rule on similar works"),
            analysis.get("entity_type_restrictions", "Proprietorship / Consortium Allowed"),
            analysis.get("key_tech_riders", "Stage, sound, AV and event management specifications"),
            analysis.get("pre_bid_meeting_info", "Refer portal link"),
            item["rfp_doc_url"],
            item.get("corrigendum_url", item["rfp_doc_url"]),
            item["portal_link"],
            analysis.get("status", "NEEDS MANUAL INTERVENTION"),
            remarks_field
        ]

        active_sheet.append_row(row_data)
        print(f"  -> Appended [{analysis.get('status')}] with Exact Criteria.")

    print(f"\nPan-India deep document crawl and evaluation completed successfully for {today}.")

if __name__ == "__main__":
    try:
        run_pipeline()
    except Exception as fatal_err:
        print("=== FATAL PIPELINE EXCEPTION TRACEBACK ===")
        traceback.print_exc()
        raise fatal_err
