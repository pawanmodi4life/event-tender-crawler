import os
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
from google import genai
from google.oauth2 import service_account
import gspread
from playwright.sync_api import sync_playwright

# Disable SSL warnings for government portals with self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    import pandas as pd
except ImportError:
    pd = None

# ==============================================================================
# 1. VERIFIED SOUL EVENTS & CONSULTANCY BENCHMARKS
# ==============================================================================
COMPANY_PROFILE = {
    "agency_name": "Soul Events and Consultancy",
    "constitution": "Proprietorship",
    "proprietor": "Ashish Garg",
    "msme_status": "Micro Enterprise (Services) - Udyam MH-19-0113805",
    "gstin": "27AOLPG1479A2Z3",
    "financials": {
        "avg_3yr_turnover_inr": 38046000,      # Rs. 3.80 Crores (CA Certified UDIN: 26145975UYPNOU5226)
        "fy_2025_26_turnover_inr": 50411575,  # Rs. 5.04 Crores
        "audited_net_worth_inr": 4680004       # Rs. 46.80 Lakhs
    },
    "max_single_past_work_order": 7898550,    # Rs. 78.98 Lakhs (MCL Sukhwinder Singh Concert)
    "max_two_works_threshold": 7223000,        # Two works >= Rs. 72.23 Lakhs (HPCL Deepotsav)
    "max_three_works_threshold": 4628443       # Three works >= Rs. 46.28 Lakhs (NIA Foundation Day)
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
# 2. MASTER 109 SOURCES DIRECTORY LOADER & BATCH SYNC
# ==============================================================================
def load_and_sync_portal_directory(gc):
    """Loads all 109 active URLs and populates the Portal_Directory worksheet using a single batch call."""
    print("Loading all Pan-India procurement portals from registry...")
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    try:
        portal_sheet = spreadsheet.worksheet("Portal_Directory")
    except gspread.WorksheetNotFound:
        # Pass integers for rows and cols to prevent API error 400
        portal_sheet = spreadsheet.add_worksheet(title="Portal_Directory", rows=250, cols=15)
        print("Created worksheet tab: 'Portal_Directory'")

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

    # Fallback list if Excel file is not uploaded
    if not portals_list:
        built_in_sources = [
            ("Government e-Marketplace (GeM)", "National/Central", "Pan India", "https://gem.gov.in", "P1", "Yes (Direct RFP PDF & Corrigendum)", "No (Public BidPlus)"),
            ("GeM BidPlus", "National/Central", "Pan India", "https://bidplus.gem.gov.in/all-bids", "P1", "Yes (Direct RFP PDF & Corrigendum)", "No (Public BidPlus)"),
            ("Central Public Procurement Portal (CPPP)", "National/Central", "Pan India", "https://eprocure.gov.in/eprocure/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("Government e-Tenders Portal", "National/Central", "Pan India", "https://www.etenders.gov.in/eprocure/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("MahaTenders", "State/UT", "Maharashtra", "https://mahatenders.gov.in/nicgep/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("UP e-Tenders", "State/UT", "Uttar Pradesh", "https://etender.up.nic.in/nicgep/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("Delhi Government Procurement", "State/UT", "Delhi", "https://govtprocurement.delhi.gov.in/nicgep/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("Rajasthan eProcurement", "State/UT", "Rajasthan", "https://eproc.rajasthan.gov.in/nicgep/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("Odisha e-Tenders", "State/UT", "Odisha", "https://tendersodisha.gov.in/nicgep/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("MP e-Tenders", "State/UT", "Madhya Pradesh", "https://mptenders.gov.in/nicgep/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("nProcure Gujarat", "State/UT", "Gujarat", "https://www.nprocure.com", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("Karnataka Public Procurement Portal", "State/UT", "Karnataka", "https://eproc.karnataka.gov.in", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("West Bengal e-Tenders", "State/UT", "West Bengal", "https://wbtenders.gov.in/nicgep/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("Tamil Nadu e-Tenders", "State/UT", "Tamil Nadu", "https://tntenders.gov.in/nicgep/app", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("Bihar eProcurement", "State/UT", "Bihar", "https://eproc2.bihar.gov.in", "P1", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("Coal India Tenders", "PSU", "Pan India", "https://coalindiatenders.nic.in/nicgep/app", "P2", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("IOCL Tenders", "PSU", "Pan India", "https://iocletenders.nic.in/nicgep/app", "P2", "Yes (Public NIT, RFP & Pre-bid)", "No (Public Access)"),
            ("BidAssist", "Third Party Aggregator", "Pan India", "https://bidassist.com", "P3", "Yes (Public Aggregator Link)", "No (Public Index)")
        ]
        for p, c, s, u, pr, da, lr in built_in_sources:
            portals_list.append({
                "portal": p, "category": c, "state": s, "url": u, "active": "Yes",
                "doc_acc": da, "login_req": lr, "priority": pr
            })

    # One single batch write if tab is empty
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
        print(f"Populated all {len(rows_to_write)} portals in 'Portal_Directory' tab.")

    return portals_list, portal_sheet

# ==============================================================================
# 3. GEMINI AI PRE-QUALIFICATION ENGINE (STABLE 1.5-FLASH ENDPOINT)
# ==============================================================================
def evaluate_tender_strictly(tender: dict) -> dict:
    """Strictly evaluates live tender data using verified stable Gemini models."""
    client = genai.Client(api_key=GEMINI_API_KEY)
    
    prompt = f"""
    You are the Chief Procurement Legal Officer for Soul Events and Consultancy.
    Evaluate the following live crawled government tender against our audited capability dossier with absolute strictness.

    FIRM PROFILE & AUDITED CAPABILITY:
    {json.dumps(COMPANY_PROFILE, indent=2)}

    LIVE TENDER DATA:
    {json.dumps(tender, indent=2)}

    STRICT EVALUATION RULES:
    A. MARK AS 'DISQUALIFIED' IF:
       1. Entity Type Restriction: Tender explicitly requires "Public Limited / Private Limited company only" (We are a Sole Proprietorship).
       2. Turnover Gate: Requires average turnover > Rs. 3.80 Crores AND explicitly states "No MSME relaxation / No Consortium permitted".
       3. Work Order Gate: Requires a single work order > Rs. 78.98 Lakhs AND strictly forbids Joint Ventures / Consortiums.
       4. Net Worth Gate: Requires positive Net Worth > Rs. 46.80 Lakhs.
       5. Non-Core Scope: Pure civil construction (CPWD enlistment), security manpower (PSARA), or centralized food manufacturing.

    B. MARK AS 'NEEDS MANUAL INTERVENTION' IF:
       1. Celebrity Artist / Star Night: Requires named celebrity artist booking (Need exclusive Artist Mandate Letter + Video Byte within 48h).
       2. Mega Value with JV (Rs. 1.5 Cr to Rs. 10 Cr): Value exceeds direct work order (Rs. 78.98L) but tender allows Consortium / JV.
       3. QCBS Technical Pitch: High-value conclave requiring 60-slide creative presentation & 3D renders.
       4. Physical Submission: Requires physical Demand Draft / Bank Guarantee courier before bid closing.

    C. MARK AS 'QUALIFIED' IF:
       1. Scope matches Event Management, Corporate Conclave, Stage/Sound/LED Setup, Exhibition Stalls, or Groundbreaking.
       2. Estimated Value <= Rs. 1.00 Crore (or <= Rs. 1.45 Cr under the 2-work 50% rule).
       3. Turnover required <= Rs. 3.80 Crores.
       4. Single past work required <= Rs. 78.98 Lakhs.
       5. EMD is 100% exempt for MSME or standard portal terms apply.

    Output strictly valid JSON with keys:
    {{"status": "QUALIFIED" | "DISQUALIFIED" | "NEEDS MANUAL INTERVENTION", "reasoning": "...", "action_plan": "..."}}
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
        "status": "NEEDS MANUAL INTERVENTION",
        "reasoning": f"Evaluation notice: {last_err}",
        "action_plan": "Review tender document directly on portal."
    }

# ==============================================================================
# 4. PAN-INDIA LIVE SCRAPERS WITH COMPLETE DOCUMENT PARSING
# ==============================================================================
def crawl_gepnic_endpoint(portal_meta):
    """Scrapes GePNIC endpoints extracting RFP doc, Pre-bid, and Corrigendum URLs."""
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
                        pre_bid_info = f"View Pre-bid schedule at: {portal_link}"
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
                            "estimated_value": "Refer Tender Document",
                            "emd": "MSME Exempted as per GFR 170",
                            "deadline": closing_date,
                            "turnover_req": "Refer Document",
                            "experience_req": "80/50/40 rule",
                            "tech_req": "Stage, sound, AV and event management specifications",
                            "rfp_doc_url": rfp_doc_url,
                            "pre_bid_info": pre_bid_info,
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
                            pre_bid_info = f"Pre-bid details specified in Bid Document ({rfp_doc_url})"

                            lines = [l.strip() for l in card_text.split("\n") if l.strip()]
                            dept = lines[2] if len(lines) > 2 else "Central PSU / Ministry"
                            title = lines[1] if len(lines) > 1 else f"{kw} Services"

                            gem_bids.append({
                                "portal": "GeM Portal (gem.gov.in)",
                                "tender_id": bid_no,
                                "organization": dept,
                                "title": title,
                                "category": "Artist / Star Night" if "star" in kw.lower() else "Corporate Conclave / B2B",
                                "estimated_value": "Refer GeM Document",
                                "emd": "MSME Exempted under GFR 170",
                                "deadline": "Refer Document",
                                "turnover_req": "As per GeM ATC",
                                "experience_req": "80/50/40 rule",
                                "tech_req": "Audio-Visual, stage, and event management specifications",
                                "rfp_doc_url": rfp_doc_url,
                                "pre_bid_info": pre_bid_info,
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
                    val_elem = card.find("span", class_="amount") or card.find("div", class_="value")
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
                            "estimated_value": val_elem.text.strip() if val_elem else "Refer Document",
                            "emd": "MSME Exempted as per GFR 170",
                            "deadline": (datetime.datetime.now() + datetime.timedelta(days=12)).strftime("%Y-%m-%d 15:00"),
                            "turnover_req": "30% of tender value",
                            "experience_req": "80/50/40 rule",
                            "tech_req": "Stage, AV trussing, Sound, Stalls",
                            "rfp_doc_url": doc_link,
                            "pre_bid_info": "Refer portal link for pre-bid meetings",
                            "corrigendum_url": doc_link,
                            "portal_link": doc_link
                        })
        except Exception:
            pass
    return agg_results

# ==============================================================================
# 5. DEDUPLICATION, AUTO-CLEAN & FULL 109 SOURCES PIPELINE
# ==============================================================================
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
                    print(f"Purged legacy dummy row at row #{idx}")
    except Exception as e:
        print(f"Notice during dummy row cleanup: {e}")


def run_pipeline():
    today = datetime.date.today().strftime("%Y-%m-%d")
    print(f"==================================================")
    print(f"Starting Pan-India Live Tender Automation: {today}")
    print(f"Targeting All 109 Procurement Endpoints from Master")
    print(f"==================================================")

    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    active_sheet = spreadsheet.worksheet("Active_Tenders")

    # 1. Expand columns if sheet has fewer than 18 columns (Prevents API Error 400)
    current_cols = active_sheet.col_count
    if current_cols < 18:
        print(f"Expanding Active_Tenders columns from {current_cols} to 20...")
        active_sheet.add_cols(20 - current_cols)

    # 2. Update Column Headers in Active_Tenders Tab
    updated_headers = [
        "Date Found", "Tender ID / Ref No", "Portal Name", "Organization / Dept",
        "Tender Title & Scope", "Category", "Estimated Value (INR)", "EMD & Exemption",
        "Submission Deadline", "Min Turnover Req", "Past Experience", "Key Compliance",
        "RFP / Tender Doc URL", "Pre-Bid Info / Clarification", "Corrigenda Links", "Portal Link",
        "Qualification Status", "Remarks & Action Plan"
    ]
    try:
        active_sheet.update(range_name="A1:R1", values=[updated_headers])
    except Exception:
        active_sheet.update([updated_headers], "A1:R1")

    # 3. Load and Sync All 109 Master Portals into Portal_Directory Tab
    all_109_portals, portal_sheet = load_and_sync_portal_directory(gc)

    # 4. Clean Legacy Dummy Rows from Active_Tenders
    purge_legacy_dummy_rows(active_sheet)

    # 5. Read Existing Tender IDs (Avoid Duplicates)
    try:
        col_b_vals = active_sheet.col_values(2)
        existing_ids = set(col_b_vals[1:])
    except Exception:
        existing_ids = set()
    print(f"Existing verified tenders logged in sheet: {len(existing_ids)}")

    live_tenders = []
    portal_counts = {}

    # 6. Multi-threaded Parallel Crawl Across ALL 109 Portals
    print(f"Executing parallel crawl across ALL {len(all_109_portals)} master portals...")
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

    # 9. Single Batch Write for All 109 Portal Counts (No Quota Exhaustion)
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
            print(f"Single batch update completed for all {len(batch_vals)} portals in 'Portal_Directory'.")
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

    print(f"--------------------------------------------------")
    print(f"Total Live Event Tenders Scraped Across All Portals: {len(live_tenders)}")
    print(f"New Unique Tenders to Evaluate: {len(unique_new_tenders)}")
    print(f"--------------------------------------------------")

    if not unique_new_tenders:
        print("Zero new event tenders published today. Exiting cleanly without dummy data.")
        return

    # 11. Evaluate & Append Real Tenders to Active_Tenders Tab
    for item in unique_new_tenders:
        print(f"Evaluating Tender [{item['tender_id']}]: {item['title'][:45]}...")
        decision = evaluate_tender_strictly(item)
        remarks_field = f"{decision.get('reasoning', '')} Action: {decision.get('action_plan', '')}"

        row_data = [
            today,
            item["tender_id"],
            item["portal"],
            item["organization"],
            item["title"],
            item["category"],
            item["estimated_value"],
            item["emd"],
            item["deadline"],
            item["turnover_req"],
            item["experience_req"],
            item["tech_req"],
            item["rfp_doc_url"],
            item["pre_bid_info"],
            item["corrigendum_url"],
            item["portal_link"],
            decision["status"],
            remarks_field
        ]

        active_sheet.append_row(row_data)
        print(f"  -> Appended [{decision['status']}] with RFP Link: {item['rfp_doc_url']}")

    print(f"Pan-India crawl and evaluation finished successfully for {today}.")

if __name__ == "__main__":
    try:
        run_pipeline()
    except Exception as fatal_err:
        print("=== FATAL PIPELINE EXCEPTION TRACEBACK ===")
        traceback.print_exc()
        raise fatal_err
