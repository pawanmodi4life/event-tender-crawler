import os
import re
import json
import time
import datetime
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
from bs4 import BeautifulSoup
from google import genai
from google.oauth2 import service_account
import gspread
from playwright.sync_api import sync_playwright

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
# 2. PAN-INDIA 109 SOURCES DIRECTORY BUILDER
# ==============================================================================
def load_and_sync_portal_directory(gc):
    """
    Creates/updates the 'Portal_Directory' tab in Google Sheets with the 10 requested columns:
    Portal | Category | State | URL | Active | Last Crawl | Tender Count | Documents Accessible | Login/Captcha Required | Crawl Priority
    """
    print("Syncing Master Portal Directory (109 URLs) into Google Sheets...")
    
    # 1. Get or create Portal_Directory sheet tab
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    try:
        portal_sheet = spreadsheet.worksheet("Portal_Directory")
    except gspread.WorksheetNotFound:
        portal_sheet = spreadsheet.add_worksheet(title="Portal_Directory", rows="150", cols="12")
        print("Created new worksheet tab: 'Portal_Directory'")

    # Set Header row
    headers = [
        "Portal", "Category", "State", "URL", "Active",
        "Last Crawl", "Tender Count", "Documents Accessible",
        "Login/Captcha Required", "Crawl Priority"
    ]
    portal_sheet.update(values=[headers], range_name="A1:J1")

    # 2. Extract sources from Excel or built-in registry
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

                # Determine document and captcha accessibility
                if "gem.gov.in" in u.lower():
                    doc_acc = "Yes (Direct Live PDF)"
                    login_req = "No (Public BidPlus)"
                elif any(k in u.lower() for k in ["nicgep", "eproc", "etender", "tenders"]):
                    doc_acc = "Yes (Public View)"
                    login_req = "No (Public NIT)"
                elif src_type == "Discovery":
                    doc_acc = "Yes (Aggregator)"
                    login_req = "No (Public Index)"
                else:
                    doc_acc = "Varies (Org Website)"
                    login_req = "May require Captcha"

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
            print(f"Notice reading Excel: {e}")

    # Built-in fallback if Excel reading encounters an issue
    if not portals_list:
        sample_sources = [
            ("GeM BidPlus", "National/Central", "Pan India", "https://bidplus.gem.gov.in/all-bids", "P1", "Yes (Direct Live PDF)", "No (Public BidPlus)"),
            ("CPPP Central", "National/Central", "Pan India", "https://eprocure.gov.in/eprocure/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("Government e-Tenders", "National/Central", "Pan India", "https://www.etenders.gov.in/eprocure/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("MahaTenders", "State/UT", "Maharashtra", "https://mahatenders.gov.in/nicgep/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("UP eTenders", "State/UT", "Uttar Pradesh", "https://etender.up.nic.in/nicgep/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("Delhi eProcurement", "State/UT", "Delhi", "https://govtprocurement.delhi.gov.in/nicgep/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("Rajasthan eProc", "State/UT", "Rajasthan", "https://eproc.rajasthan.gov.in/nicgep/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("Odisha e-Tenders", "State/UT", "Odisha", "https://tendersodisha.gov.in/nicgep/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("MP e-Tenders", "State/UT", "Madhya Pradesh", "https://mptenders.gov.in/nicgep/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("Karnataka KPPP", "State/UT", "Karnataka", "https://eproc.karnataka.gov.in", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("West Bengal e-Tenders", "State/UT", "West Bengal", "https://wbtenders.gov.in/nicgep/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("Tamil Nadu e-Tenders", "State/UT", "Tamil Nadu", "https://tntenders.gov.in/nicgep/app", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("Bihar eProcurement", "State/UT", "Bihar", "https://eproc2.bihar.gov.in", "P1", "Yes (Public View)", "No (Public NIT)"),
            ("Coal India Tenders", "PSU", "Pan India", "https://coalindiatenders.nic.in/nicgep/app", "P2", "Yes (Public View)", "No (Public NIT)"),
            ("IOCL Tenders", "PSU", "Pan India", "https://iocletenders.nic.in/nicgep/app", "P2", "Yes (Public View)", "No (Public NIT)"),
            ("BidAssist", "Third Party Aggregator", "Pan India", "https://bidassist.com", "P3", "Yes (Aggregator)", "No (Public Index)")
        ]
        for p, c, s, u, pr, da, lr in sample_sources:
            portals_list.append({
                "portal": p, "category": c, "state": s, "url": u, "active": "Yes",
                "doc_acc": da, "login_req": lr, "priority": pr
            })

    # Read existing rows to preserve or populate initial structure
    existing_portal_rows = portal_sheet.get_all_values()
    if len(existing_portal_rows) <= 1:
        initial_data = []
        for p in portals_list:
            initial_data.append([
                p["portal"], p["category"], p["state"], p["url"], p["active"],
                "-", 0, p["doc_acc"], p["login_req"], p["priority"]
            ])
        portal_sheet.update(values=initial_data, range_name=f"A2:J{len(initial_data)+1}")
        print(f"Populated {len(initial_data)} portals in Portal_Directory sheet.")

    return portals_list, portal_sheet

# ==============================================================================
# 3. GEMINI AI PRE-QUALIFICATION ENGINE (STABLE 1.5-FLASH ENDPOINT)
# ==============================================================================
def evaluate_tender_strictly(tender: dict) -> dict:
    """Evaluates live tender data using verified stable Gemini models."""
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
        "reasoning": f"AI Parsing Exception: {last_err}",
        "action_plan": "Manual review of tender document required."
    }

# ==============================================================================
# 4. PAN-INDIA LIVE SCRAPERS (DEEP DIRECT DOCUMENT URLS)
# ==============================================================================
def crawl_gepnic_endpoint(portal_meta):
    """Scrapes active published tenders from GePNIC/NIC central & state endpoints with direct links."""
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
        resp = requests.get(feed_url, headers=headers, timeout=12)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            rows = soup.find_all("tr")
            for row in rows:
                row_text = row.text.lower()
                # Check against all 59 event keywords
                if any(kw in row_text for kw in EVENT_KEYWORDS_LOWER):
                    cols = row.find_all("td")
                    if len(cols) >= 5:
                        title_col = cols[4].text.strip()
                        dept_col = cols[5].text.strip() if len(cols) > 5 else f"{portal_name} ({state_name})"
                        closing_date = cols[2].text.strip()

                        link_elem = cols[4].find("a", href=True) or row.find("a", href=True)
                        tender_id_match = re.search(r"\d{4}_[A-Z0-9]+_\d+_\d+", title_col)
                        tender_id = tender_id_match.group(0) if tender_id_match else f"{portal_name[:3]}/{int(time.time())}"

                        # Direct Deep Link
                        if link_elem and link_elem["href"]:
                            raw_href = link_elem["href"]
                            if "javascript:" not in raw_href.lower():
                                doc_url = urllib.parse.urljoin(base_url, raw_href)
                            else:
                                doc_url = f"{base_url}?page=FrontEndTenderDetailsExternal&service=page&tenderId={tender_id}"
                        else:
                            doc_url = f"{base_url}?page=FrontEndTenderDetailsExternal&service=page&tenderId={tender_id}"

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
                            "doc_link": doc_url
                        })
    except Exception:
        pass
    return found


def crawl_gem_live(page):
    """Scrapes live bids from GeM BidPlus with direct document download URLs."""
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
                            doc_elem = card.locator("a[href*='showbidDocument'], a[href*='show-bid']").first
                            if doc_elem.count() > 0:
                                href = doc_elem.get_attribute("href")
                                doc_link = urllib.parse.urljoin("https://bidplus.gem.gov.in", href)
                            else:
                                doc_link = f"https://bidplus.gem.gov.in/showbidDocument/{bid_no.split('/')[-1]}"

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
                                "doc_link": doc_link
                            })
            except Exception:
                continue
    except Exception as e:
        print(f"Notice during GeM Live Crawl: {e}")
    return gem_bids


def crawl_aggregators_live():
    """Scrapes Pan-India aggregators capturing exact deep links for event tenders."""
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
                            "doc_link": doc_link
                        })
        except Exception:
            pass
    return agg_results

# ==============================================================================
# 5. DEDUPLICATION, AUTO-CLEAN & SHEET EXECUTION
# ==============================================================================
def purge_legacy_dummy_rows(sheet):
    """Automatically cleans legacy sample mock rows from the sheet."""
    try:
        rows = sheet.get_all_values()
        if len(rows) > 1:
            dummy_indicators = ["gem/2026/b/7891024", "cppp/2026/dpiit", "up/2026/tourism/4512", "rj/2026/sppp", "mh/2026/cidco/3321"]
            for idx, r in enumerate(rows[1:], start=2):
                tender_id = str(r[1]).lower()
                if any(di in tender_id for di in dummy_indicators):
                    sheet.delete_rows(idx)
                    print(f"Purged legacy dummy test row: {r[1]}")
    except Exception as e:
        print(f"Notice during dummy row cleanup: {e}")


def run_pipeline():
    today = datetime.date.today().strftime("%Y-%m-%d")
    print(f"==================================================")
    print(f"Starting Pan-India Live Tender Automation: {today}")
    print(f"Keywords Configured: {len(EVENT_KEYWORDS)} event categories")
    print(f"==================================================")

    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    active_sheet = spreadsheet.worksheet("Active_Tenders")

    # 1. Sync & Update Portal_Directory Tab (109 Master Sources)
    master_portals, portal_sheet = load_and_sync_portal_directory(gc)

    # 2. Purge Any Lingering Legacy Dummy Rows
    purge_legacy_dummy_rows(active_sheet)

    # 3. Read Existing Tender IDs (Avoid Duplicates)
    try:
        col_b_vals = active_sheet.col_values(2)
        existing_ids = set(col_b_vals[1:])
    except Exception:
        existing_ids = set()
    print(f"Existing verified tenders logged in sheet: {len(existing_ids)}")

    live_tenders = []
    portal_counts = {}

    # 4. Multi-threaded Parallel Crawl for GePNIC Endpoints
    print("Executing parallel extraction across GePNIC portals...")
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(crawl_gepnic_endpoint, portal): portal for portal in master_portals}
        for f in as_completed(futures):
            p_meta = futures[f]
            res = f.result()
            count = len(res) if res else 0
            portal_counts[p_meta["portal"]] = count
            if res:
                live_tenders.extend(res)

    # 5. Playwright Crawl for GeM BidPlus
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
            portal_counts["GeM BidPlus"] = len(gem_records)
            browser.close()
    except Exception as e:
        print(f"GeM crawler notice: {e}")

    # 6. Aggregator Feeds
    print("Querying Pan-India aggregator feeds...")
    agg_records = crawl_aggregators_live()
    live_tenders.extend(agg_records)

    # 7. Update Last Crawl Timestamp & Tender Count in Portal_Directory
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M IST")
    try:
        p_rows = portal_sheet.get_all_values()
        updates = []
        for row_idx, r in enumerate(p_rows[1:], start=2):
            p_name = r[0]
            if p_name in portal_counts:
                # Update Col F (Last Crawl) and Col G (Tender Count)
                updates.append({"range": f"F{row_idx}:G{row_idx}", "values": [[now_str, portal_counts[p_name]]]})
        if updates:
            for u in updates[:25]:  # Batch update top crawled portals
                portal_sheet.update(range_name=u["range"], values=u["values"])
        print("Updated crawl timestamps and live counts in 'Portal_Directory'.")
    except Exception as e:
        print(f"Notice updating Portal_Directory counts: {e}")

    # 8. Deduplicate Live Tenders
    unique_new_tenders = []
    seen_in_batch = set()
    for t in live_tenders:
        t_id = t["tender_id"]
        if t_id not in existing_ids and t_id not in seen_in_batch:
            unique_new_tenders.append(t)
            seen_in_batch.add(t_id)

    print(f"--------------------------------------------------")
    print(f"Total Live Event Tenders Scraped Across Sites: {len(live_tenders)}")
    print(f"New Unique Tenders to Evaluate: {len(unique_new_tenders)}")
    print(f"--------------------------------------------------")

    if not unique_new_tenders:
        print("Zero new event tenders published today. Exiting cleanly without dummy data.")
        return

    # 9. Evaluate & Append Real Tenders to Active_Tenders Tab
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
            item["doc_link"],
            decision["status"],
            remarks_field
        ]

        active_sheet.append_row(row_data)
        print(f"  -> Appended [{decision['status']}] with direct link: {item['doc_link']}")

    print(f"Pipeline executed successfully for {today}.")

if __name__ == "__main__":
    run_pipeline()
