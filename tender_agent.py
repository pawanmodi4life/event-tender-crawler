import os
import json
import datetime
from google import genai
from google.oauth2 import service_account
import gspread
from playwright.sync_api import sync_playwright

# ==============================================================================
# 1. VERIFIED SOUL EVENTS & CONSULTANCY PROFILE (AUDITED DOSSIER BASELINE)
# ==============================================================================
COMPANY_PROFILE = {
    "agency_name": "Soul Events and Consultancy",
    "constitution": "Proprietorship",
    "proprietor": "Ashish Garg",
    "vintage_years": 13,
    "msme_status": "Micro Enterprise (Services) - Valid Udyam MH-19-0113805",
    "gstin": "27AOLPG1479A2Z3",
    "financials": {
        "avg_3yr_turnover_inr": 38046000,      # Rs. 3.80 Crores (CA Certified UDIN: 26145975UYPNOU5226)
        "fy_2025_26_turnover_inr": 50411575,  # Rs. 5.04 Crores
        "audited_net_worth_inr": 4680004       # Rs. 46.80 Lakhs (Positive across all years)
    },
    "verified_work_orders": [
        {"client": "Mahanadi Coalfields Ltd (CIL)", "title": "Celebrity Star Night (Shri Sukhwinder Singh)", "val": 7898550, "type": "PSU"},
        {"client": "Hindustan Petroleum Corp Ltd (HPCL)", "title": "Stage & Venue Decor Deepotsav", "val": 7223000, "type": "PSU"},
        {"client": "Epiroc Mining India Pvt Ltd", "title": "Groundbreaking Ceremony & German Hangar", "val": 6857653, "type": "Private/MNC"},
        {"client": "New India Assurance Co Ltd (NIA)", "title": "108th Foundation Day at NCPA Mumbai", "val": 4628443, "type": "PSU"},
        {"client": "General Insurance Corp (GIC Re)", "title": "Underwriters Meet Indore (Hospitality & Event)", "val": 3338320, "type": "PSU"},
        {"client": "Central Inst of Fisheries Education (ICAR)", "title": "Fish Fair Exhibition Stalls & Pandal", "val": 1540950, "type": "Govt Autonomous"}
    ],
    "max_single_past_work_order": 7898550,    # Rs. 78.98 Lakhs
    "max_two_works_threshold": 7223000,        # Two works >= Rs. 72.23 Lakhs
    "max_three_works_threshold": 4628443       # Three works >= Rs. 46.28 Lakhs
}

SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1YgsKZNiamaZ-E2ZZhIY2MLm8diC5okmeUMGFQdgQzHQ")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

# ==============================================================================
# 2. STRICT 3-TIER EVALUATION ENGINE
# ==============================================================================
def evaluate_tender_strictly(tender: dict) -> dict:
    """
    Evaluates tender eligibility strictly enforcing:
    1. QUALIFIED
    2. DISQUALIFIED
    3. NEEDS MANUAL INTERVENTION
    """
    client = genai.Client(api_key=GEMINI_API_KEY)
    
    prompt = f"""
    You are the Chief Procurement Legal Officer for Soul Events and Consultancy.
    Evaluate the following tender against our audited capability dossier with absolute strictness.
    
    FIRM PROFILE & AUDITED CAPABILITY:
    {json.dumps(COMPANY_PROFILE, indent=2)}
    
    TENDER DATA TO EVALUATE:
    {json.dumps(tender, indent=2)}
    
    STRICT DECISION RULES (DO NOT DEVIATE):
    
    A. MARK AS 'DISQUALIFIED' IF ANY OF THESE OCCUR:
       1. Entity Type Restriction: The tender explicitly restricts bidding to "Public Limited / Private Limited company only" (We are a Proprietorship).
       2. Turnover Gate: Tender requires average turnover > Rs. 3.80 Crores AND explicitly states "No MSME relaxation / No Consortium permitted".
       3. Work Order Gate: Requires a single work order > Rs. 78.98 Lakhs AND forbids Joint Ventures / Consortiums.
       4. Net Worth Gate: Requires Net Worth > Rs. 46.80 Lakhs.
       5. Non-Core Scope: Pure civil engineering construction (CPWD enlistment), security guard staffing (PSARA license), or centralized food manufacturing.
    
    B. MARK AS 'NEEDS MANUAL INTERVENTION' IF:
       1. Celebrity Artist / Star Night: Requires named celebrity artist booking (Requires securing exclusive Artist Mandate Letter + Video Byte within 48h).
       2. Mega Value with JV (Rs. 1.5 Cr to Rs. 10 Cr): Tender exceeds our direct single work order (Rs. 78.98L) but tender allows Consortium / JV bidding.
       3. QCBS Technical Pitch: High-value conclave/branding requiring comprehensive 60-slide creative presentation & 3D renders.
       4. Physical Submission: Requires physical Demand Draft / Bank Guarantee courier before bid closing.
    
    C. MARK AS 'QUALIFIED' IF:
       1. Scope matches Event Management, Corporate Conclave, Stage/Sound/LED Setup, Exhibition Stalls, or Groundbreaking.
       2. Estimated Value <= Rs. 1.00 Crore (or <= Rs. 1.45 Cr under the 2-work 50% rule).
       3. Turnover required <= Rs. 3.80 Crores.
       4. Single past work required <= Rs. 78.98 Lakhs.
       5. EMD is 100% exempt for MSME or standard portal terms apply.
    
    Output strictly JSON format with keys:
    {{"status": "QUALIFIED" | "DISQUALIFIED" | "NEEDS MANUAL INTERVENTION", "reasoning": "...", "action_plan": "..."}}
    """
    
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config={"response_mime_type": "application/json"}
        )
        return json.loads(response.text)
    except Exception as e:
        return {
            "status": "NEEDS MANUAL INTERVENTION",
            "reasoning": f"AI Parsing Exception: {str(e)}",
            "action_plan": "Manual review of tender document required."
        }

# ==============================================================================
# 3. SHEET LOGGER & PIPELINE EXECUTION
# ==============================================================================
def get_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID).worksheet("Active_Tenders")

def run_agent():
    today = datetime.date.today().strftime("%Y-%m-%d")
    print(f"Executing Strict Tender Evaluation Agent: {today}")
    sheet = get_google_sheet()
    
    # Active incoming tenders feed for evaluation
    tenders_pool = [
        {
            "portal": "GeM Portal",
            "tender_id": f"GEM/2026/B/STAR_{datetime.datetime.now().strftime('%M%S')}",
            "organization": "NTPC Limited (Central PSU)",
            "title": "Celebrity Star Night & Cultural Evening for Raising Day Celebrations",
            "category": "Artist / Star Night",
            "estimated_value": "85,00,000",
            "emd": "Exempted for MSME",
            "deadline": (datetime.datetime.now() + datetime.timedelta(days=12)).strftime("%Y-%m-%d 15:00"),
            "turnover_req": "Rs. 25 Lakhs in last 3 FYs",
            "experience_req": "1 celebrity concert work of Rs. 68 Lakhs (80%)",
            "tech_req": "Exclusive Artist Authorization Letter, Video Byte within 48h, Official Sound & Light Rider",
            "doc_link": "https://gem.gov.in"
        },
        {
            "portal": "UP eTender",
            "tender_id": f"UP/TOURISM/2026/{datetime.datetime.now().strftime('%M%S')}",
            "organization": "UP State Tourism Board",
            "title": "Turnkey Event Management, P3 LED Wall, Trussing & German Hangar for Handicraft Expo",
            "category": "Turnkey Production",
            "estimated_value": "65,00,000",
            "emd": "Exempted for MSME",
            "deadline": (datetime.datetime.now() + datetime.timedelta(days=14)).strftime("%Y-%m-%d 17:00"),
            "turnover_req": "Rs. 20 Lakhs",
            "experience_req": "1 similar event of Rs. 52 Lakhs or 2 of Rs. 32 Lakhs",
            "tech_req": "German Hangar setup, Sound line array, 3D branding, barricading",
            "doc_link": "https://etender.up.nic.in"
        },
        {
            "portal": "CPPP Portal",
            "tender_id": f"CPPP/2026/CIVIL/{datetime.datetime.now().strftime('%M%S')}",
            "organization": "Military Engineer Services (MES)",
            "title": "Construction of Permanent Officers Mess Building & Civil Structural Renovation",
            "category": "Civil Works",
            "estimated_value": "1,80,00,000",
            "emd": "Rs. 3,60,000",
            "deadline": (datetime.datetime.now() + datetime.timedelta(days=7)).strftime("%Y-%m-%d 14:00"),
            "turnover_req": "Rs. 90 Lakhs",
            "experience_req": "Class 'A' CPWD Enlistment & Civil RCC structural experience",
            "tech_req": "Concrete, brick masonry, structural engineering registration",
            "doc_link": "https://eprocure.gov.in"
        }
    ]

    for item in tenders_pool:
        decision = evaluate_tender_strictly(item)
        remarks_field = f"{decision['reasoning']} Action: {decision['action_plan']}"
        
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
        
        sheet.append_row(row_data)
        print(f"Processed: {item['tender_id']} --> Status: [{decision['status']}]")

if __name__ == "__main__":
    run_agent()
