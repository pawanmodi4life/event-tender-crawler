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
