# Event Tender Crawler — Pan-India Tender Intelligence Agent

Daily tender discovery and pre-qualification workflow for **Soul Events and Consultancy**.

## What it does

1. Loads active sources from `Pan_India_Tender_URL_Master.xlsx`.
2. Crawls GePNIC/state tender portals, GeM, and generic PSU/organisation tender pages.
3. Searches for event, exhibition, experiential, conference, activation, AV, branding and related tenders.
4. Downloads publicly accessible tender documents.
5. Extracts text from PDF, DOCX, XLSX and ZIP files.
6. Uses Gemini only to extract documented eligibility facts.
7. Applies deterministic Python qualification rules against Soul Events' configured benchmarks.
8. Writes results to Google Sheets and records crawler health for every portal.

## Important safety rule

The agent **does not invent** EMD exemption, turnover, past-experience requirements, MSME relaxation, deadlines or entity restrictions. If the official documents cannot be read or a criterion is unclear, the tender is marked:

`DOCUMENT REVIEW REQUIRED`

## Qualification statuses

- `VERIFIED QUALIFIED`
- `VERIFIED DISQUALIFIED`
- `NEEDS MANUAL INTERVENTION`
- `DOCUMENT REVIEW REQUIRED`

## Google Sheet tabs

### `Active_Tenders`
Stores tender details, document URLs, extracted eligibility evidence, qualification result, evidence quotes, and action plan.

### `Portal_Directory`
Stores each source URL plus last crawl time, tender count, crawler status, and crawler message.

Typical crawler statuses include:

- `SUCCESS`
- `TIMEOUT`
- `HTTP_ERROR`
- `LOGIN_OR_BLOCKED`
- `CAPTCHA_OR_LOGIN`
- `PARSER_FAILED`
- `SKIPPED_DEDICATED_ADAPTER`

## GitHub configuration

### Required repository secrets

Create these under **Settings → Secrets and variables → Actions → Secrets**:

- `GEMINI_API_KEY`
- `GCP_SA_KEY` — full Google service-account JSON

### Recommended repository variable

Under **Settings → Secrets and variables → Actions → Variables**:

- `SPREADSHEET_ID`

If `SPREADSHEET_ID` is not configured, the script currently keeps the existing project sheet ID as its fallback.

## Google service account

Share the target Google Sheet with the `client_email` contained in the service-account JSON and give it Editor access.

## Local installation

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
```

Place these files in the repository root:

- `Pan_India_Tender_URL_Master.xlsx`
- `service_account.json` (local use only; never commit it)

Then configure:

```bash
export GEMINI_API_KEY="YOUR_KEY"
export SPREADSHEET_ID="YOUR_SHEET_ID"
python tender_agent.py
```

## Daily schedule

GitHub Actions runs at:

- `03:00 UTC`
- `08:30 AM IST`

The workflow can also be started manually using **Run workflow**.

## Coverage limitation

The source Excel contains 109 active URLs, but government procurement platforms do not all use the same technical interface. The agent uses:

- a GePNIC-style crawler for compatible e-procurement portals,
- a dedicated GeM browser adapter,
- a generic tender-link crawler for organisation/PSU websites.

Special systems such as IREPS, MSTC, login-only portals, CAPTCHA-protected pages, and portals with unusual JavaScript interfaces may require dedicated adapters. The agent intentionally does not bypass authentication or CAPTCHA controls.

## Recommended next adapters

For higher Pan-India coverage, add dedicated adapters for:

1. IREPS
2. MSTC
3. high-priority PSU tender systems
4. third-party discovery aggregators (discovery only; always verify from the issuing authority)

## Files

- `tender_agent.py` — main crawler and qualification pipeline
- `Pan_India_Tender_URL_Master.xlsx` — source directory
- `requirements.txt` — Python dependencies
- `.github/workflows/daily_crawler.yml` — daily automation
- `.gitignore` — prevents credentials/downloads from being committed
