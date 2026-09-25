const GEM_BASE = "https://bidplus.gem.gov.in";
const ALL_BIDS = GEM_BASE + "/all-bids";
const ALL_BIDS_DATA = GEM_BASE + "/all-bids-data";

const USER_AGENT =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36";

function arrayFirst(value, fallback = "") {
  if (Array.isArray(value)) return value.length ? value[0] : fallback;
  return value ?? fallback;
}

function extractCookies(response) {
  const raw = response.headers.get("set-cookie") || "";
  if (!raw) return "";
  return raw
    .split(/,(?=[^;,]+=[^;,]+)/)
    .map((part) => part.split(";")[0].trim())
    .filter(Boolean)
    .join("; ");
}

function extractCsrf(html) {
  const patterns = [
    /csrf_bd_gem_nk[^'"]*['"]([a-f0-9]{32})['"]/i,
    /name=["']csrf_bd_gem_nk["'][^>]*value=["']([^"']+)["']/i,
    /csrf_bd_gem_nk\s*[:=]\s*["']([^"']+)["']/i,
  ];

  for (const pattern of patterns) {
    const match = html.match(pattern);
    if (match) return match[1];
  }
  return "";
}

async function initGemSession() {
  const response = await fetch(ALL_BIDS, {
    method: "GET",
    redirect: "follow",
    headers: {
      "user-agent": USER_AGENT,
      accept: "text/html,application/xhtml+xml",
      "accept-language": "en-US,en;q=0.9",
      "cache-control": "no-cache",
    },
  });

  const html = await response.text();
  const csrf = extractCsrf(html);
  const cookie = extractCookies(response);

  return {
    ok: response.ok,
    status: response.status,
    finalUrl: response.url,
    csrf,
    cookie,
    htmlLength: html.length,
    html,
  };
}

async function searchGem({ query, page = 1, bidType = "service" }) {
  const session = await initGemSession();

  if (!session.ok) {
    const error = new Error("GeM all-bids returned HTTP " + session.status);
    error.code = "GEM_HTTP_" + session.status;
    throw error;
  }

  if (!session.csrf) {
    const error = new Error("GeM page loaded but CSRF token was not found.");
    error.code = "GEM_CSRF_NOT_FOUND";
    throw error;
  }

  const payload = {
    page,
    param: {
      searchBid: query,
      searchType: "fullText",
    },
    filter: {
      bidStatusType: "ongoing_bids",
      byType: bidType,
      highBidValue: "",
      byEndDate: {
        from: "",
        to: "",
      },
      sort: "Bid-End-Date-Latest",
    },
  };

  const form = new URLSearchParams();
  form.set("payload", JSON.stringify(payload));
  form.set("csrf_bd_gem_nk", session.csrf);

  const response = await fetch(ALL_BIDS_DATA, {
    method: "POST",
    redirect: "follow",
    headers: {
      "user-agent": USER_AGENT,
      accept: "application/json,text/plain,*/*",
      "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
      "x-requested-with": "XMLHttpRequest",
      referer: ALL_BIDS,
      origin: GEM_BASE,
      cookie: session.cookie,
    },
    body: form.toString(),
  });

  const raw = await response.text();

  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    const error = new Error(
      "GeM all-bids-data did not return JSON. HTTP " + response.status
    );
    error.code = "GEM_NON_JSON_RESPONSE";
    error.preview = raw.slice(0, 300);
    throw error;
  }

  if (String(data.code) === "404") {
    return { total: 0, docs: [], session };
  }

  if (String(data.code) !== "200") {
    const error = new Error(
      "GeM all-bids-data returned code " + String(data.code)
    );
    error.code = "GEM_API_CODE_" + String(data.code);
    throw error;
  }

  const responseBody = data?.response?.response || {};
  return {
    total: Number(responseBody.numFound || 0),
    docs: Array.isArray(responseBody.docs) ? responseBody.docs : [],
    session,
  };
}

function parseGemDate(value) {
  const raw = arrayFirst(value, "");
  if (!raw) return null;

  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function isFutureDeadline(value) {
  const date = parseGemDate(value);
  if (!date) return false;
  return date.getTime() >= Date.now();
}

const SERVICE_PHRASES = [
  "event management",
  "event agency",
  "event or seminar or workshop or exhibition or expo management service",
  "exhibition management",
  "exhibition stall",
  "stall fabrication",
  "stall design",
  "pavilion fabrication",
  "pavilion design",
  "conference management",
  "convention management",
  "summit management",
  "conclave management",
  "creative agency",
  "advertising agency",
  "communication agency",
  "social media agency",
  "outreach campaign",
  "publicity campaign",
  "brand activation",
  "experiential marketing",
  "audio visual coverage",
  "photography videography",
  "photography and videography",
  "stage production",
  "mela management",
  "festival management",
];

function normalize(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function isRelevantService(doc) {
  const type = Number(arrayFirst(doc.b_type, 0));
  const title = normalize(arrayFirst(doc.bbt_title, ""));
  const category = normalize(arrayFirst(doc.b_category_name, ""));
  const text = (title + " " + category).toLowerCase();

  if (type !== 1) return false;
  return SERVICE_PHRASES.some((phrase) => text.includes(phrase));
}

function formatDoc(doc) {
  const bidNumber = normalize(arrayFirst(doc.b_bid_number, ""));
  const bidId = normalize(arrayFirst(doc.b_id, ""));
  const title = normalize(arrayFirst(doc.bbt_title, ""));
  const category = normalize(arrayFirst(doc.b_category_name, ""));
  const department = normalize(arrayFirst(doc.ba_official_details_deptName, ""));
  const ministry = normalize(arrayFirst(doc.ba_official_details_minName, ""));
  const start = parseGemDate(doc.final_start_date_sort);
  const end = parseGemDate(doc.final_end_date_sort);

  return {
    bid_number: bidNumber,
    bid_id: bidId,
    title,
    category,
    department,
    ministry,
    is_service: Number(arrayFirst(doc.b_type, 0)) === 1,
    opening_date: start ? start.toISOString() : null,
    submission_deadline: end ? end.toISOString() : null,
    url: bidId ? GEM_BASE + "/showbidDocument/" + bidId : ALL_BIDS,
  };
}

function authorize(req) {
  const token = process.env.GEM_PROXY_TOKEN || "";
  if (!token) return true;

  const auth = req.headers.authorization || "";
  return auth === "Bearer " + token;
}

function sendJson(res, status, body) {
  res.status(status);
  res.setHeader("content-type", "application/json; charset=utf-8");
  res.setHeader("cache-control", "no-store");
  res.json(body);
}

module.exports = {
  ALL_BIDS,
  ALL_BIDS_DATA,
  authorize,
  formatDoc,
  initGemSession,
  isFutureDeadline,
  isRelevantService,
  searchGem,
  sendJson,
};
