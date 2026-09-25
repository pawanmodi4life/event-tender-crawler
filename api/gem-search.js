const {
  authorize,
  formatDoc,
  isFutureDeadline,
  isRelevantService,
  searchGem,
  sendJson,
} = require("./_gem-common");

module.exports = async function handler(req, res) {
  if (!authorize(req)) {
    return sendJson(res, 401, {
      ok: false,
      error: "Unauthorized",
    });
  }

  if (req.method !== "GET") {
    return sendJson(res, 405, {
      ok: false,
      error: "Method not allowed",
    });
  }

  const query = String(req.query.q || "Event Management").trim();
  const requestedPages = Math.max(
    1,
    Math.min(Number(req.query.pages || 3) || 3, 10)
  );

  const started = Date.now();
  const dedupe = new Map();
  let reportedTotal = 0;
  let pagesFetched = 0;

  try {
    for (let page = 1; page <= requestedPages; page += 1) {
      const result = await searchGem({
        query,
        page,
        bidType: "service",
      });

      reportedTotal = Math.max(reportedTotal, result.total);
      pagesFetched += 1;

      if (!result.docs.length) break;

      for (const doc of result.docs) {
        if (!isRelevantService(doc)) continue;
        if (!isFutureDeadline(doc.final_end_date_sort)) continue;

        const item = formatDoc(doc);
        const key = item.bid_number || item.bid_id || item.url;
        dedupe.set(key, item);
      }
    }

    const results = Array.from(dedupe.values()).sort((a, b) =>
      String(a.submission_deadline).localeCompare(
        String(b.submission_deadline)
      )
    );

    return sendJson(res, 200, {
      ok: true,
      query,
      source: "GeM standard BidPlus",
      bid_type: "service",
      deadline_filter: "submission_deadline >= current time",
      pages_fetched: pagesFetched,
      gem_reported_total: reportedTotal,
      matched_active_event_service_bids: results.length,
      fetched_at: new Date().toISOString(),
      elapsed_ms: Date.now() - started,
      results,
    });
  } catch (error) {
    return sendJson(res, 502, {
      ok: false,
      query,
      source: "GeM standard BidPlus",
      error: error?.code || error?.name || "GEM_SEARCH_ERROR",
      message: String(error?.message || error),
      preview: error?.preview || undefined,
      elapsed_ms: Date.now() - started,
    });
  }
};
