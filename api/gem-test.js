const {
  ALL_BIDS,
  ALL_BIDS_DATA,
  authorize,
  initGemSession,
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

  const started = Date.now();

  try {
    const session = await initGemSession();

    return sendJson(res, 200, {
      ok: session.ok && Boolean(session.csrf),
      provider: "Vercel",
      gem_host: "bidplus.gem.gov.in",
      all_bids_url: ALL_BIDS,
      all_bids_data_url: ALL_BIDS_DATA,
      http_status: session.status,
      final_url: session.finalUrl,
      csrf_found: Boolean(session.csrf),
      cookie_found: Boolean(session.cookie),
      html_length: session.htmlLength,
      elapsed_ms: Date.now() - started,
      message:
        session.ok && session.csrf
          ? "GeM is reachable from this Vercel function and a session was created."
          : "GeM responded, but a usable CSRF session could not be established.",
    });
  } catch (error) {
    return sendJson(res, 502, {
      ok: false,
      provider: "Vercel",
      gem_host: "bidplus.gem.gov.in",
      all_bids_url: ALL_BIDS,
      error: error?.code || error?.name || "GEM_CONNECTION_ERROR",
      message: String(error?.message || error),
      elapsed_ms: Date.now() - started,
    });
  }
};
