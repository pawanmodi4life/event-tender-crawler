# Vercel GeM Proxy

This repository now includes a lightweight Vercel-only GeM connector.

## Endpoints

- `GET /api/gem-test`
  - Tests whether `https://bidplus.gem.gov.in/all-bids` is reachable from Vercel.
  - Confirms whether a GeM CSRF session can be created.

- `GET /api/gem-search?q=Event%20Management&pages=3`
  - Searches the standard GeM BidPlus catalogue.
  - Requests only GeM `service` bids.
  - Keeps only event/exhibition/creative service opportunities.
  - Keeps only bids whose submission deadline is still in the future.
  - Returns opening date and submission deadline separately.

## Optional protection

Set a Vercel environment variable:

`GEM_PROXY_TOKEN=<your-random-secret>`

When configured, requests must include:

`Authorization: Bearer <your-random-secret>`

If the variable is not configured, the endpoints are public.

## Deployment

1. Import this GitHub repository into Vercel.
2. Deploy with the repository root as the project root.
3. Do not add a framework preset; Vercel should detect the `api/*.js` functions.
4. Test `/api/gem-test`.
5. If `csrf_found` is true, test `/api/gem-search?q=Event%20Management&pages=3`.

This service is intentionally separate from the long-running 109-source GitHub crawler.
