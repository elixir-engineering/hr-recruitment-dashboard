/* ---------------------------------------------------------------------------
   Cloudflare Pages Function — GET /api/data

   Same public URL and same JSON shape as the Vercel route in api/, so the
   front end is identical on both hosts. Only the header x-backend differs,
   which is how the page reports which host answered.

   THE CPU RULE
   Cloudflare's free tier allows 10 ms of CPU per invocation. Measured on this
   workbook, parsing it costs 103 ms — 10x over — so the parse CANNOT happen
   here. A scheduled GitHub Action (.github/workflows/refresh-hr-data.yml)
   does the Graph download and the SheetJS parse, then writes finished JSON to
   KV. This function only streams those bytes.

   Do not "simplify" this by calling .get(key, 'json') or JSON.parse() on the
   payload — that puts a 476 KB parse back in the request path. The payload is
   read with { type: 'stream' } and handed straight to the Response, so its
   bytes are never touched by JS. Measured cost: ~0.2 ms.

   KV keys (binding name HR_DATA):
     payload  big JSON { tracker, reqs, reqsError, hash, fetchedAt }
              with KV metadata { hash, fetchedAt, trackerRows, reqsRows }
     meta     small JSON used only by ?debug=1
--------------------------------------------------------------------------- */

const JSON_HEADERS = { 'content-type': 'application/json; charset=utf-8', 'x-backend': 'cloudflare' };

const cacheHeader = (fresh, debug) => (fresh || debug)
  ? 'no-store'
  : 'public, s-maxage=300, stale-while-revalidate=600';

function reply(obj, status, extra) {
  return new Response(JSON.stringify(obj), { status, headers: { ...JSON_HEADERS, ...extra } });
}

export async function onRequest(context) {
  const { request, env } = context;

  if (request.method !== 'GET' && request.method !== 'HEAD') {
    return reply({ error: 'method', message: 'Use GET.' }, 405, { 'cache-control': 'no-store' });
  }

  const url = new URL(request.url);
  const debug = url.searchParams.has('debug');
  const fresh = url.searchParams.has('fresh');
  const have = url.searchParams.get('v');
  const cc = { 'cache-control': cacheHeader(fresh, debug) };

  if (!env.HR_DATA) {
    return reply({
      error: 'not_configured',
      message: 'KV binding HR_DATA is missing. Cloudflare Pages -> Settings -> Functions '
        + '-> KV namespace bindings -> add variable HR_DATA pointing at the namespace.'
    }, 503, { 'cache-control': 'no-store' });
  }

  /* ?debug=1 reads the SMALL meta key, never the payload */
  if (debug) {
    const meta = await env.HR_DATA.get('meta', { type: 'json' });
    if (!meta) {
      return reply({
        ok: false, backend: 'cloudflare',
        error: 'empty',
        message: 'KV holds no data yet. Run the "Refresh HR data" GitHub Action once '
          + '(Actions tab -> Refresh HR data -> Run workflow).'
      }, 503, { 'cache-control': 'no-store' });
    }
    return reply({ ok: true, backend: 'cloudflare', ...meta }, 200, { 'cache-control': 'no-store' });
  }

  let got;
  try {
    got = await env.HR_DATA.getWithMetadata('payload', { type: 'stream' });
  } catch (e) {
    return reply({ error: 'kv', message: 'KV read failed: ' + (e && e.message || e) }, 502,
      { 'cache-control': 'no-store' });
  }

  if (!got || !got.value) {
    return reply({
      error: 'empty',
      message: 'No data in KV yet. Run the "Refresh HR data" GitHub Action once to populate it.'
    }, 503, { 'cache-control': 'no-store' });
  }

  const meta = got.metadata || {};

  /* The page sends back the hash it already holds. If nothing changed we answer
     in ~80 bytes and discard the stream unread — cancelling it matters, an
     abandoned stream would otherwise stay open. */
  if (have && meta.hash && have === meta.hash && !fresh) {
    try { await got.value.cancel(); } catch (e) { /* already closed */ }
    return reply({
      unchanged: true, hash: meta.hash, fetchedAt: meta.fetchedAt || null,
      ageSeconds: meta.fetchedAt ? Math.round((Date.now() - Date.parse(meta.fetchedAt)) / 1000) : null
    }, 200, cc);
  }

  /* Stream the stored bytes straight through. No parse, no copy. */
  return new Response(got.value, {
    status: 200,
    headers: {
      ...JSON_HEADERS,
      ...cc,
      'x-data-hash': meta.hash || '',
      'x-data-fetched-at': meta.fetchedAt || ''
    }
  });
}
