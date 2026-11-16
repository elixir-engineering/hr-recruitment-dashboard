/* ---------------------------------------------------------------------------
   Cloudflare Pages Function — POST /api/refresh

   Triggers the "Refresh HR data" GitHub Action (workflow_dispatch) so the
   dashboard's Rebuild button can pull a fresh copy of the workbook without
   anyone opening the Actions tab.

   It only STARTS the job. The Action does the Graph download and the SheetJS
   parse (~13 s) and writes to KV; the page then polls /api/data until the
   content hash changes. No parsing happens here, so this stays inside the
   10 ms CPU budget — it is one outbound fetch and nothing else.

   Environment variables (Cloudflare Pages -> Settings -> Environment variables):
     GH_TOKEN   GitHub fine-grained PAT, Actions: Read and write, this repo only
                MUST be marked as a Secret.
     GH_REPO    owner/repo, e.g. elixir-engineering/hr-recruitment-dashboard
     GH_WORKFLOW  optional, defaults to refresh-hr-data.yml
     GH_REF       optional, defaults to main

   NOTE ON ABUSE: this endpoint is public, like the rest of the dashboard.
   Anyone who finds the URL can start a build. The cooldown below caps that at
   one dispatch per COOLDOWN_MS per edge location, so the worst case is wasted
   GitHub Actions minutes rather than anything destructive — the token is
   scoped to triggering this one workflow and cannot read or write code.
--------------------------------------------------------------------------- */

const COOLDOWN_MS = 60000;      // one manual rebuild a minute is plenty
let LAST_DISPATCH = 0;          // module scope: per warm isolate, best-effort

const json = (obj, status) => new Response(JSON.stringify(obj), {
  status,
  headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' }
});

export async function onRequest(context) {
  const { request, env } = context;

  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 204, headers: { 'allow': 'POST, OPTIONS' } });
  }
  if (request.method !== 'POST') {
    return json({ error: 'method', message: 'Use POST.' }, 405);
  }

  const token = env.GH_TOKEN;
  const repo = env.GH_REPO;
  const workflow = env.GH_WORKFLOW || 'refresh-hr-data.yml';
  const ref = env.GH_REF || 'main';

  if (!token || !repo) {
    return json({
      error: 'not_configured',
      missing: [!token && 'GH_TOKEN', !repo && 'GH_REPO'].filter(Boolean),
      message: 'Set these in Cloudflare Pages -> Settings -> Environment variables, then redeploy. '
        + 'GH_TOKEN must be marked Secret.'
    }, 503);
  }

  const since = Date.now() - LAST_DISPATCH;
  if (since < COOLDOWN_MS) {
    return json({
      error: 'cooldown',
      retryInSeconds: Math.ceil((COOLDOWN_MS - since) / 1000),
      message: 'A rebuild was just started. Give it a moment.'
    }, 429);
  }

  try {
    const r = await fetch(
      'https://api.github.com/repos/' + repo + '/actions/workflows/' + workflow + '/dispatches',
      {
        method: 'POST',
        headers: {
          Authorization: 'Bearer ' + token,
          Accept: 'application/vnd.github+json',
          'X-GitHub-Api-Version': '2022-11-28',
          'User-Agent': 'hr-dashboard-refresh',
          'content-type': 'application/json'
        },
        body: JSON.stringify({ ref })
      }
    );

    /* GitHub answers 204 No Content on success and says nothing useful. */
    if (r.status === 204) {
      LAST_DISPATCH = Date.now();
      return json({ ok: true, startedAt: new Date().toISOString(),
        message: 'Rebuild started. It takes about 20 seconds.' }, 200);
    }

    let detail = 'HTTP ' + r.status;
    try { detail = (JSON.parse(await r.text()).message) || detail; } catch { /* non-JSON */ }

    if (r.status === 401 || r.status === 403) {
      return json({ error: 'auth', message: 'GitHub rejected the token (' + detail + '). '
        + 'Check GH_TOKEN has Actions: Read and write on this repository and has not expired.' }, 502);
    }
    if (r.status === 404) {
      return json({ error: 'not_found', message: 'GitHub returned 404 (' + detail + '). '
        + 'Check GH_REPO is owner/repo, the workflow file name is right, and the token can see this repo.' }, 502);
    }
    return json({ error: 'github', message: detail }, 502);

  } catch (e) {
    return json({ error: 'failed', message: String((e && e.message) || e) }, 502);
  }
}
