/* ---------------------------------------------------------------------------
   HR Dashboard — SharePoint / Microsoft Graph data endpoint (Vercel, Node)

   Replaces the old Google Sheets gviz call. A static page cannot read
   SharePoint directly: the file needs a Microsoft sign-in, and SharePoint
   sends no CORS header, so even a public file would be unreadable from
   JavaScript. This function does the reading server-side and hands the
   browser plain JSON from its own origin.

   GET /api/data        -> { tracker: [[...]], reqs: [[...]], fetchedAt, cached }
   GET /api/data?debug=1-> configuration check, no data

   Returns rows as arrays of DISPLAYED text, column A at index 0 — the same
   shape the gviz endpoint produced, so buildCands()/buildReqs() are unchanged.

   Required environment variables (Vercel → Settings → Environment Variables):
     MS_TENANT_ID       Directory (tenant) ID of the Azure AD app
     MS_CLIENT_ID       Application (client) ID
     MS_CLIENT_SECRET   Client secret VALUE (not the secret ID)
     MS_SHARE_URL       The SharePoint sharing link to HR_data.xlsx
   Optional:
     MS_TAB_TRACKER     default "Recruitment Tracker"
     MS_TAB_REQS        default "Sheet3"
     CACHE_SECONDS      default 60
--------------------------------------------------------------------------- */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const TIMEOUT_MS = 20000;

const cfg = () => ({
  tenant: process.env.MS_TENANT_ID,
  clientId: process.env.MS_CLIENT_ID,
  secret: process.env.MS_CLIENT_SECRET,
  shareUrl: process.env.MS_SHARE_URL,
  tabTracker: process.env.MS_TAB_TRACKER || 'Recruitment Tracker',
  tabReqs: process.env.MS_TAB_REQS || 'Sheet3',
  cacheSec: Math.max(0, parseInt(process.env.CACHE_SECONDS || '60', 10) || 60)
});

/* module-scope caches survive between invocations on a warm lambda */
let TOKEN = { value: null, exp: 0 };
let ITEM = { driveId: null, itemId: null, at: 0 };
let DATA = { payload: null, at: 0 };

async function token(c) {
  if (TOKEN.value && Date.now() < TOKEN.exp - 60000) return TOKEN.value;
  const body = new URLSearchParams({
    client_id: c.clientId,
    client_secret: c.secret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const r = await fetch('https://login.microsoftonline.com/' + c.tenant + '/oauth2/v2.0/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) {
    const e = new Error('Azure AD token request failed: ' + (j.error_description || j.error || r.status));
    e.step = 'token'; throw e;
  }
  TOKEN = { value: j.access_token, exp: Date.now() + (j.expires_in || 3600) * 1000 };
  return TOKEN.value;
}

/* a sharing link becomes a Graph share id: "u!" + unpadded base64url of the URL */
const shareId = url => 'u!' + Buffer.from(url, 'utf8').toString('base64')
  .replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');

async function resolveItem(c, tok) {
  if (ITEM.driveId && Date.now() - ITEM.at < 3600000) return ITEM;
  const r = await fetch(GRAPH + '/shares/' + shareId(c.shareUrl) + '/driveItem?$select=id,name,parentReference',
    { headers: { Authorization: 'Bearer ' + tok } });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.id) {
    const e = new Error('Could not resolve the shared file: ' + (j.error?.message || r.status)
      + ' — check MS_SHARE_URL and that the app has Files.Read.All granted.');
    e.step = 'resolve'; throw e;
  }
  ITEM = { driveId: j.parentReference?.driveId, itemId: j.id, name: j.name, at: Date.now() };
  return ITEM;
}

/* "Sheet1!C2:T811" -> 2  (zero-based index of the first used column) */
function colOffset(address) {
  const m = String(address || '').match(/!([A-Z]+)\d+/);
  if (!m) return 0;
  let n = 0;
  for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

async function sheet(c, tok, item, tab) {
  const url = GRAPH + '/drives/' + item.driveId + '/items/' + item.itemId
    + "/workbook/worksheets('" + encodeURIComponent(tab) + "')/usedRange(valuesOnly=true)?$select=text,address";
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + tok } });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) {
    const e = new Error('Tab "' + tab + '" could not be read: ' + (j.error?.message || r.status));
    e.step = 'sheet'; throw e;
  }
  /* "text" is the cell as DISPLAYED, which keeps dates as dd-mm-yyyy strings
     rather than Excel serial numbers — matching what gviz used to return. */
  const rows = j.text || [];
  const pad = colOffset(j.address);
  return pad ? rows.map(r0 => new Array(pad).fill('').concat(r0)) : rows;
}

module.exports = async (req, res) => {
  const c = cfg();
  res.setHeader('Cache-Control', 'no-store');

  const missing = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MS_SHARE_URL']
    .filter(k => !process.env[k]);

  if (req.query && req.query.debug) {
    const out = { configured: !missing.length, missingEnvVars: missing,
      tabs: { tracker: c.tabTracker, requisitions: c.tabReqs }, cacheSeconds: c.cacheSec };
    if (!missing.length) {
      try {
        const tok = await token(c);
        const item = await resolveItem(c, tok);
        out.tokenAcquired = true;
        out.file = { name: item.name, driveId: item.driveId ? item.driveId.slice(0, 12) + '…' : null };
      } catch (e) { out.tokenAcquired = false; out.failedAt = e.step || 'unknown'; out.error = e.message; }
    }
    return res.status(200).json(out);
  }

  if (missing.length) {
    return res.status(503).json({ error: 'not_configured', missingEnvVars: missing,
      message: 'Set these in Vercel → Settings → Environment Variables, then redeploy. Open /api/data?debug=1 to re-check.' });
  }

  if (DATA.payload && Date.now() - DATA.at < c.cacheSec * 1000) {
    return res.status(200).json({ ...DATA.payload, cached: true });
  }

  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);
  try {
    const tok = await token(c);
    const item = await resolveItem(c, tok);
    /* the requisition tab is optional — the dashboard works on the tracker alone */
    const [tr, rq] = await Promise.all([
      sheet(c, tok, item, c.tabTracker),
      sheet(c, tok, item, c.tabReqs).catch(e => ({ __error: e.message }))
    ]);
    const payload = {
      tracker: tr,
      reqs: Array.isArray(rq) ? rq : [],
      reqsError: Array.isArray(rq) ? null : rq.__error,
      fetchedAt: new Date().toISOString(),
      file: item.name
    };
    DATA = { payload, at: Date.now() };
    return res.status(200).json({ ...payload, cached: false });
  } catch (e) {
    const aborted = e && (e.name === 'AbortError' || /abort/i.test(String(e.message || '')));
    TOKEN = { value: null, exp: 0 };          // force a fresh token next time
    return res.status(aborted ? 504 : 502).json({
      error: aborted ? 'timeout' : (e.step || 'failed'),
      message: aborted ? 'SharePoint did not respond within 20 seconds.' : String(e.message || e)
    });
  } finally { clearTimeout(timer); }
};
