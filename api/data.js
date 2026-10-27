/* ---------------------------------------------------------------------------
   HR Dashboard — reads the workbook from Microsoft 365 via the Graph API.

   App-only (client credentials): the function signs in as an Azure AD
   application, so the file can stay PRIVATE in SharePoint/OneDrive. Nothing is
   shared publicly and no credential is ever sent to the browser.

   GET /api/data          -> { tracker: [[...]], reqs: [[...]], hash, fetchedAt }
   GET /api/data?v=<hash> -> { unchanged: true } when nothing moved (cheap poll)
   GET /api/data?debug=1  -> configuration / connection check, no row data

   Rows are arrays of strings with column A at index 0 — the same shape the old
   loader produced, so buildCands()/buildReqs() are unchanged.

   Environment variables (Vercel -> Settings -> Environment Variables):
     TENANT_ID        Directory (tenant) ID
     CLIENT_ID        Application (client) ID
     CLIENT_SECRET    Client secret VALUE (not the secret ID)
     SITE_HOSTNAME    e.g. elixirengineering.sharepoint.com
     SITE_PATH        e.g. /sites/HR          (site-relative path)
     FILE_PATH        e.g. Shared Documents/HR_data.xlsx   (path inside the drive)
     SHEET_NAME       e.g. Recruitment Tracker
   Also used:
     SHEET_NAME_REQS  requisitions tab, default "Sheet3" (the dashboard reads two
                      tabs; set this if yours is named differently)
     DRIVE_USER       use INSTEAD of SITE_HOSTNAME/SITE_PATH when the file lives on
                      a person's OneDrive rather than a SharePoint site, e.g.
                      aliasgark@elixirengg.com
     CACHE_SECONDS    default 60

   Graph permission required (Application, with admin consent):
     Sites.Read.All   when using SITE_HOSTNAME + SITE_PATH
     Files.Read.All   when using DRIVE_USER
--------------------------------------------------------------------------- */

const crypto = require('crypto');

const GRAPH = 'https://graph.microsoft.com/v1.0';
const TIMEOUT_MS = 25000;

const env = k => (process.env[k] || '').trim();
const CFG = () => ({
  tenant: env('TENANT_ID'),
  clientId: env('CLIENT_ID'),
  secret: env('CLIENT_SECRET'),
  host: env('SITE_HOSTNAME'),
  sitePath: env('SITE_PATH'),
  filePath: env('FILE_PATH'),
  sheet: env('SHEET_NAME') || 'Recruitment Tracker',
  sheetReqs: env('SHEET_NAME_REQS') || 'Sheet3',
  driveUser: env('DRIVE_USER'),
  cacheMs: Math.max(0, parseInt(env('CACHE_SECONDS') || '60', 10) || 60) * 1000
});

/* module-scope caches survive between invocations on a warm lambda */
let TOKEN = { value: null, exp: 0 };
let DRIVE = { base: null, at: 0 };
let CACHE = { payload: null, at: 0 };

const fail = (msg, step) => { const e = new Error(msg); e.step = step; return e; };

/* ---------- auth ---------- */
async function getToken(c, signal) {
  if (TOKEN.value && Date.now() < TOKEN.exp - 60000) return TOKEN.value;
  const r = await fetch('https://login.microsoftonline.com/' + encodeURIComponent(c.tenant) + '/oauth2/v2.0/token', {
    method: 'POST', signal,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: c.clientId,
      client_secret: c.secret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    })
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) {
    throw fail('Sign-in to Microsoft failed: ' + (j.error_description || j.error || ('HTTP ' + r.status))
      + ' — check TENANT_ID, CLIENT_ID and CLIENT_SECRET (the secret VALUE, not its ID).', 'token');
  }
  TOKEN = { value: j.access_token, exp: Date.now() + ((j.expires_in || 3600) * 1000) };
  return TOKEN.value;
}

async function gget(url, tok, signal) {
  const r = await fetch(url, { signal, headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json' } });
  const txt = await r.text();
  let j = null; try { j = JSON.parse(txt); } catch { /* non-JSON body */ }
  return { ok: r.ok, status: r.status, json: j, text: txt };
}

/* ---------- locate the drive that holds the file ---------- */
async function driveBase(c, tok, signal) {
  if (DRIVE.base && Date.now() - DRIVE.at < 3600000) return DRIVE.base;

  let base;
  if (c.driveUser) {
    base = GRAPH + '/users/' + encodeURIComponent(c.driveUser) + '/drive';
  } else {
    const path = c.sitePath.replace(/^\/+|\/+$/g, '');
    const url = GRAPH + '/sites/' + encodeURIComponent(c.host) + ':/' + path;
    const r = await gget(url, tok, signal);
    if (!r.ok || !r.json || !r.json.id) {
      throw fail('Could not find the SharePoint site "' + c.host + c.sitePath + '": '
        + ((r.json && r.json.error && r.json.error.message) || ('HTTP ' + r.status))
        + ' — check SITE_HOSTNAME and SITE_PATH, and that Sites.Read.All has admin consent.', 'site');
    }
    base = GRAPH + '/sites/' + r.json.id + '/drive';
  }
  DRIVE = { base, at: Date.now() };
  return base;
}

/* ---------- dates ----------
   Graph hands back Excel date cells as serial numbers. Converting them through
   any local-time Date path shifts them a day, which would silently corrupt every
   figure in the dashboard, so we do pure UTC arithmetic.
   25569 = days between the Excel epoch (1899-12-30) and the Unix epoch. */
const two = n => String(n).padStart(2, '0');
function serialToDMY(n) {
  const d = new Date(Math.round((n - 25569) * 86400000));
  if (isNaN(d)) return String(n);
  return two(d.getUTCDate()) + '/' + two(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
}
/* "dd\-mm\-yyyy" and "m/d/yy" are dates; "General", "0.00", "#,##0" are not */
const isDateFmt = f => typeof f === 'string' && !/^general$/i.test(f.trim())
  && /y/i.test(f) && /[dm]/i.test(f) && !/[#?]/.test(f);

/* "Recruitment Tracker!C3:T1153" -> 2 (zero-based first used column) */
function colOffset(address) {
  const m = String(address || '').match(/!\$?([A-Z]+)\$?\d+/i);
  if (!m) return 0;
  let n = 0;
  for (const ch of m[1].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/* ---------- read one worksheet ---------- */
async function sheetRows(base, tok, name, signal) {
  const url = base + "/root:/" + encodeURI(CFG().filePath) + ":/workbook/worksheets('"
    + encodeURIComponent(name) + "')/usedRange(valuesOnly=true)"
    + "?$select=address,text,values,numberFormat";
  const r = await gget(url, tok, signal);
  if (!r.ok) {
    const m = (r.json && r.json.error && r.json.error.message) || ('HTTP ' + r.status);
    return { error: 'worksheet "' + name + '" could not be read: ' + m };
  }
  const text = (r.json && r.json.text) || [];
  const vals = (r.json && r.json.values) || [];
  const fmts = (r.json && r.json.numberFormat) || [];
  const pad = colOffset(r.json && r.json.address);

  const rows = text.map((row, i) => {
    const out = new Array(pad).fill('');
    for (let c = 0; c < row.length; c++) {
      const v = vals[i] ? vals[i][c] : null;
      const f = fmts[i] ? fmts[i][c] : null;
      if (typeof v === 'number' && isDateFmt(f)) out.push(serialToDMY(v));
      else if (row[c] !== '' && row[c] != null) out.push(String(row[c]).trim());
      else out.push(v == null ? '' : String(v).trim());
    }
    return out;
  });
  return { rows, address: r.json && r.json.address };
}

/* ---------- the whole read ---------- */
async function load(c, signal) {
  const tok = await getToken(c, signal);
  const base = await driveBase(c, tok, signal);

  const tr = await sheetRows(base, tok, c.sheet, signal);
  if (tr.error) throw fail(tr.error + ' — check FILE_PATH and SHEET_NAME.', 'sheet');

  /* the requisitions tab is optional; the dashboard runs on the tracker alone */
  const rq = await sheetRows(base, tok, c.sheetReqs, signal);

  const tracker = tr.rows, reqs = rq.error ? [] : rq.rows;
  const hash = crypto.createHash('sha1')
    .update(JSON.stringify([tracker, reqs])).digest('hex').slice(0, 16);

  return {
    tracker, reqs,
    reqsError: rq.error || null,
    hash,
    addresses: { tracker: tr.address, reqs: rq.address || null },
    fetchedAt: new Date().toISOString()
  };
}

/* ---------- handler ---------- */
module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  const c = CFG();
  const debug = req.query && req.query.debug;
  const have = req.query && req.query.v;
  const unchanged = p => ({ unchanged: true, hash: p.hash, fetchedAt: p.fetchedAt });

  const required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'FILE_PATH', 'SHEET_NAME'];
  const missing = required.filter(k => !env(k));
  if (!c.driveUser && (!c.host || !c.sitePath)) missing.push('SITE_HOSTNAME + SITE_PATH (or DRIVE_USER)');

  if (missing.length) {
    return res.status(503).json({
      error: 'not_configured', missingEnvVars: missing,
      message: 'Set these in Vercel -> Settings -> Environment Variables, then redeploy.'
    });
  }

  /* a poll inside the cache window costs ~60 bytes when nothing has changed */
  if (CACHE.payload && Date.now() - CACHE.at < c.cacheMs && !debug) {
    if (have && have === CACHE.payload.hash) return res.status(200).json(unchanged(CACHE.payload));
    return res.status(200).json({ ...CACHE.payload, cached: true });
  }

  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);
  try {
    const payload = await load(c, ctl.signal);
    CACHE = { payload, at: Date.now() };

    if (debug) {
      return res.status(200).json({
        ok: true,
        mode: c.driveUser ? 'user drive (DRIVE_USER)' : 'sharepoint site (SITE_HOSTNAME + SITE_PATH)',
        site: c.driveUser || (c.host + c.sitePath),
        filePath: c.filePath,
        sheets: { tracker: c.sheet, requisitions: c.sheetReqs },
        addresses: payload.addresses,
        trackerRows: payload.tracker.length,
        requisitionRows: payload.reqs.length,
        reqsError: payload.reqsError,
        hash: payload.hash,
        cacheSeconds: c.cacheMs / 1000,
        /* first rows of each tab, so column positions can be eyeballed —
           the parser reads by position, so a column inserted mid-sheet shifts
           every field after it without raising any error */
        trackerFirstRows: payload.tracker.slice(0, 5),
        reqsFirstRows: payload.reqs.slice(0, 3),
        fetchedAt: payload.fetchedAt
      });
    }
    if (have && have === payload.hash) return res.status(200).json(unchanged(payload));
    return res.status(200).json({ ...payload, cached: false });

  } catch (e) {
    const aborted = e && (e.name === 'AbortError' || /abort/i.test(String(e.message || '')));
    if (e && e.step === 'token') TOKEN = { value: null, exp: 0 };
    if (e && e.step === 'site') DRIVE = { base: null, at: 0 };
    return res.status(aborted ? 504 : 502).json({
      error: aborted ? 'timeout' : (e && e.step) || 'failed',
      failedAt: (e && e.step) || null,
      message: aborted ? 'Microsoft Graph did not respond within 25 seconds.'
        : String((e && e.message) || e)
    });
  } finally { clearTimeout(timer); }
};
