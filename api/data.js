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
const XLSX = require('xlsx');

const GRAPH = 'https://graph.microsoft.com/v1.0';
const TIMEOUT_MS = 45000;

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
let INFLIGHT = null;        // one shared download, however many callers arrive at once

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
   Excel stores dates as serial numbers. Converting them through any local-time
   Date path shifts them a day, which would silently corrupt every figure in the
   dashboard, so we do pure UTC arithmetic.
   25569 = days between the Excel epoch (1899-12-30) and the Unix epoch. */
const two = n => String(n).padStart(2, '0');
function serialToDMY(n) {
  const d = new Date(Math.round((n - 25569) * 86400000));
  if (isNaN(d)) return String(n);
  return two(d.getUTCDate()) + '/' + two(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
}

function cellText(c) {
  if (!c) return '';
  if (c.t === 'n') {
    const isDate = (c.z && XLSX.SSF && XLSX.SSF.is_date(c.z))
      || (!c.z && typeof c.w === 'string' && /^\d{1,4}[-\/]\d{1,2}[-\/]\d{1,4}$/.test(c.w.trim()));
    return isDate ? serialToDMY(c.v) : String(c.v);
  }
  if (c.t === 'd' && c.v instanceof Date) {
    return two(c.v.getUTCDate()) + '/' + two(c.v.getUTCMonth() + 1) + '/' + c.v.getUTCFullYear();
  }
  if (c.t === 'b') return c.v ? 'TRUE' : 'FALSE';
  return String(c.v == null ? '' : c.v).trim();
}

/* Walk the sheet by address so column A is always index 0, even when the used
   range starts further right. Trailing all-blank rows are dropped: stray
   formatting can stretch a sheet's declared range to thousands of empty rows. */
function gridOf(ws) {
  if (!ws || !ws['!ref']) return [];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const out = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    const row = [];
    for (let Cc = 0; Cc <= range.e.c; Cc++) {
      row.push(cellText(ws[XLSX.utils.encode_cell({ r: R, c: Cc })]));
    }
    out.push(row);
  }
  while (out.length && out[out.length - 1].every(v => v === '')) out.pop();
  return out;
}

function rowsOf(wb, wanted) {
  /* tolerate a renamed or re-cased tab rather than failing outright */
  const names = wb.SheetNames;
  const hit = names.find(n => n === wanted)
    || names.find(n => n.trim().toLowerCase() === wanted.trim().toLowerCase());
  if (!hit) return { error: 'tab "' + wanted + '" not found. Tabs present: ' + names.join(', ') };
  const rows = gridOf(wb.Sheets[hit]);
  return { name: hit, rows, address: wb.Sheets[hit]['!ref'] || null };
}

/* ---------- download the workbook ----------
   We deliberately do NOT use the Graph Excel workbook API (usedRange etc).
   It caps how many cells one request may touch and fails with
   "The cell count in the range has exceeded the maximum supported number"
   once a sheet grows or stray formatting inflates its used range. Downloading
   the file and parsing it locally has no such limit. */
async function download(base, tok, filePath, signal) {
  const url = base + '/root:/' + encodeURI(filePath) + ':/content';
  const r = await fetch(url, { signal, redirect: 'follow', headers: { Authorization: 'Bearer ' + tok } });
  if (!r.ok) {
    let detail = 'HTTP ' + r.status;
    try { detail = (JSON.parse(await r.text()).error || {}).message || detail; } catch {}
    throw fail('Could not download "' + filePath + '": ' + detail
      + ' — check FILE_PATH (case-sensitive, no leading slash) and that consent was granted.', 'download');
  }
  const buf = Buffer.from(await r.arrayBuffer());
  if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) {
    throw fail('The download did not return an .xlsx file (' + buf.length + ' bytes). '
      + 'Check FILE_PATH points at the workbook itself.', 'download');
  }
  return buf;
}

/* ---------- the whole read ---------- */
async function load(c, signal) {
  const tok = await getToken(c, signal);
  const base = await driveBase(c, tok, signal);
  const buf = await download(base, tok, c.filePath, signal);

  const wb = XLSX.read(buf, { type: 'buffer', cellNF: true, cellText: true });
  const tr = rowsOf(wb, c.sheet);
  if (tr.error) throw fail(tr.error, 'sheet');

  /* the requisitions tab is optional; the dashboard runs on the tracker alone */
  const rq = rowsOf(wb, c.sheetReqs);

  const tracker = tr.rows, reqs = rq.error ? [] : rq.rows;
  const hash = crypto.createHash('sha1')
    .update(JSON.stringify([tracker, reqs])).digest('hex').slice(0, 16);

  return {
    tracker, reqs,
    reqsError: rq.error || null,
    hash,
    tabsFound: wb.SheetNames,
    bytes: buf.length,
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

  /* The page polls every 5s but a cold download takes seconds. Without this,
     each poll starts its own download and they slow each other down until
     everything times out. One load runs; the rest await the same promise. */
  if (!INFLIGHT) {
    const ctl = new AbortController();
    const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);
    INFLIGHT = load(c, ctl.signal)
      .then(p => { CACHE = { payload: p, at: Date.now() }; return p; })
      .finally(() => { clearTimeout(timer); INFLIGHT = null; });
  }

  /* If we hold anything at all, serve it now and let the refresh finish in the
     background — a slow SharePoint must never blank the dashboard. */
  if (CACHE.payload) {
    INFLIGHT.catch(() => {});
    if (have && have === CACHE.payload.hash) return res.status(200).json(unchanged(CACHE.payload));
    return res.status(200).json({ ...CACHE.payload, cached: true, refreshing: true });
  }

  try {
    const payload = await INFLIGHT;

    if (debug) {
      return res.status(200).json({
        ok: true,
        mode: c.driveUser ? 'user drive (DRIVE_USER)' : 'sharepoint site (SITE_HOSTNAME + SITE_PATH)',
        site: c.driveUser || (c.host + c.sitePath),
        filePath: c.filePath,
        sheets: { tracker: c.sheet, requisitions: c.sheetReqs },
        tabsFound: payload.tabsFound,
        workbookBytes: payload.bytes,
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
      message: aborted
        ? 'Reading the workbook took longer than ' + (TIMEOUT_MS / 1000) + ' seconds. '
          + 'The file may have grown large — press Ctrl+End in Excel to see whether stray '
          + 'formatting has stretched the sheet far past the real data.'
        : String((e && e.message) || e)
    });
  }
};
