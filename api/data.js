/* ---------------------------------------------------------------------------
   HR Dashboard — reads HR_data.xlsx from SharePoint. No configuration needed.

   Why this file exists at all: the browser cannot fetch the workbook itself,
   because SharePoint sends no Access-Control-Allow-Origin header. CORS is a
   browser rule, not a server one — so this function downloads the file
   server-side and hands the browser plain JSON from its own origin.

   The only requirement is that the sharing link is set to
   "Anyone with the link" in SharePoint. No Azure app, no keys, no env vars.

   GET /api/data         -> { tracker: [[...]], reqs: [[...]], fetchedAt }
   GET /api/data?debug=1 -> what it managed to read, without the row data

   Rows come back as arrays of strings, column A at index 0 — the same shape the
   old Google gviz endpoint produced, so buildCands()/buildReqs() are unchanged.
--------------------------------------------------------------------------- */

const XLSX = require('xlsx');

/* Override with HR_XLSX_URL if the file ever moves. */
const SHARE_URL = process.env.HR_XLSX_URL ||
  'https://elixirengineering-my.sharepoint.com/:x:/g/personal/aliasgark_elixirengg_com/IQCXhg2hv-SoUK7mASENzSQvAR5IbolCEfJkzhBzrE7VFbc?e=PWaljg';

const TAB_TRACKER = process.env.HR_TAB_TRACKER || 'Recruitment Tracker';
const TAB_REQS = process.env.HR_TAB_REQS || 'Sheet3';
const CACHE_MS = Math.max(0, parseInt(process.env.CACHE_SECONDS || '60', 10) || 60) * 1000;
const TIMEOUT_MS = 25000;

let CACHE = { payload: null, at: 0 };

/* Getting the bytes of an anonymously-shared SharePoint file is fiddly: a plain
   sharing URL serves an HTML viewer, and ?download=1 alone is not always enough
   on /:x:/g/personal/ links. So we try several known forms and take the first
   that actually returns a zip (every .xlsx starts with "PK").

   The first one is the important one: api.onedrive.com resolves a share token
   with no authentication at all, provided the link is "Anyone with the link".
   The share id is "u!" + unpadded base64url of the full sharing URL. */
const shareId = u => 'u!' + Buffer.from(u, 'utf8').toString('base64')
  .replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');

function strategies(u) {
  const out = [];
  const id = shareId(u);
  out.push({ name: 'onedrive-shares-root', url: 'https://api.onedrive.com/v1.0/shares/' + id + '/root/content' });
  out.push({ name: 'onedrive-shares-driveitem', url: 'https://api.onedrive.com/v1.0/shares/' + id + '/driveItem/content' });
  out.push({ name: 'download-param', url: u + (u.includes('?') ? '&' : '?') + 'download=1' });
  if (u.includes('/:x:/')) {
    const generic = u.replace('/:x:/', '/:u:/');
    out.push({ name: 'generic-u-download', url: generic + (generic.includes('?') ? '&' : '?') + 'download=1' });
  }
  /* host/personal/<user>/_layouts/15/download.aspx?share=<token> */
  const m = u.match(/^(https:\/\/[^/]+)\/:[a-z]:\/[a-z]\/(personal\/[^/]+)\/([^?]+)/i);
  if (m) out.push({ name: 'layouts-download-aspx', url: m[1] + '/' + m[2] + '/_layouts/15/download.aspx?share=' + m[3] });
  return out;
}

const isZip = b => b.length > 4 && b[0] === 0x50 && b[1] === 0x4b;

async function grab(url, signal) {
  const r = await fetch(url, {
    signal, redirect: 'follow',
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36',
      'Accept': 'application/octet-stream,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*'
    }
  });
  const buf = Buffer.from(await r.arrayBuffer());
  return {
    status: r.status, ok: r.ok, bytes: buf.length,
    contentType: r.headers.get('content-type') || '',
    zip: isZip(buf),
    head: isZip(buf) ? 'PK (xlsx)' : buf.slice(0, 120).toString('utf8').replace(/\s+/g, ' ').trim(),
    buf
  };
}

/* Tries each strategy, returns the buffer from the first that yields a zip.
   The attempt log comes back either way so ?debug=1 can show what happened. */
async function fetchWorkbook(signal) {
  const log = [];
  for (const s of strategies(SHARE_URL)) {
    try {
      const r = await grab(s.url, signal);
      log.push({ strategy: s.name, status: r.status, bytes: r.bytes, contentType: r.contentType, got: r.head.slice(0, 90) });
      if (r.zip) return { buf: r.buf, via: s.name, log };
    } catch (e) {
      log.push({ strategy: s.name, error: String((e && e.message) || e) });
      if (e && e.name === 'AbortError') throw e;
    }
  }
  return { buf: null, via: null, log };
}

/* The dashboard's pdate() expects dd/mm/yyyy.

   Dates are the one thing that can silently corrupt this whole dashboard, so we
   do NOT let SheetJS hand us Date objects: cellDates:true converts through the
   server's local timezone and round-trips a day early (9 Jul comes back as 8
   Jul). Instead we take the raw Excel serial and convert it with pure UTC
   arithmetic, which no timezone can touch. 25569 = days between the Excel
   epoch (1899-12-30) and the Unix epoch. */
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
  if (c.t === 'd' && c.v instanceof Date) {            // only if some path yields a Date
    return two(c.v.getUTCDate()) + '/' + two(c.v.getUTCMonth() + 1) + '/' + c.v.getUTCFullYear();
  }
  if (c.t === 'b') return c.v ? 'TRUE' : 'FALSE';
  return String(c.v == null ? '' : c.v).trim();
}

/* Walk the sheet by address rather than using sheet_to_json, so column A is
   always index 0 even when the used range starts further right. */
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
  return out;
}

function rowsOf(wb, wanted) {
  /* tolerate a renamed or re-cased tab rather than failing outright */
  const names = wb.SheetNames;
  const hit = names.find(n => n === wanted)
    || names.find(n => n.trim().toLowerCase() === wanted.trim().toLowerCase());
  if (!hit) return { error: 'tab "' + wanted + '" not found. Tabs present: ' + names.join(', ') };
  return { name: hit, rows: gridOf(wb.Sheets[hit]) };
}

async function load() {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);
  try {
    const { buf, via, log } = await fetchWorkbook(ctl.signal);
    if (!buf) {
      const e = new Error('None of the download methods returned the file. '
        + 'Every attempt came back as a web page, which means the link still needs a sign-in. '
        + 'Check the link opens in a private/incognito window without logging in.');
      e.attempts = log;
      throw e;
    }

    const wb = XLSX.read(buf, { type: 'buffer', cellNF: true, cellText: true });
    const tr = rowsOf(wb, TAB_TRACKER);
    const rq = rowsOf(wb, TAB_REQS);
    if (tr.error) { const e = new Error(tr.error); e.attempts = log; throw e; }

    return {
      tracker: tr.rows,
      reqs: rq.error ? [] : rq.rows,
      reqsError: rq.error || null,
      tabsFound: wb.SheetNames,
      bytes: buf.length,
      via,
      attempts: log,
      fetchedAt: new Date().toISOString()
    };
  } finally { clearTimeout(timer); }
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  const debug = req.query && req.query.debug;

  if (CACHE.payload && Date.now() - CACHE.at < CACHE_MS && !debug) {
    return res.status(200).json({ ...CACHE.payload, cached: true });
  }

  try {
    const payload = await load();
    CACHE = { payload, at: Date.now() };
    if (debug) {
      return res.status(200).json({
        ok: true, via: payload.via, attempts: payload.attempts,
        url: SHARE_URL.split('?')[0] + '?…', bytes: payload.bytes,
        tabsFound: payload.tabsFound,
        lookingFor: { tracker: TAB_TRACKER, requisitions: TAB_REQS },
        trackerRows: payload.tracker.length, requisitionRows: payload.reqs.length,
        reqsError: payload.reqsError,
        firstTrackerRow: payload.tracker[0], secondTrackerRow: payload.tracker[1],
        fetchedAt: payload.fetchedAt
      });
    }
    return res.status(200).json({ ...payload, cached: false });
  } catch (e) {
    const aborted = e && (e.name === 'AbortError' || /abort/i.test(String(e.message || '')));
    return res.status(aborted ? 504 : 502).json({
      error: aborted ? 'timeout' : 'failed',
      message: aborted ? 'SharePoint did not respond within 25 seconds.' : String((e && e.message) || e),
      attempts: e && e.attempts || null,
      url: SHARE_URL.split('?')[0] + '?…'
    });
  }
};
