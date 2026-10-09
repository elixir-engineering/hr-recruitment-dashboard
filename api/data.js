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

/* A SharePoint sharing link serves an HTML viewer by default; download=1 makes
   it serve the file. We follow redirects and then check we really got a file. */
const directUrl = u => u + (u.includes('?') ? '&' : '?') + 'download=1';

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
    const r = await fetch(directUrl(SHARE_URL), {
      signal: ctl.signal, redirect: 'follow',
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; HRDashboard/1.0)' }
    });
    if (!r.ok) throw new Error('SharePoint returned HTTP ' + r.status + ' for the sharing link.');

    const buf = Buffer.from(await r.arrayBuffer());

    /* Every .xlsx is a zip, so it starts with "PK". Anything else means we were
       handed a sign-in page instead of the file. */
    if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) {
      const head = buf.slice(0, 400).toString('utf8');
      const signIn = /sign ?in|login|redirect|authenticat/i.test(head);
      throw new Error(signIn
        ? 'SharePoint served a sign-in page instead of the file. The sharing link is restricted. '
          + 'In SharePoint open HR_data.xlsx → Share → the gear/settings on the link → set it to '
          + '"Anyone with the link" (Can view), then use that new link.'
        : 'The sharing link did not return an .xlsx file (got ' + buf.length + ' bytes starting "'
          + head.slice(0, 40).replace(/\s+/g, ' ') + '").');
    }

    const wb = XLSX.read(buf, { type: 'buffer', cellNF: true, cellText: true });
    const tr = rowsOf(wb, TAB_TRACKER);
    const rq = rowsOf(wb, TAB_REQS);
    if (tr.error) throw new Error(tr.error);

    return {
      tracker: tr.rows,
      reqs: rq.error ? [] : rq.rows,
      reqsError: rq.error || null,
      tabsFound: wb.SheetNames,
      bytes: buf.length,
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
        ok: true, url: SHARE_URL.split('?')[0] + '?…', bytes: payload.bytes,
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
      url: SHARE_URL.split('?')[0] + '?…'
    });
  }
};
