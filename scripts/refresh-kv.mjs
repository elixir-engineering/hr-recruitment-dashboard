/* ---------------------------------------------------------------------------
   Build the dashboard payload and write it to Cloudflare KV.

   Runs in GitHub Actions on a schedule, NOT in a request. That is the whole
   point: parsing the workbook costs ~103 ms, and a Cloudflare Function gets
   10 ms of CPU. Here it is free.

   This must produce byte-identical output to the Vercel route in api/data.js,
   otherwise the two hosts would disagree. The date, cell and grid logic below
   is a deliberate copy of that file — if you change one, change both.

   Writes to KV, never to the repo: a commit per refresh would trigger a Pages
   rebuild, and the free tier allows 500 builds/month (a 15-minute schedule
   alone would be ~2,880).

   Env (all from GitHub secrets):
     TENANT_ID CLIENT_ID CLIENT_SECRET
     DRIVE_USER | (SITE_HOSTNAME + SITE_PATH)
     FILE_PATH SHEET_NAME [SHEET_NAME_REQS]
     CF_ACCOUNT_ID CF_KV_NAMESPACE_ID CF_API_TOKEN
--------------------------------------------------------------------------- */

import crypto from 'node:crypto';
import XLSX from 'xlsx';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const env = k => (process.env[k] || '').trim();

const CFG = {
  tenant: env('TENANT_ID'),
  clientId: env('CLIENT_ID'),
  secret: env('CLIENT_SECRET'),
  host: env('SITE_HOSTNAME'),
  sitePath: env('SITE_PATH'),
  driveUser: env('DRIVE_USER'),
  filePath: env('FILE_PATH'),
  sheet: env('SHEET_NAME') || 'Recruitment Tracker',
  sheetReqs: env('SHEET_NAME_REQS') || 'Sheet3',
  cfAccount: env('CF_ACCOUNT_ID'),
  cfNamespace: env('CF_KV_NAMESPACE_ID'),
  cfToken: env('CF_API_TOKEN')
};

function requireEnv() {
  const need = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'FILE_PATH', 'SHEET_NAME',
    'CF_ACCOUNT_ID', 'CF_KV_NAMESPACE_ID', 'CF_API_TOKEN'];
  const missing = need.filter(k => !env(k));
  if (!CFG.driveUser && !(CFG.host && CFG.sitePath)) missing.push('DRIVE_USER (or SITE_HOSTNAME + SITE_PATH)');
  if (missing.length) {
    console.error('Missing secrets: ' + missing.join(', '));
    process.exit(1);
  }
}

/* ---------- Microsoft Graph ---------- */
async function getToken() {
  const r = await fetch('https://login.microsoftonline.com/' + encodeURIComponent(CFG.tenant) + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CFG.clientId, client_secret: CFG.secret,
      scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials'
    })
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) {
    throw new Error('Azure AD token failed: ' + (j.error_description || j.error || r.status));
  }
  return j.access_token;
}

async function driveBase(tok) {
  if (CFG.driveUser) return GRAPH + '/users/' + encodeURIComponent(CFG.driveUser) + '/drive';
  const path = CFG.sitePath.replace(/^\/+|\/+$/g, '');
  const r = await fetch(GRAPH + '/sites/' + encodeURIComponent(CFG.host) + ':/' + path,
    { headers: { Authorization: 'Bearer ' + tok } });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.id) throw new Error('Site lookup failed: ' + ((j.error || {}).message || r.status));
  return GRAPH + '/sites/' + j.id + '/drive';
}

/* FILE_PATH is a GitHub secret, so its value is masked as *** in the logs and
   a wrong value is invisible. Accept the shapes people actually paste — a full
   SharePoint URL, a leading slash, a ?web=1 suffix, %20 escapes — and print a
   masked-safe description of what we ended up with. */
function normalisePath(raw) {
  let s = String(raw || '').trim();
  const notes = [];
  if (/^https?:\/\//i.test(s)) {
    notes.push('stripped full URL');
    const m = s.match(/\/(?:personal|sites)\/[^/]+\/(?:Documents|Shared%20Documents|Shared Documents)\/(.+)$/i);
    s = m ? m[1] : s.split('/').pop();
  }
  if (s.includes('?')) { notes.push('stripped query string'); s = s.split('?')[0]; }
  if (s.includes('#')) { s = s.split('#')[0]; }
  if (/^\/+/.test(s)) { notes.push('stripped leading slash'); s = s.replace(/^\/+/, ''); }
  if (/%[0-9a-f]{2}/i.test(s)) { notes.push('decoded %-escapes'); try { s = decodeURIComponent(s); } catch { /* leave as-is */ } }
  return { path: s, notes };
}

async function download(base, tok) {
  const { path, notes } = normalisePath(CFG.filePath);
  /* describe the path without printing it — GitHub masks the secret anyway */
  console.log('FILE_PATH check: ' + path.length + ' chars, '
    + (path.split('/').length - 1) + ' folder level(s), '
    + (/\.xlsx$/i.test(path) ? 'ends .xlsx' : 'DOES NOT end in .xlsx')
    + (notes.length ? ' · fixed up: ' + notes.join(', ') : ''));

  const r = await fetch(base + '/root:/' + encodeURI(path) + ':/content',
    { headers: { Authorization: 'Bearer ' + tok }, redirect: 'follow' });
  if (!r.ok) {
    let d = 'HTTP ' + r.status;
    try { d = (JSON.parse(await r.text()).error || {}).message || d; } catch { /* non-JSON */ }
    throw new Error('Download failed (' + d + '). FILE_PATH must be the path inside the drive '
      + 'with no leading slash and no ?web=1 — for a file in the OneDrive root that is just '
      + '"HR_data.xlsx". It is case-sensitive. Set LIST_FILES=1 in the workflow to print the '
      + 'files the app can actually see.');
  }
  const buf = Buffer.from(await r.arrayBuffer());
  if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) {
    throw new Error('Not an .xlsx (got ' + buf.length + ' bytes). Check FILE_PATH.');
  }
  return buf;
}

/* ---------- parsing (must match api/data.js exactly) ----------
   Excel dates are serial numbers. Any local-time conversion shifts them a day
   and silently corrupts every figure. 25569 = Excel epoch to Unix epoch. */
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
      || (!c.z && typeof c.w === 'string' && /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/.test(c.w.trim()));
    return isDate ? serialToDMY(c.v) : String(c.v);
  }
  if (c.t === 'd' && c.v instanceof Date) {
    return two(c.v.getUTCDate()) + '/' + two(c.v.getUTCMonth() + 1) + '/' + c.v.getUTCFullYear();
  }
  if (c.t === 'b') return c.v ? 'TRUE' : 'FALSE';
  return String(c.v == null ? '' : c.v).trim();
}

/* Find the real extent of the data by looking at the cells that actually exist
   and hold a value, rather than trusting the sheet's declared !ref.

   This matters enormously. Stray formatting stretches !ref far past the real
   data — on this workbook a 570 KB file declared a range big enough to produce
   a 71 MB payload of mostly empty strings, which blew the 25 MiB KV limit and
   took 10 s to build. Measuring the true extent gives ~0.5 MB. */
function realExtent(ws) {
  let maxR = -1, maxC = -1;
  for (const key in ws) {
    if (key.charCodeAt(0) === 33) continue;            // '!' metadata keys
    const c = ws[key];
    if (!c || c.v == null || String(c.v).trim() === '') continue;
    const a = XLSX.utils.decode_cell(key);
    if (a.r > maxR) maxR = a.r;
    if (a.c > maxC) maxC = a.c;
  }
  return { maxR, maxC };
}

/* Walk by address so column A is always index 0. */
function gridOf(ws) {
  if (!ws || !ws['!ref']) return [];
  const { maxR, maxC } = realExtent(ws);
  if (maxR < 0) return [];
  const out = [];
  for (let R = 0; R <= maxR; R++) {
    const row = [];
    for (let C = 0; C <= maxC; C++) {
      row.push(cellText(ws[XLSX.utils.encode_cell({ r: R, c: C })]));
    }
    out.push(row);
  }
  while (out.length && out[out.length - 1].every(v => v === '')) out.pop();
  return out;
}

function rowsOf(wb, wanted) {
  const names = wb.SheetNames;
  const hit = names.find(n => n === wanted)
    || names.find(n => n.trim().toLowerCase() === wanted.trim().toLowerCase());
  if (!hit) return { error: 'tab "' + wanted + '" not found. Tabs present: ' + names.join(', ') };
  return { name: hit, rows: gridOf(wb.Sheets[hit]), address: wb.Sheets[hit]['!ref'] || null };
}

/* ---------- Cloudflare KV ---------- */
async function kvPut(key, value, metadata) {
  const url = 'https://api.cloudflare.com/client/v4/accounts/' + CFG.cfAccount
    + '/storage/kv/namespaces/' + CFG.cfNamespace + '/values/' + encodeURIComponent(key);
  const fd = new FormData();
  fd.set('value', value);
  fd.set('metadata', JSON.stringify(metadata || {}));
  const r = await fetch(url, { method: 'PUT', headers: { Authorization: 'Bearer ' + CFG.cfToken }, body: fd });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || j.success === false) {
    throw new Error('KV write "' + key + '" failed: ' + JSON.stringify(j.errors || j).slice(0, 300));
  }
}

/* ---------- main ---------- */
/* LIST_FILES=1 -> print what the app can see, so a wrong FILE_PATH can be
   corrected from evidence instead of guesswork. Filenames only, no contents. */
async function listFiles(base, tok) {
  const r = await fetch(base + '/root/children?$select=name,size,folder&$top=100',
    { headers: { Authorization: 'Bearer ' + tok } });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) { console.log('Could not list the drive root: ' + ((j.error || {}).message || r.status)); return; }
  console.log('--- files in the drive root ---');
  for (const it of (j.value || [])) {
    console.log('  ' + (it.folder ? '[folder] ' : '         ') + it.name
      + (it.size ? '  (' + Math.round(it.size / 1024) + ' KB)' : ''));
  }
  console.log('--- use one of these names as FILE_PATH (add "Folder/" if it is inside one) ---');
}

async function main() {
  requireEnv();
  const t0 = Date.now();

  const tok = await getToken();
  const base = await driveBase(tok);
  if (env('LIST_FILES')) await listFiles(base, tok);
  const buf = await download(base, tok);
  console.log('downloaded ' + (buf.length / 1024).toFixed(0) + ' KB in ' + (Date.now() - t0) + ' ms');

  const tParse = Date.now();
  const wb = XLSX.read(buf, { type: 'buffer', cellNF: true, cellText: true });
  const tr = rowsOf(wb, CFG.sheet);
  if (tr.error) throw new Error(tr.error);
  const rq = rowsOf(wb, CFG.sheetReqs);
  console.log('parsed in ' + (Date.now() - tParse) + ' ms');

  const tracker = tr.rows;
  const reqs = rq.error ? [] : rq.rows;
  console.log('tracker: declared ' + (tr.address || '?') + ' -> real '
    + tracker.length + ' rows x ' + (tracker[0] ? tracker[0].length : 0) + ' cols');
  console.log('reqs   : declared ' + (rq.address || '?') + ' -> real '
    + reqs.length + ' rows x ' + (reqs[0] ? reqs[0].length : 0) + ' cols');
  const hash = crypto.createHash('sha1').update(JSON.stringify([tracker, reqs])).digest('hex').slice(0, 16);
  const fetchedAt = new Date().toISOString();

  /* Exactly the shape api/data.js returns, so the front end cannot tell them
     apart apart from the x-backend header. */
  const payload = JSON.stringify({
    tracker, reqs,
    reqsError: rq.error || null,
    hash, fetchedAt,
    tabsFound: wb.SheetNames,
    bytes: buf.length,
    addresses: { tracker: tr.address, reqs: rq.address || null },
    cached: false, ageSeconds: 0
  });

  /* KV metadata is capped at 1024 bytes — keep it to what the function needs
     to answer an unchanged poll without reading the payload. */
  const metadata = { hash, fetchedAt, trackerRows: tracker.length, reqsRows: reqs.length };

  /* The small key ?debug=1 reads. Separate so debug never touches the big one. */
  const meta = {
    ...metadata,
    workbookBytes: buf.length,
    tabsFound: wb.SheetNames,
    sheets: { tracker: CFG.sheet, requisitions: CFG.sheetReqs },
    addresses: { tracker: tr.address, reqs: rq.address || null },
    reqsError: rq.error || null,
    payloadBytes: payload.length,
    trackerFirstRows: tracker.slice(0, 5),
    reqsFirstRows: reqs.slice(0, 3)
  };

  /* KV caps a value at 25 MiB. Fail loudly rather than sending a doomed write. */
  const MAX = 24 * 1024 * 1024;
  if (payload.length > MAX) {
    throw new Error('Payload is ' + (payload.length / 1048576).toFixed(1) + ' MB, over the 25 MiB KV limit. '
      + 'That means the sheet still has a huge amount of real data or stray values far from the table. '
      + 'In Excel press Ctrl+End to see where it thinks the data ends, delete the empty rows/columns '
      + 'below and right of the real table, and save.');
  }

  await kvPut('payload', payload, metadata);
  await kvPut('meta', JSON.stringify(meta), { hash, fetchedAt });

  console.log('wrote payload ' + (payload.length / 1024).toFixed(0) + ' KB · '
    + tracker.length + ' tracker rows · ' + reqs.length + ' requisition rows · hash ' + hash);
  console.log('total ' + (Date.now() - t0) + ' ms');
}

main().catch(e => { console.error('FAILED: ' + (e && e.message || e)); process.exit(1); });
