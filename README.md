# HR Recruitment Dashboard

Single-page dashboard over `HR_data.xlsx` on Microsoft 365, deployed on Vercel.

- `index.html` — the whole dashboard. Reads live data from `/api/data`, falls back to an embedded snapshot.
- `api/data.js` — signs in to Microsoft and reads the workbook through the Graph API.
- `package.json` — no dependencies; the route uses plain `fetch`.

## How it reads the data

`HR_data.xlsx` is **private** on Ali Asgar's OneDrive. The route signs in as an Azure AD
application (client credentials) and reads it through Microsoft Graph, so no sharing link
exists and no credential ever reaches the browser.

### Live configuration

Set in Vercel → Settings → Environment Variables:

| Name | Value |
|---|---|
| `TENANT_ID` | Directory (tenant) ID |
| `CLIENT_ID` | Application (client) ID |
| `CLIENT_SECRET` | client secret **Value** |
| `DRIVE_USER` | `aliasgark@elixirengg.com` |
| `FILE_PATH` | `HR_data.xlsx` |
| `SHEET_NAME` | `Recruitment Tracker` |

Optional: `SHEET_NAME_REQS` (default `Sheet3`), `CACHE_SECONDS` (default `60`).

`SITE_HOSTNAME` / `SITE_PATH` are deliberately **unset** — setting them switches the route
into SharePoint-site mode, which cannot find a file on a personal OneDrive. They exist for
if the workbook ever moves into a SharePoint document library.

### The Azure app

Registered as `HR Dashboard Reader`, single tenant, with the Microsoft Graph **application**
permission `Files.Read.All` and tenant admin consent granted. Read-only on purpose: the
dashboard never writes, and `Files.ReadWrite.All` would mean a leaked secret could alter any
file in the tenant.

Note the app has **no redirect URI**, which is correct for client credentials. One side
effect: the `/adminconsent` URL does not work for it (AADSTS500113) — an admin must grant
consent from the portal's API permissions page instead.

### Checking it

`/api/data?debug=1` reports missing variables, which step failed (`token`, `site`, `sheet`),
the resolved used-range addresses, and the first rows of each tab. No secret is exposed.

### Sheet layout

Columns are read **by position**, not by header name:

- Tracker: A Date, B Name, D CV date, E Job code + position, F Source, G Mode,
  H Interviewer, I Recruiter, J Location, K Status, L Interview date, M Current salary,
  N Expected salary, O Selected date, P Joining date, R Remark.
  Columns S–V (computed day-counts) are ignored.
- Sheet3: B Designation, C Location, D Status, E Start, F End, G Department, H Nos,
  I On board, J Balance, K Owner, L Remark.

**Insert a column in the middle and every field after it silently shifts** — wrong numbers,
no error. Add new columns at the right-hand end only.

Salary cells accept `40k`, `22.5k`, `40000` and `4.5 lpa` (converted to monthly).

## Notes

- **Dates** arrive from Graph as Excel serial numbers and are converted with pure UTC
  arithmetic. Any local-time conversion shifts them a day and silently corrupts every
  figure in the dashboard — do not "simplify" this.
- **Client secrets expire** (24 months max, often 6). When one does, the dashboard falls
  back to the snapshot and `?debug=1` reports `failedAt: "token"`. Set a reminder.
- Columns are read by position from column A, so inserting a column in the middle of the
  sheet will shift everything. Add new columns at the right-hand end.
- **No auto-refresh.** The workbook is read once when the page opens and then only when
  the user presses **↻ Refresh**. The header shows "Updated HH:MM:SS" so nobody mistakes a
  stale screen for a live one. This was a deliberate cost decision: polling every 5s was
  consuming ~17,000 serverless invocations a day per open tab and nearly exhausted the
  Vercel Hobby Fluid Active CPU allowance.
- Responses carry `s-maxage=60, stale-while-revalidate=300`, so repeat page loads are
  served by Vercel's CDN rather than invoking the function. Refresh sends `fresh=1`, which
  bypasses both the CDN and the 60s in-memory cache.
- Do **not** add a cache-busting query parameter to the normal request — a unique URL per
  request makes every CDN lookup a guaranteed miss, which is what defeated edge caching
  the first time.
- `.gitattributes` pins `*.html` to LF endings, so Windows CRLF does not turn every commit
  into a full-file diff.
- The Period filter selects candidates by **CV month**. Joined / Still with us / Joined &
  left instead count people by the month they actually started — toggle with the
  `Joins by:` button.

## Access

The dashboard and `/api/data` are both public on the Vercel URL. Anyone with the link can
read every candidate name, salary and remark — the same as when the Google Sheet was shared
as "anyone with the link". If that needs restricting, turn on Vercel Authentication
(Project → Settings → Deployment Protection).
