# HR Recruitment Dashboard

Single-page dashboard over `HR_data.xlsx` on SharePoint, deployed on Vercel.

- `index.html` — the whole dashboard. Reads live data from `/api/data`, falls back to an embedded snapshot.
- `api/data.js` — downloads the workbook server-side and returns its rows as JSON.
- `package.json` — one dependency, SheetJS, for reading the .xlsx.

## There is nothing to configure

The file's SharePoint link is in `api/data.js`. No keys, no Azure app, no environment variables.

The only reason this function exists is that the browser cannot fetch the workbook itself —
SharePoint sends no `Access-Control-Allow-Origin` header, so JavaScript on `vercel.app` is
blocked. CORS is a browser rule, not a server one, so the function downloads the file and
hands the browser JSON from its own origin.

## The one requirement

**The sharing link must be set to "Anyone with the link".**

In SharePoint: open `HR_data.xlsx` → **Share** → click the settings/gear on the link →
**Anyone with the link** → *Can view* → copy that link.

If it is restricted instead, SharePoint serves a sign-in page rather than the file. The
dashboard detects this exactly and says so, rather than showing wrong numbers.

If your tenant blocks anonymous links entirely, tell me — the fallback is a Microsoft Graph
app registration, which works but needs an admin.

## If the link ever changes

Either edit `SHARE_URL` at the top of `api/data.js`, or set `HR_XLSX_URL` in Vercel's
environment variables to override it without touching the code.

Optional overrides: `HR_TAB_TRACKER` (default `Recruitment Tracker`),
`HR_TAB_REQS` (default `Sheet3`), `CACHE_SECONDS` (default `60`).

## Checking it

Open `https://your-site.vercel.app/api/data?debug=1`. It shows the tabs found in the
workbook, how many rows each returned, and the first two tracker rows — enough to spot a
renamed tab or a shifted column immediately.

## Notes

- **Dates** are converted from Excel serial numbers using UTC arithmetic, not SheetJS's
  `cellDates`, which shifts them a day through the server's timezone. Do not "simplify"
  this — it silently corrupts every date in the dashboard.
- Tab lookup is case-insensitive and tolerates stray spaces.
- Columns are read by position from column A, so inserting a column in the middle of the
  sheet will shift everything. Add new columns at the right-hand end.
- 60-second server-side cache; the page itself re-reads every 5 minutes.
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
