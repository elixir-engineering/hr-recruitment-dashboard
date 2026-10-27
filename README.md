# HR Recruitment Dashboard

Single-page dashboard over `HR_data.xlsx` on SharePoint, deployed on Vercel.

- `index.html` — the whole dashboard. Reads live data from `/api/data`, falls back to an embedded snapshot.
- `api/data.js` — signs in to Microsoft and reads the workbook through the Graph API.
- `package.json` — no dependencies; the route uses plain `fetch`.

## Setup

The workbook stays **private**. The route signs in as an Azure AD application
(client credentials) and reads it through the Microsoft Graph API. No sharing
link, no anonymous access, and no credential ever reaches the browser.

### 1. Register an app (needs a Microsoft 365 admin, once)

<https://portal.azure.com> → **Microsoft Entra ID** → **App registrations** → **New registration**

- Name: `HR Dashboard Reader`, accounts: *this organizational directory only*, no redirect URI.
- **Overview** → copy **Application (client) ID** and **Directory (tenant) ID**.
- **Certificates & secrets** → **New client secret** → copy the **Value** immediately
  (shown once; the "Secret ID" is *not* it).
- **API permissions** → **Microsoft Graph** → **Application permissions** →
  add **`Sites.Read.All`** (or `Files.Read.All` if using `DRIVE_USER`) →
  **Grant admin consent**.

### 2. Environment variables in Vercel

**Settings → Environment Variables** (Production, Preview and Development):

| Name | Example |
|---|---|
| `TENANT_ID` | Directory (tenant) ID |
| `CLIENT_ID` | Application (client) ID |
| `CLIENT_SECRET` | the secret **Value** |
| `SITE_HOSTNAME` | `elixirengineering.sharepoint.com` |
| `SITE_PATH` | `/sites/HR` |
| `FILE_PATH` | `Shared Documents/HR_data.xlsx` |
| `SHEET_NAME` | `Recruitment Tracker` |

Also available: `SHEET_NAME_REQS` (default `Sheet3`), `CACHE_SECONDS` (default `60`).

**If the file is on someone's OneDrive** rather than a SharePoint site, leave
`SITE_HOSTNAME`/`SITE_PATH` unset and instead set `DRIVE_USER` to their UPN
(e.g. `aliasgark@elixirengg.com`) with `Files.Read.All` consented. `FILE_PATH`
is then the path inside their OneDrive, e.g. `Documents/HR_data.xlsx`.

Then **redeploy** — environment variables only apply to new builds.

### 3. Check it

`https://your-site.vercel.app/api/data?debug=1` reports which variables are
missing, which step failed (`token`, `site`, `sheet`), the resolved used-range
addresses and the first two rows. No secret is exposed.

## Notes

- **Dates** arrive from Graph as Excel serial numbers and are converted with pure UTC
  arithmetic. Any local-time conversion shifts them a day and silently corrupts every
  figure in the dashboard — do not "simplify" this.
- **Client secrets expire** (24 months max, often 6). When one does, the dashboard falls
  back to the snapshot and `?debug=1` reports `failedAt: "token"`. Set a reminder.
- Columns are read by position from column A, so inserting a column in the middle of the
  sheet will shift everything. Add new columns at the right-hand end.
- 60-second server-side cache. The page polls every 5 seconds but sends the content
  hash it already holds, so an unchanged poll returns ~80 bytes instead of the whole sheet.
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
