# HR Recruitment Dashboard

Single-page dashboard over `HR_data.xlsx` on SharePoint, deployed on Vercel.

- `index.html` — the whole dashboard. Reads live data from `/api/data`, falls back to an embedded snapshot.
- `api/data.js` — reads the workbook server-side through Microsoft Graph.

## Why there is a backend now

The old version read a Google Sheet directly from the browser using Google's `gviz`
JSONP endpoint. SharePoint has no equivalent, and there are two separate blockers:

- The file needs a Microsoft sign-in — there is no anonymous data URL.
- SharePoint sends no `Access-Control-Allow-Origin` header, so even a public file
  could not be read by JavaScript running on `vercel.app`.

`api/data.js` does the reading server-side and hands the browser plain JSON from its
own origin, which solves both.

## Setup — one-time

### 1. Register an app in Azure AD

Someone with Microsoft 365 admin rights needs to do this once.

1. <https://portal.azure.com> → **Microsoft Entra ID** → **App registrations** → **New registration**
   - Name: `HR Dashboard Reader`
   - Accounts: *this organizational directory only*
   - Redirect URI: leave blank
2. From the **Overview** page copy **Application (client) ID** and **Directory (tenant) ID**.
3. **Certificates & secrets** → **New client secret** → copy the **Value** immediately
   (it is only shown once — the "Secret ID" is *not* the value).
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
   → `Files.Read.All` → **Add**, then **Grant admin consent**.

`Files.Read.All` is read-only but tenant-wide. If your admin would rather scope it to a
single file, ask for `Sites.Selected` instead and tell me — it needs a different lookup
in `api/data.js`.

### 2. Add the environment variables in Vercel

**Project → Settings → Environment Variables** (apply to Production, Preview and Development):

| Name | Value |
|---|---|
| `MS_TENANT_ID` | Directory (tenant) ID |
| `MS_CLIENT_ID` | Application (client) ID |
| `MS_CLIENT_SECRET` | the secret **Value** from step 3 |
| `MS_SHARE_URL` | the SharePoint sharing link to `HR_data.xlsx` |

Optional: `MS_TAB_TRACKER` (default `Recruitment Tracker`), `MS_TAB_REQS` (default `Sheet3`),
`CACHE_SECONDS` (default `60`).

Then **redeploy**.

### 3. Check it

Open `https://your-site.vercel.app/api/data?debug=1`. It reports which variables are
missing, whether the token was acquired, and the resolved file name. No secret is
exposed. If it looks right, load the dashboard — the status chip should read
**Live · <time>** and the header **Live from SharePoint**.

## Notes

- Client secrets **expire** (24 months max, often 6). When it does, the dashboard falls
  back to the snapshot and `?debug=1` will say the token failed. Put a calendar reminder
  a month before.
- The dashboard holds a 60-second server-side cache so a page refresh does not hit
  SharePoint every time. The page itself re-reads every 5 minutes.
- Tab names must match the workbook exactly. They are set in two places: `TAB_TRACKER` /
  `TAB_REQS` in `index.html` and the `MS_TAB_*` variables.
- `.gitattributes` pins `*.html` to LF endings. Without it, Windows CRLF turns every
  commit into a full-file diff.
- The Period filter selects candidates by **CV month**. Joined / Still with us / Joined &
  left instead count people by the month they actually started — toggle with the
  `Joins by:` button.

## Access

The dashboard and `/api/data` are both public on the Vercel URL. Anyone with the link can
read every candidate name, salary and remark — the same as when the Google Sheet was shared
as "anyone with the link". Moving the file into SharePoint does **not** by itself make the
dashboard private. If it needs to be restricted, enable Vercel Authentication (Project →
Settings → Deployment Protection) or put the site behind Entra sign-in.
