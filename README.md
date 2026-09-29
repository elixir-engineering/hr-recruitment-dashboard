# HR Recruitment Dashboard

Single-page dashboard over the recruitment Google Sheet, deployed on Vercel.

- `index.html` — the whole dashboard. Reads the Google Sheet live, falls back to an embedded snapshot.
- `api/ask.js` — serverless endpoint for the AI copilot. Holds the Groq key server-side.

## Setting up the AI copilot

The copilot works in two tiers:

1. **From the dashboard** — counting and ranking questions ("how many joined in August", "which recruiter converts best") are answered in the browser from data already loaded. No network call, no API key, typically under 20 ms.
2. **AI analysis** — questions asking *why*, *what should we fix*, *compare*, or anything needing judgement go to Groq through `api/ask.js`.

Tier 1 works with no setup at all. Tier 2 needs a key.

### Add the Groq key

1. Get a free key at <https://console.groq.com> → API Keys.
2. In Vercel: **Project → Settings → Environment Variables**
   - Name: `GROQ_API_KEY`
   - Value: your key
   - Apply to Production, Preview and Development.
3. Redeploy (Deployments → ⋯ → Redeploy).

Optional: set `GROQ_MODEL` to override the default `llama-3.3-70b-versatile`. Use
`llama-3.1-8b-instant` if you want faster, shallower answers.

### What gets sent to Groq

Only the question and an **aggregate brief** — counts, rates, medians, and per-recruiter /
per-source / per-department / per-month rollups. Roughly 7 KB.

**No candidate names, no individual salaries, no remarks ever leave the browser.**

This is deliberate. If you ever want the AI to read free-text remarks, that is a conscious
decision to make, not something to add by accident.

### Testing locally

Opening `index.html` directly from disk gives you tier 1 only — there is no `/api` route on
`file://`. The copilot says so plainly and still answers what it can. For the full thing run
`vercel dev`, or just test on the deployed site.

## Notes

- `.gitattributes` pins `*.html` to LF endings. Without it, Windows CRLF turns every commit
  into a full-file diff.
- The Period filter selects candidates by **CV month**. Joined / Still with us / Joined & left
  instead count people by the month they actually started — toggle with the `Joins by:` button.
