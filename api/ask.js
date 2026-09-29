/* ---------------------------------------------------------------------------
   HR Dashboard — AI copilot endpoint (Vercel serverless, Node runtime)

   The browser answers most questions itself from data it already holds. This
   endpoint is only reached when a question needs reasoning rather than counting.

   It receives:  { question, brief }   brief = AGGREGATES ONLY, never candidate rows.
   It returns :  { answer, model, ms }

   The Groq API key lives in the GROQ_API_KEY environment variable and never
   reaches the browser. Set it in Vercel → Project → Settings → Environment
   Variables, then redeploy.

   Optional env:
     GROQ_MODEL   default llama-3.3-70b-versatile
--------------------------------------------------------------------------- */

const GROQ_URL = 'https://api.groq.com/openai/v1/chat/completions';
const DEFAULT_MODEL = 'llama-3.3-70b-versatile';
const TIMEOUT_MS = 12000;          // fail fast — a slow answer is a failed answer
const MAX_BRIEF_CHARS = 24000;     // hard ceiling so a runaway payload can't stall us

const SYSTEM = `You are the analyst built into an HR recruitment dashboard for an Indian engineering company (Elixir).

You are given a JSON brief of PRE-AGGREGATED recruitment figures. That brief is your only source of truth.

Rules:
- Use only numbers present in the brief. Never invent or estimate a figure that is not there.
- If the brief cannot answer the question, say so plainly in one sentence and name what is missing.
- Lead with the direct answer. Then at most 3 short supporting points.
- Be specific and quantitative. Quote the actual numbers.
- Keep the whole reply under 150 words. No preamble, no restating the question.
- Plain text with simple markdown (**bold**, "- " bullets). No headings, no tables, no code blocks.
- Amounts are Indian rupees; write them as ₹40,000 or ₹1.2L.
- "joined" means the person actually started. "CV cohort" means candidates grouped by the month their CV arrived.
- Where a number looks like a data-entry problem rather than a real result, say so.`;

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(204).end();
  }
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Use POST.' });
  }

  const key = process.env.GROQ_API_KEY;
  if (!key) {
    // Explicit, actionable — this is the single most likely setup mistake.
    return res.status(503).json({
      error: 'no_key',
      answer: 'The AI side is not configured yet. Add a GROQ_API_KEY environment variable in Vercel (Settings → Environment Variables) and redeploy. Everything the dashboard can answer from its own data still works without it.'
    });
  }

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch { body = {}; } }
  if (!body || typeof body !== 'object') body = {};

  const question = String(body.question || '').slice(0, 800).trim();
  if (!question) return res.status(400).json({ error: 'Empty question.' });

  let brief = '';
  try { brief = JSON.stringify(body.brief ?? {}); } catch { brief = '{}'; }
  if (brief.length > MAX_BRIEF_CHARS) brief = brief.slice(0, MAX_BRIEF_CHARS) + '…(truncated)';

  const model = process.env.GROQ_MODEL || DEFAULT_MODEL;
  const started = Date.now();
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);

  try {
    const r = await fetch(GROQ_URL, {
      method: 'POST',
      signal: ctl.signal,
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + key },
      body: JSON.stringify({
        model,
        temperature: 0.2,       // analysis, not creative writing
        max_tokens: 420,        // caps the slowest part of the round trip
        messages: [
          { role: 'system', content: SYSTEM },
          { role: 'user', content: 'FIGURES (JSON):\n' + brief + '\n\nQUESTION: ' + question }
        ]
      })
    });

    const txt = await r.text();
    if (!r.ok) {
      let detail = txt.slice(0, 300);
      try { detail = JSON.parse(txt).error?.message || detail; } catch {}
      return res.status(502).json({
        error: 'upstream',
        answer: 'The AI service returned an error (' + r.status + '): ' + detail
      });
    }

    const data = JSON.parse(txt);
    const answer = data?.choices?.[0]?.message?.content?.trim();
    if (!answer) return res.status(502).json({ error: 'empty', answer: 'The AI service returned an empty reply. Try rephrasing the question.' });

    return res.status(200).json({ answer, model, ms: Date.now() - started });

  } catch (e) {
    const aborted = e && (e.name === 'AbortError' || /abort/i.test(String(e.message || '')));
    return res.status(aborted ? 504 : 500).json({
      error: aborted ? 'timeout' : 'failed',
      answer: aborted
        ? 'The AI service took longer than ' + (TIMEOUT_MS / 1000) + ' seconds and was cut off. Ask something narrower, or try again.'
        : 'Could not reach the AI service: ' + String(e && e.message || e)
    });
  } finally {
    clearTimeout(timer);
  }
};
