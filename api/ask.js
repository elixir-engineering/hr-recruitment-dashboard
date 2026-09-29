/* ---------------------------------------------------------------------------
   HR Dashboard — AI copilot endpoint (Vercel serverless, Node runtime)

   The browser answers most questions itself from data it already holds. This
   endpoint is only reached when a question needs reasoning rather than counting.

   POST { question, brief }   brief = AGGREGATES ONLY, never candidate rows.
    ->  { answer, model, ms }

   GET  -> a small diagnostic: is the key set, which models can it reach, which
           one would be used. Open /api/ask in a browser to check the setup.

   The Groq API key lives in GROQ_API_KEY and never reaches the browser.

   Model selection: Groq moves models between free, preview and enterprise
   tiers, so a hardcoded name goes stale (llama-3.3-70b-versatile became
   enterprise-only and started returning 404 to free keys). Instead we ask the
   API which models THIS key can see, then take the best one we recognise.
   Set GROQ_MODEL to pin a specific model and skip all of that.
--------------------------------------------------------------------------- */

const BASE = 'https://api.groq.com/openai/v1';
const TIMEOUT_MS = 12000;
const MAX_BRIEF_CHARS = 24000;

/* best first — quality, then speed. Anything not on this list is still used as
   a last resort if the account only has something exotic. */
const PREFERRED = [
  'openai/gpt-oss-120b',        // free tier, ~500 tok/s, best reasoning of the open set
  'openai/gpt-oss-20b',         // ~1000 tok/s, noticeably faster, a little shallower
  'qwen/qwen3.8-27b',
  'llama-3.3-70b-versatile',    // enterprise on most accounts, kept in case yours has it
  'llama-3.1-8b-instant'
];
/* never route a chat question to these */
const NOT_CHAT = /whisper|tts|orpheus|prompt-guard|safeguard|embed|moderation/i;

let MODEL_CACHE = { ids: null, at: 0 };
const CACHE_MS = 10 * 60 * 1000;

async function listModels(key) {
  if (MODEL_CACHE.ids && Date.now() - MODEL_CACHE.at < CACHE_MS) return MODEL_CACHE.ids;
  try {
    const r = await fetch(BASE + '/models', { headers: { Authorization: 'Bearer ' + key } });
    if (!r.ok) return null;
    const j = await r.json();
    const ids = (j.data || []).filter(m => m && m.active !== false).map(m => m.id);
    if (!ids.length) return null;
    MODEL_CACHE = { ids, at: Date.now() };
    return ids;
  } catch { return null; }
}

/* ordered list of models worth trying for this key */
async function candidates(key) {
  if (process.env.GROQ_MODEL) return [process.env.GROQ_MODEL];
  const ids = await listModels(key);
  if (!ids) return PREFERRED.slice();                       // /models unreachable — try blind
  const known = PREFERRED.filter(m => ids.includes(m));
  const rest = ids.filter(id => !NOT_CHAT.test(id) && !known.includes(id));
  return known.concat(rest).slice(0, 4);                    // cap the retry budget
}

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

async function chat(key, model, question, brief, signal) {
  const r = await fetch(BASE + '/chat/completions', {
    method: 'POST',
    signal,
    headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + key },
    body: JSON.stringify({
      model,
      temperature: 0.2,
      max_tokens: 420,
      messages: [
        { role: 'system', content: SYSTEM },
        { role: 'user', content: 'FIGURES (JSON):\n' + brief + '\n\nQUESTION: ' + question }
      ]
    })
  });
  const txt = await r.text();
  return { ok: r.ok, status: r.status, txt };
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  const key = process.env.GROQ_API_KEY;

  /* ---- diagnostic: open /api/ask in a browser ---- */
  if (req.method === 'GET') {
    if (!key) return res.status(200).json({ keySet: false, hint: 'Add GROQ_API_KEY in Vercel → Settings → Environment Variables, then redeploy.' });
    const ids = await listModels(key);
    const list = await candidates(key);
    return res.status(200).json({
      keySet: true,
      pinnedByEnv: process.env.GROQ_MODEL || null,
      willUse: list[0] || null,
      fallbacks: list.slice(1),
      modelsVisibleToThisKey: ids ? ids.filter(i => !NOT_CHAT.test(i)) : 'could not list (check the key is valid)'
    });
  }

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(204).end();
  }
  if (req.method !== 'POST') return res.status(405).json({ error: 'Use POST.' });

  if (!key) {
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

  const started = Date.now();
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);

  try {
    const list = await candidates(key);
    if (!list.length) {
      return res.status(502).json({ error: 'no_model', answer: 'This Groq key cannot reach any chat model. Open /api/ask in a browser to see what it can see.' });
    }

    let lastDetail = '';
    for (const model of list) {
      const { ok, status, txt } = await chat(key, model, question, brief, ctl.signal);

      if (ok) {
        const data = JSON.parse(txt);
        const answer = data?.choices?.[0]?.message?.content?.trim();
        if (answer) return res.status(200).json({ answer, model, ms: Date.now() - started });
        lastDetail = 'empty reply';
        continue;
      }

      try { lastDetail = JSON.parse(txt).error?.message || txt.slice(0, 200); }
      catch { lastDetail = txt.slice(0, 200); }

      /* no access / gone / bad model -> drop the cache and try the next one */
      if (status === 404 || status === 403 || /model/i.test(lastDetail)) {
        MODEL_CACHE = { ids: null, at: 0 };
        continue;
      }
      if (status === 429) {
        return res.status(429).json({ error: 'rate_limit', answer: 'Groq rate limit hit on the free tier. Wait a moment and ask again.' });
      }
      return res.status(502).json({ error: 'upstream', answer: 'The AI service returned an error (' + status + '): ' + lastDetail });
    }

    return res.status(502).json({
      error: 'all_models_failed',
      answer: 'None of the models this key can reach accepted the request (' + list.join(', ') + '). Last error: ' + lastDetail
        + ' — open /api/ask in a browser to see what the key can access.'
    });

  } catch (e) {
    const aborted = e && (e.name === 'AbortError' || /abort/i.test(String(e.message || '')));
    return res.status(aborted ? 504 : 500).json({
      error: aborted ? 'timeout' : 'failed',
      answer: aborted
        ? 'The AI service took longer than ' + (TIMEOUT_MS / 1000) + ' seconds and was cut off. Ask something narrower, or try again.'
        : 'Could not reach the AI service: ' + String((e && e.message) || e)
    });
  } finally {
    clearTimeout(timer);
  }
};
