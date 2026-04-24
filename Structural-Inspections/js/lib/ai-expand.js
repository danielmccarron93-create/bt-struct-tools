/**
 * AI comment expansion (Phase V2) — the north-star §8.2 feature.
 *
 * Takes an engineer's short-form dictation ("short cover on P3 bottom") and
 * asks Claude Sonnet 4.6 to expand it into a polished, AS-referenced
 * rectification comment in the Bligh Tanner house voice.
 *
 * Design choices:
 *   • User-provided API key, stored per-device in IndexedDB under
 *     settings['user.profile'].anthropicKey. Each engineer manages their own
 *     quota and spend — no shared secret baked into the app.
 *   • Direct browser call to https://api.anthropic.com/v1/messages with
 *     `anthropic-dangerous-direct-browser-access: true`. No backend required.
 *   • Offline-safe: throws a friendly error if navigator.onLine is false, so
 *     the caller can fall back to the raw dictated text.
 *   • Never mutates the engineer's comment — the caller shows a preview and
 *     lets the engineer accept, edit, or reject.
 */

const API_URL = 'https://api.anthropic.com/v1/messages';
const MODEL   = 'claude-sonnet-4-6';
const MAX_TOKENS = 400;

const SYSTEM_PROMPT = `You are a senior structural engineer at Bligh Tanner (Brisbane, Australia) rewriting short-form inspection notes into polished rectification comments for a site inspection report.

Bligh Tanner house style:
- Measured, factual, defensible. Never speculative or inflammatory.
- Always references the relevant Australian Standard clause when one applies — AS 3600 (concrete), AS 4100 (steel), AS 1720.1 (timber), AS 2870 (residential slabs), AS 3700 (masonry), AS 1684 (light framing), AS 1379 (concrete supply).
- Uses Australian English spelling.
- Active voice for defects ("Install additional ligatures…"), passive for observations ("Ligature spacing confirmed…").
- One to two short sentences — no preamble, no repetition of the location (the report's Location column handles that).
- Uses square brackets for values the engineer still needs to fill in, e.g. "[X] mm", "[N] off".
- If the input clearly describes a hold-point situation (works must not proceed), begin with "HOLD POINT — ".

Output a single JSON object: {"text": "...", "asClause": "AS 3600:2018 Cl. 4.10.3"}. No markdown, no commentary. If no AS clause applies, set asClause to "".`;

function buildUserPrompt(ctx, note) {
  const lines = [];
  lines.push(`Inspection type: ${ctx.inspectionTypeName || '(unspecified)'}`);
  if (ctx.drawingSheet) lines.push(`Drawing: ${ctx.drawingSheet}${ctx.drawingRevision ? ` Rev ${ctx.drawingRevision}` : ''}`);
  if (ctx.gridRef)      lines.push(`Location: ${ctx.gridRef}`);
  if (ctx.severity)     lines.push(`Severity: ${ctx.severity}`);
  lines.push('');
  lines.push('Engineer\u2019s short note:');
  lines.push(note.trim());
  lines.push('');
  lines.push('Rewrite this into a single polished rectification comment. Respond with JSON only.');
  return lines.join('\n');
}

/**
 * Expand a short-form note using Claude. Throws on error (caller shows toast).
 *
 * opts.apiKey        — Anthropic API key (required)
 * opts.note          — short-form engineer note (required)
 * opts.context       — { inspectionTypeName, drawingSheet, drawingRevision, gridRef, severity }
 *
 * Returns: { text: string, asClause: string }
 */
export async function expandComment({ apiKey, note, context = {} }) {
  if (!apiKey || !apiKey.trim()) {
    throw new Error('Anthropic API key not set — add one in Settings to use AI expansion.');
  }
  if (!note || !note.trim()) {
    throw new Error('Type or dictate a short note first, then tap Expand.');
  }
  if (navigator.onLine === false) {
    throw new Error('Offline — AI expansion needs a connection. Your short note is kept as-is.');
  }

  const body = {
    model: MODEL,
    max_tokens: MAX_TOKENS,
    system: SYSTEM_PROMPT,
    messages: [
      { role: 'user', content: buildUserPrompt(context, note) }
    ]
  };

  let resp;
  try {
    resp = await fetch(API_URL, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': apiKey.trim(),
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true'
      },
      body: JSON.stringify(body)
    });
  } catch (err) {
    throw new Error(`Network error calling Claude: ${err.message || err}`);
  }

  if (!resp.ok) {
    let detail = '';
    try { detail = (await resp.json())?.error?.message || ''; } catch {}
    throw new Error(`Claude API ${resp.status}${detail ? ' — ' + detail : ''}`);
  }

  const payload = await resp.json();
  const text = (payload.content || []).map((c) => c.text || '').join('\n').trim();
  if (!text) throw new Error('Empty response from Claude — try again.');

  // Parse the JSON reply — tolerate occasional leading/trailing whitespace or
  // a stray ```json fence even though the system prompt forbids it.
  const cleaned = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (err) {
    throw new Error('Claude returned a non-JSON response — try again, or refine your note.');
  }

  return {
    text:     String(parsed.text || '').trim(),
    asClause: String(parsed.asClause || '').trim()
  };
}
