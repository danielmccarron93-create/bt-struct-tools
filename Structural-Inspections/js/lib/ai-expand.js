/**
 * AI comment expansion (north-star §8.2).
 *
 * Takes an engineer's short-form dictation ("short cover on P3 bottom") and
 * asks Claude to expand it into a polished, AS-referenced rectification
 * comment in the Bligh Tanner house voice.
 *
 * Refactored in Phase 11 to use the shared lib/anthropic.js client. Behaviour
 * is unchanged for callers — same opts in, same { text, asClause } out.
 */

import { callClaude } from './anthropic.js';

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
 * Expand a short-form note. Returns { text, asClause }.
 */
export async function expandComment({ apiKey, note, context = {} }) {
  if (!note || !note.trim()) {
    throw new Error('Type or dictate a short note first, then tap Expand.');
  }

  // Mark the system prompt as cacheable so repeated expansions during one
  // inspection only pay for the user prompt + output tokens.
  const result = await callClaude({
    apiKey,
    label:  'comment-expand',
    system: [{ text: SYSTEM_PROMPT, cache: true }],
    messages: [
      { role: 'user', content: buildUserPrompt(context, note) }
    ],
    maxTokens: 400,
    responseSchema: { required: ['text'] }
  });

  return {
    text:     String(result.parsed.text || '').trim(),
    asClause: String(result.parsed.asClause || '').trim()
  };
}
