/**
 * Generic Claude / Anthropic API client.
 *
 * One source of truth for every LLM call in the app — comment expansion,
 * title-block extraction, drawing classification, inspection planning,
 * checklist generation, etc.
 *
 * Design choices:
 *
 *   • User-provided API key, stored per-device in IndexedDB under
 *     settings['user.profile'].anthropicKey. Never leaves the browser
 *     (direct call to api.anthropic.com with
 *     `anthropic-dangerous-direct-browser-access: true`).
 *
 *   • Prompt-cache aware. Pass `system` as an array of blocks with
 *     `{ text, cache: true }` to mark cache_control: ephemeral. The
 *     cached system prompt amortises across batch calls.
 *
 *   • Structured-output friendly. Pass `responseSchema` to declare the
 *     expected JSON shape; the client strips ```json fences, parses,
 *     and validates required keys before returning `parsed`.
 *
 *   • Retry with backoff on 429 / 5xx. Friendly errors otherwise.
 *
 *   • Batch helper with concurrency control and progress callbacks.
 *
 *   • Usage tracker — every call logs input/output/cached token counts
 *     to settings['ai.usage'] so the Settings screen can show an
 *     approximate running spend. Pricing table is illustrative; update
 *     when Anthropic publishes new rates.
 */

import { db, getUserProfile } from '../db.js';

const API_URL  = 'https://api.anthropic.com/v1/messages';
const VERSION  = '2023-06-01';

/* Default model — used unless the caller overrides. Sonnet 4.6 is the
 * current sweet spot for cost/quality on the heavy passes. Bump as new
 * Claude families ship. */
export const DEFAULT_MODEL = 'claude-sonnet-4-6';

/* Approximate per-million-token pricing (USD). Used only for the on-device
 * cost tally — billing of record is on console.anthropic.com. Update when
 * Anthropic publishes new rates. */
const PRICING = {
  'claude-opus-4-7':       { in: 15.00, out: 75.00, cached: 1.50 },
  'claude-sonnet-4-6':     { in:  3.00, out: 15.00, cached: 0.30 },
  'claude-haiku-4-5':      { in:  1.00, out:  5.00, cached: 0.10 }
};

const USAGE_KEY = 'ai.usage';

/* --------------------------------------------------------------------------
   Public API
   -------------------------------------------------------------------------- */

/**
 * Resolve the API key from the saved user profile.
 * Throws a friendly error if missing — callers can catch and toast.
 */
export async function getApiKey() {
  const profile = await getUserProfile();
  const key = (profile?.anthropicKey || '').trim();
  if (!key) {
    throw new Error('Anthropic API key not set — add one in Settings to use AI features.');
  }
  return key;
}

/**
 * Single LLM call.
 *
 * @param {Object}  opts
 * @param {string}  opts.apiKey               — Anthropic API key (use getApiKey()).
 * @param {string|Array} opts.system          — System prompt. String for plain,
 *                                              or array of { text, cache?: bool }
 *                                              blocks for cache_control: ephemeral.
 * @param {Array}   opts.messages             — Anthropic messages array. Content can
 *                                              be plain string or full content blocks
 *                                              (text + image base64 etc.).
 * @param {Object}  [opts.responseSchema]     — { required: ['key1', 'key2'] }. If
 *                                              present, the response text is parsed
 *                                              as JSON and validated; the result is
 *                                              returned as `parsed`.
 * @param {number}  [opts.maxTokens=2000]     — Max tokens for the response.
 * @param {string}  [opts.model]              — Model id; defaults to DEFAULT_MODEL.
 * @param {number}  [opts.temperature=0]      — Determinism control.
 * @param {string}  [opts.label]              — Free-form label tagged on usage records.
 * @param {AbortSignal} [opts.signal]
 *
 * @returns {Promise<{
 *   content: string,        // raw text reply
 *   parsed?: any,           // parsed JSON if responseSchema was provided
 *   usage: {
 *     model: string,
 *     inputTokens: number,
 *     outputTokens: number,
 *     cachedInputTokens: number,
 *     cacheCreationTokens: number,
 *     approxCostUsd: number,
 *     latencyMs: number
 *   }
 * }>}
 */
export async function callClaude(opts) {
  const {
    apiKey,
    system,
    messages,
    responseSchema,
    maxTokens = 2000,
    model = DEFAULT_MODEL,
    temperature = 0,
    label = '',
    signal
  } = opts;

  if (!apiKey || !apiKey.trim()) {
    throw new Error('Anthropic API key not set — add one in Settings.');
  }
  if (!Array.isArray(messages) || messages.length === 0) {
    throw new Error('callClaude: messages array is required and must be non-empty.');
  }
  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
    throw new Error('Offline — AI features need a connection.');
  }

  const body = {
    model,
    max_tokens: maxTokens,
    temperature,
    messages
  };
  if (system !== undefined) {
    body.system = normaliseSystem(system);
  }

  const startedAt = Date.now();
  let payload;
  try {
    payload = await fetchWithRetry(API_URL, {
      method:  'POST',
      headers: {
        'content-type':                                'application/json',
        'x-api-key':                                   apiKey.trim(),
        'anthropic-version':                           VERSION,
        'anthropic-dangerous-direct-browser-access':   'true'
      },
      body:    JSON.stringify(body),
      signal
    });
  } catch (err) {
    throw new Error(`Claude request failed: ${err.message || err}`);
  }
  const latencyMs = Date.now() - startedAt;

  // Extract text content
  const contentText = (payload.content || [])
    .filter((c) => c?.type === 'text')
    .map((c) => c.text || '')
    .join('\n')
    .trim();
  if (!contentText) {
    throw new Error('Empty response from Claude — try again.');
  }

  // Usage accounting — Anthropic returns input_tokens, output_tokens,
  // cache_creation_input_tokens, cache_read_input_tokens.
  const u = payload.usage || {};
  const inputTokens         = Number(u.input_tokens || 0);
  const outputTokens        = Number(u.output_tokens || 0);
  const cachedInputTokens   = Number(u.cache_read_input_tokens || 0);
  const cacheCreationTokens = Number(u.cache_creation_input_tokens || 0);
  const approxCostUsd       = approxCost(model, inputTokens, outputTokens, cachedInputTokens, cacheCreationTokens);

  const usage = {
    model,
    label,
    inputTokens,
    outputTokens,
    cachedInputTokens,
    cacheCreationTokens,
    approxCostUsd,
    latencyMs
  };
  // Fire-and-forget — don't block the caller on the tally write.
  recordUsage(usage).catch(() => {});

  // Parse JSON if a schema was declared.
  let parsed;
  if (responseSchema) {
    parsed = parseJsonReply(contentText);
    validateAgainstSchema(parsed, responseSchema);
  }

  return { content: contentText, parsed, usage };
}

/**
 * Fan out N independent LLM calls with a concurrency limit. Each item is
 * passed to `fn` which must return a Promise. Returns an array of results in
 * the SAME order as the input items. Errors propagate (not eaten).
 *
 * @param {Array}    items
 * @param {Function} fn               — async (item, index) => result
 * @param {Object}   [opts]
 * @param {number}   [opts.concurrency=3]
 * @param {Function} [opts.onProgress] — (doneCount, totalCount, lastResult) => void
 *
 * @returns {Promise<Array>}
 */
export async function callClaudeBatch(items, fn, { concurrency = 3, onProgress } = {}) {
  const results = new Array(items.length);
  let nextIndex = 0;
  let doneCount = 0;

  const total = items.length;
  if (total === 0) return results;

  const workers = Array.from({ length: Math.min(concurrency, total) }, async () => {
    while (true) {
      const i = nextIndex++;
      if (i >= total) return;
      try {
        results[i] = await fn(items[i], i);
      } catch (err) {
        results[i] = { error: err };
      } finally {
        doneCount += 1;
        if (typeof onProgress === 'function') {
          try { onProgress(doneCount, total, results[i]); } catch {}
        }
      }
    }
  });

  await Promise.all(workers);

  // If any worker errored, surface the first one — keeps debugging easy.
  const firstErr = results.find((r) => r && r.error);
  if (firstErr) throw firstErr.error;
  return results;
}

/**
 * Read the cumulative running usage tally (across all calls since install).
 * Surfaced on the Settings screen.
 */
export async function readUsageTally() {
  const row = await db.settings.get(USAGE_KEY);
  return row?.value || {
    totalCalls: 0,
    totalInputTokens: 0,
    totalOutputTokens: 0,
    totalCachedInputTokens: 0,
    totalCacheCreationTokens: 0,
    totalApproxCostUsd: 0,
    byLabel: {},
    byModel: {},
    lastCallAt: null
  };
}

/**
 * Reset the usage tally — for users who want to zero things between projects
 * or after a billing period.
 */
export async function resetUsageTally() {
  await db.settings.delete(USAGE_KEY);
}

/* --------------------------------------------------------------------------
   Internals
   -------------------------------------------------------------------------- */

/**
 * Convert a system spec (string or array of blocks) into the Anthropic
 * messages-API format. When the caller passes an array of blocks, blocks
 * marked `cache: true` get cache_control: ephemeral.
 */
function normaliseSystem(system) {
  if (typeof system === 'string') return system;
  if (!Array.isArray(system)) {
    throw new Error('system must be a string or an array of blocks');
  }
  return system.map((block) => {
    if (typeof block === 'string') return { type: 'text', text: block };
    const out = { type: 'text', text: String(block.text || '') };
    if (block.cache) out.cache_control = { type: 'ephemeral' };
    return out;
  });
}

/**
 * Strip ```json fences, parse JSON. Throws a friendly error on malformed
 * output so callers can catch and toast.
 */
function parseJsonReply(text) {
  // Tolerate markdown fences and stray leading/trailing prose.
  let cleaned = String(text || '').trim();
  cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');
  // If there's prose before/after the JSON, try to grab the outermost {...} or [...].
  if (!/^[\[{]/.test(cleaned)) {
    const firstBracket = cleaned.search(/[\[{]/);
    const lastBracket  = cleaned.search(/[\]}](?=\s*$)/);
    // search returns -1 on miss. Bail if either side missing.
    if (firstBracket >= 0 && lastBracket >= 0) {
      cleaned = cleaned.slice(firstBracket);
    }
  }
  try {
    return JSON.parse(cleaned);
  } catch (err) {
    throw new Error(`Claude returned non-JSON: ${truncate(cleaned, 120)}`);
  }
}

/**
 * Trivial schema check — only validates `required` keys for now. Good enough
 * for the structured calls we make. Avoids a dependency on Ajv.
 */
function validateAgainstSchema(value, schema) {
  if (!schema) return;
  const required = schema.required || [];
  const target = schema.itemRequired ? value?.[0] : value;
  if (!target || typeof target !== 'object') {
    throw new Error('LLM response is not an object as expected.');
  }
  for (const key of required) {
    if (!(key in target)) {
      throw new Error(`LLM response is missing required key "${key}".`);
    }
  }
}

/**
 * fetch with exponential backoff on 429 / 5xx. Up to 3 retries.
 * Honours Retry-After header when present.
 */
async function fetchWithRetry(url, init, { maxRetries = 3 } = {}) {
  let lastErr;
  for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
    let resp;
    try {
      resp = await fetch(url, init);
    } catch (err) {
      lastErr = err;
      if (attempt < maxRetries) {
        await sleep(backoffMs(attempt));
        continue;
      }
      throw err;
    }

    if (resp.ok) {
      return resp.json();
    }

    // Retry-eligible statuses
    if ((resp.status === 429 || resp.status >= 500) && attempt < maxRetries) {
      const retryAfter = Number(resp.headers.get('retry-after')) || 0;
      const wait = retryAfter > 0 ? retryAfter * 1000 : backoffMs(attempt);
      await sleep(wait);
      continue;
    }

    let detail = '';
    try { detail = (await resp.json())?.error?.message || ''; } catch {}
    throw new Error(`Claude API ${resp.status}${detail ? ' — ' + detail : ''}`);
  }
  throw lastErr || new Error('Claude API: max retries exceeded');
}

function backoffMs(attempt) {
  // Exponential backoff with jitter — 1s, 2s, 4s base, ±20% jitter.
  const base = 1000 * 2 ** attempt;
  const jitter = base * 0.2 * (Math.random() * 2 - 1);
  return Math.max(250, Math.round(base + jitter));
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function truncate(s, n) {
  return s.length > n ? s.slice(0, n) + '…' : s;
}

function approxCost(model, inputTokens, outputTokens, cachedInputTokens, cacheCreationTokens) {
  const p = PRICING[model];
  if (!p) return 0;
  const freshInput = Math.max(0, inputTokens - cachedInputTokens);
  return (
    (freshInput / 1e6) * p.in +
    (cachedInputTokens / 1e6) * p.cached +
    (cacheCreationTokens / 1e6) * p.in * 1.25 +  // cache writes are 1.25x input
    (outputTokens / 1e6) * p.out
  );
}

async function recordUsage(call) {
  const tally = await readUsageTally();

  tally.totalCalls += 1;
  tally.totalInputTokens         += call.inputTokens;
  tally.totalOutputTokens        += call.outputTokens;
  tally.totalCachedInputTokens   += call.cachedInputTokens;
  tally.totalCacheCreationTokens += call.cacheCreationTokens;
  tally.totalApproxCostUsd       += call.approxCostUsd;
  tally.lastCallAt                = new Date().toISOString();

  if (call.label) {
    const lab = tally.byLabel[call.label] || { calls: 0, costUsd: 0 };
    lab.calls   += 1;
    lab.costUsd += call.approxCostUsd;
    tally.byLabel[call.label] = lab;
  }
  const m = tally.byModel[call.model] || { calls: 0, costUsd: 0 };
  m.calls   += 1;
  m.costUsd += call.approxCostUsd;
  tally.byModel[call.model] = m;

  await db.settings.put({ key: USAGE_KEY, value: tally });
}
