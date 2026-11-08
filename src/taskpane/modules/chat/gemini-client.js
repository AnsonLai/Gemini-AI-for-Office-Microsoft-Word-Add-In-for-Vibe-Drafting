// Shared HTTP boundary. Retries generateContent requests only; never executes tools.
export class GeminiRequestError extends Error {
  constructor(code, message, status) {
    super(message);
    this.name = 'GeminiRequestError';
    this.code = code;
    if (status !== undefined) this.status = status;
  }
}

const cancelled = () => new GeminiRequestError('REQUEST_CANCELLED', 'Request cancelled by user');
const deadlineReached = () => new GeminiRequestError('REQUEST_DEADLINE', 'The request timed out before Gemini responded. Please try again.');
// Below this much remaining time, a new attempt or retry is not started (KB-13).
const MIN_ATTEMPT_MS = 5000;
const transientStatuses = new Set([408, 429, 500, 502, 503, 504]);

function pause(ms, signal) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) { reject(cancelled()); return; }
    const onAbort = () => { clearTimeout(timer); signal.removeEventListener('abort', onAbort); reject(cancelled()); };
    const timer = setTimeout(() => { signal?.removeEventListener('abort', onAbort); resolve(); }, ms);
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

// The key travels in the x-goog-api-key header, never in the URL, so it stays
// out of URL logs, proxies and error reports.
export function geminiEndpoint(model) {
  if (!model) throw new GeminiRequestError('REQUEST_CONFIGURATION', 'Set a Gemini model and API key in Settings.');
  return `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`;
}

export async function requestGemini({ model, apiKey, payload, ...options }) {
  return requestGeminiJson(geminiEndpoint(model), payload, { ...options, apiKey });
}

export async function requestGeminiJson(url, payload, {
  apiKey, signal, timeoutMs = 90000, maxAttempts = 3, backoffMs = 1000, deadline = Infinity,
  fetchImpl = globalThis.fetch, sleep = pause, random = Math.random, now = Date.now,
} = {}) {
  if (typeof apiKey !== 'string' || !apiKey.trim()) throw new GeminiRequestError('REQUEST_CONFIGURATION', 'Set a Gemini model and API key in Settings.');
  // A header cannot carry what a URL could encode (spaces, invisible or non-ASCII
  // characters from a bad paste); say so instead of failing as a network error.
  const key = apiKey.trim();
  if (/[^\x21-\x7e]/.test(key)) throw new GeminiRequestError('REQUEST_CONFIGURATION', 'The saved API key contains characters that are not allowed. Paste it again in Settings.');
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new GeminiRequestError('REQUEST_CONFIGURATION', 'Request timeout must be positive.');
  const headers = { 'Content-Type': 'application/json', 'x-goog-api-key': key };
  const attempts = Math.min(3, Math.max(1, Math.floor(maxAttempts) || 1));
  const body = JSON.stringify(payload);
  for (let attempt = 0; attempt < attempts; attempt++) {
    if (signal?.aborted) throw cancelled();
    // An optional absolute deadline (ms since epoch) bounds every attempt and retry,
    // so a request cannot outlast its overall time limit.
    const remaining = deadline - now();
    if (remaining < MIN_ATTEMPT_MS) throw deadlineReached();
    const attemptTimeoutMs = Math.min(timeoutMs, remaining);
    const controller = new AbortController();
    let timer;
    let onAbort;
    const interrupt = new Promise((_, reject) => {
      onAbort = () => { controller.abort(); reject(cancelled()); };
      signal?.addEventListener('abort', onAbort, { once: true });
      timer = setTimeout(() => {
        controller.abort();
        reject(new GeminiRequestError('REQUEST_TIMEOUT', 'Gemini request timed out. Please try again.'));
      }, attemptTimeoutMs);
    });
    let failure;
    try {
      return await Promise.race([interrupt, (async () => {
        let response;
        try {
          response = await fetchImpl(url, { method: 'POST', headers, body, signal: controller.signal });
        } catch {
          throw new GeminiRequestError('REQUEST_NETWORK', 'Could not reach Gemini. Check your connection and try again.');
        }
        if (!response.ok) {
          // Never include provider bodies: they can echo prompts and credentials.
          throw new GeminiRequestError('REQUEST_HTTP', response.status === 401 || response.status === 403
            ? 'Gemini authentication failed. Check your API key and access.'
            : response.status === 400 || response.status === 422
              ? `Gemini rejected the request (HTTP ${response.status}). Check the selected model and conversation settings.`
              : response.status === 429
                ? 'Gemini rate limit reached. Wait before trying again.'
                : `Gemini request failed (HTTP ${response.status}). Please try again later.`, response.status);
        }
        try { return await response.json(); }
        catch { throw new GeminiRequestError('REQUEST_RESPONSE', 'Gemini returned an invalid JSON response.'); }
      })()]);
    } catch (error) {
      failure = signal?.aborted ? cancelled() : error;
    } finally {
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
    }
    const retryable = failure.code === 'REQUEST_NETWORK' || failure.code === 'REQUEST_TIMEOUT'
      || (failure.code === 'REQUEST_HTTP' && transientStatuses.has(failure.status));
    if (!retryable || attempt === attempts - 1) throw failure;
    const delay = Math.min(10000, Math.max(0, backoffMs) * (2 ** attempt)) * (0.5 + random() * 0.5);
    if (deadline - now() - delay < MIN_ATTEMPT_MS) throw failure;
    await sleep(delay, signal);
  }
}
