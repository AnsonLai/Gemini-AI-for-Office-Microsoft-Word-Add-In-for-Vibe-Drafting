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
const transientStatuses = new Set([408, 429, 500, 502, 503, 504]);

function pause(ms, signal) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) { reject(cancelled()); return; }
    const onAbort = () => { clearTimeout(timer); signal.removeEventListener('abort', onAbort); reject(cancelled()); };
    const timer = setTimeout(() => { signal?.removeEventListener('abort', onAbort); resolve(); }, ms);
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

export function geminiEndpoint(model, apiKey) {
  if (!model || !apiKey) throw new GeminiRequestError('REQUEST_CONFIGURATION', 'Set a Gemini model and API key in Settings.');
  return `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`;
}

export async function requestGemini({ model, apiKey, payload, ...options }) {
  return requestGeminiJson(geminiEndpoint(model, apiKey), payload, options);
}

export async function requestGeminiJson(url, payload, {
  signal, timeoutMs = 90000, maxAttempts = 3, backoffMs = 1000,
  fetchImpl = globalThis.fetch, sleep = pause, random = Math.random,
} = {}) {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new GeminiRequestError('REQUEST_CONFIGURATION', 'Request timeout must be positive.');
  const attempts = Math.min(3, Math.max(1, Math.floor(maxAttempts) || 1));
  const body = JSON.stringify(payload);
  for (let attempt = 0; attempt < attempts; attempt++) {
    if (signal?.aborted) throw cancelled();
    const controller = new AbortController();
    let timer;
    let onAbort;
    const interrupt = new Promise((_, reject) => {
      onAbort = () => { controller.abort(); reject(cancelled()); };
      signal?.addEventListener('abort', onAbort, { once: true });
      timer = setTimeout(() => {
        controller.abort();
        reject(new GeminiRequestError('REQUEST_TIMEOUT', 'Gemini request timed out. Please try again.'));
      }, timeoutMs);
    });
    let failure;
    try {
      return await Promise.race([interrupt, (async () => {
        let response;
        try {
          response = await fetchImpl(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body, signal: controller.signal });
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
    await sleep(delay, signal);
  }
}
