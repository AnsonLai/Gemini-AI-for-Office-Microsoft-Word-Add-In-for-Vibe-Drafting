import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { requestGemini, requestGeminiJson, geminiEndpoint } from '../src/taskpane/modules/chat/gemini-client.js';

const payload = { contents: [{ parts: [{ text: 'private document' }] }] };
const url = geminiEndpoint('configured-model', 'secret+key');
assert.match(url, /configured-model:generateContent\?key=secret%2Bkey$/);
const ok = { ok: true, json: async () => ({ candidates: [] }) };
let calls = 0;
const delays = [];
const options = {
  fetchImpl: async (_url, init) => {
    calls++;
    assert.equal(_url, url);
    assert.equal(init.body, JSON.stringify(payload));
    return calls < 3 ? { ok: false, status: calls === 1 ? 429 : 503 } : ok;
  },
  sleep: async ms => { delays.push(ms); }, random: () => 0,
};
assert.deepEqual(await requestGemini({ model: 'configured-model', apiKey: 'secret+key', payload, ...options }), { candidates: [] });
assert.equal(calls, 3);
assert.deepEqual(delays, [500, 1000]);

for (const status of [400, 401, 403, 404, 422]) {
  calls = 0;
  await assert.rejects(requestGeminiJson(url, payload, { fetchImpl: async () => {
    calls++; return { ok: false, status, text: async () => 'secret+key private document' };
  } }), error => error.code === 'REQUEST_HTTP' && error.status === status && !/secret|private/.test(error.message));
  assert.equal(calls, 1, `HTTP ${status} must not retry`);
}
calls = 0;
await assert.rejects(requestGeminiJson(url, payload, {
  maxAttempts: 99, fetchImpl: async () => { calls++; throw new TypeError('URL contains secret+key'); }, sleep: async () => {},
}), error => error.code === 'REQUEST_NETWORK' && !error.message.includes('secret'));
assert.equal(calls, 3, 'Hard maximum of three transport attempts');

calls = 0;
await assert.rejects(requestGeminiJson(url, payload, {
  fetchImpl: async () => { calls++; return { ok: true, json: async () => { throw new Error('private document'); } }; },
}), error => error.code === 'REQUEST_RESPONSE');
assert.equal(calls, 1);

const before = new AbortController(); before.abort();
await assert.rejects(requestGeminiJson(url, payload, { signal: before.signal, fetchImpl: () => assert.fail('Cancelled request sent') }), { code: 'REQUEST_CANCELLED' });
const during = new AbortController();
let internalSignal;
await assert.rejects(requestGeminiJson(url, payload, {
  signal: during.signal, fetchImpl: async (_url, init) => { internalSignal = init.signal; during.abort(); return new Promise(() => {}); },
}), { code: 'REQUEST_CANCELLED' });
assert.equal(internalSignal.aborted, true);

calls = 0;
await assert.rejects(requestGeminiJson(url, payload, {
  timeoutMs: 5, maxAttempts: 2, sleep: async () => {},
  fetchImpl: async (_url, init) => { calls++; internalSignal = init.signal; return new Promise(() => {}); },
}), { code: 'REQUEST_TIMEOUT' });
assert.equal(calls, 2);
assert.equal(internalSignal.aborted, true);
// Timeout covers response parsing, too, including implementations that ignore abort.
await assert.rejects(requestGeminiJson(url, payload, {
  timeoutMs: 5, maxAttempts: 1, fetchImpl: async () => ({ ok: true, json: () => new Promise(() => {}) }),
}), { code: 'REQUEST_TIMEOUT' });

const waiting = new AbortController();
calls = 0;
await assert.rejects(requestGeminiJson(url, payload, {
  signal: waiting.signal, backoffMs: 10000,
  fetchImpl: async () => { calls++; setTimeout(() => waiting.abort(), 1); return { ok: false, status: 503 }; },
}), { code: 'REQUEST_CANCELLED' });
assert.equal(calls, 1);

// Observe listener lifecycle directly rather than relying on a process-exit timeout.
const tracked = new AbortController();
let listeners = 0;
const add = tracked.signal.addEventListener.bind(tracked.signal);
const remove = tracked.signal.removeEventListener.bind(tracked.signal);
tracked.signal.addEventListener = (...args) => { listeners++; return add(...args); };
tracked.signal.removeEventListener = (...args) => { listeners--; return remove(...args); };
await requestGeminiJson(url, payload, { signal: tracked.signal, fetchImpl: async () => ok });
assert.equal(listeners, 0);
await assert.rejects(requestGeminiJson(url, payload, { signal: tracked.signal, fetchImpl: async () => ({ ok: false, status: 400 }) }));
assert.equal(listeners, 0);

// Keep every known caller on the audited boundary and avoid raw request diagnostics.
for (const file of ['src/taskpane/taskpane.js', 'src/taskpane/modules/commands/agentic-tools.js', 'browser-demo/demo.js', 'tests/evals/run-evals.mjs']) {
  const source = readFileSync(new URL(`../${file}`, import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\bfetch\s*\(/, `${file}: direct Gemini transport bypass`);
}
const taskpane = readFileSync(new URL('../src/taskpane/taskpane.js', import.meta.url), 'utf8');
assert.ok(taskpane.indexOf('if (toolsExecutedInCurrentRequest.length > 0)', taskpane.indexOf('catch (apiError)')) < taskpane.indexOf('throw apiError;', taskpane.indexOf('catch (apiError)')), 'Post-tool provider errors stop before history recovery');
console.log('PASS: centralized Gemini requests, bounded recovery, cancellation, timeout and safe diagnostics');
