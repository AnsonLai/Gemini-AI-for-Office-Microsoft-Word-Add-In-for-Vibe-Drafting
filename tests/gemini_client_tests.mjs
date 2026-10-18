import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { requestGemini, requestGeminiJson, geminiEndpoint } from '../src/taskpane/modules/chat/gemini-client.js';
import { runDocumentAgentLoop } from '../src/taskpane/modules/chat/document-agent-loop.js';
import { createToolRegistry } from '../src/taskpane/modules/chat/tool-registry.js';

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
for (const file of [
  'src/taskpane/taskpane.js',
  'src/taskpane/modules/chat/document-agent-loop.js',
  'src/taskpane/modules/chat/tool-registry.js',
  'src/taskpane/modules/chat/document-tools.js'
]) {
  const source = readFileSync(new URL(`../${file}`, import.meta.url), 'utf8');
  assert.doesNotMatch(source, /\bfetch\s*\(/, `${file}: direct Gemini transport bypass`);
}
const taskpane = readFileSync(new URL('../src/taskpane/taskpane.js', import.meta.url), 'utf8');
const providerWrapper = taskpane.match(/async function callGeminiWithRetry\([\s\S]*?\n\}/)?.[0];
assert.ok(providerWrapper, 'sendChatMessage must route provider calls through its audited wrapper');
assert.match(providerWrapper, /return requestGeminiJson\(url, payload,/);
assert.match(providerWrapper, /signal = currentRequestController\?\.signal/);
assert.match(providerWrapper, /\n\s*signal,/);
assert.match(providerWrapper, /timeoutMs:\s*TIMEOUT_LIMITS\.FETCH_TIMEOUT_MS/);

// Exercise the extracted loop through the same audited JSON transport. A tool
// call is executed once; a later provider error and its transport retries must
// stop the loop without replaying the tool.
{
  const state = {
    history: [{ role: 'user', parts: [{ text: 'Correct the typo.' }] }],
    docText: '[P1] Original',
    docSourceBaseline: null,
    records: []
  };
  let modelRequests = 0;
  let transportCalls = 0;
  let executions = 0;
  let capturedEdits;
  let checkpointCalls = 0;
  const registry = createToolRegistry({
    applyDocumentEdits: async args => {
      executions++;
      capturedEdits = args.edits;
      return {
        status: 'ok', success: true, completion: true, verified: true,
        message: 'Edit applied and verified.', written: true, writeAttempted: true,
        mutationOutcome: 'applied', checkpointIndex: 99,
        receipts: [{ operationIndex: 0, status: 'applied', verified: true }]
      };
    }
  });
  const fetchImpl = async (_url, init) => {
    transportCalls++;
    const requestBody = JSON.parse(init.body);
    assert.ok(Array.isArray(requestBody.contents), 'the audited client sends the loop history');
    if (transportCalls === 1) {
      return {
        ok: true,
        json: async () => ({ candidates: [{ content: { parts: [{ functionCall: {
          name: 'apply_document_edits', args: { edits: [{
            type: 'replace_text', target: 'opaque-handle', find: 'Original', replace: 'Corrected'
          }] }
        } }] } }] })
      };
    }
    return { ok: false, status: 503 };
  };

  await runDocumentAgentLoop({
    state,
    registry,
    loadRegistry: async () => registry,
    requestModel: history => {
      modelRequests++;
      return requestGeminiJson(url, { contents: history }, {
        fetchImpl, sleep: async () => {}, random: () => 0, maxAttempts: 3, timeoutMs: 1000
      });
    },
    refreshContext: async () => ({
      formattedText: '[P1] Corrected',
      sourceBaseline: [{ index: 1, exactText: 'Corrected', fingerprint: 'second' }]
    }),
    createCheckpoint: async () => { checkpointCalls++; return 101; },
    requestStartTime: 0,
    now: () => 1,
    totalTimeoutMs: 10000
  });

  assert.equal(state.stopReason, 'provider_error');
  assert.equal(modelRequests, 2, 'the loop stops after the provider error instead of replaying the tool');
  assert.equal(transportCalls, 4, 'only the audited client retries the failed provider request');
  assert.equal(executions, 1, 'the already completed redline is never replayed');
  assert.equal(capturedEdits[0].replace, 'Corrected', 'the direct executor receives the typed batch once');
  assert.equal(checkpointCalls, 0, 'the source-aligned apply executor owns checkpoint capture');
  assert.equal(state.records[0].checkpointIndex, 99, 'the executor-owned checkpoint evidence survives the loop event');
  assert.equal(state.records[0].result.receipts[0].verified, true, 'the direct tool receipt is retained for post-call recovery');
  assert.equal(state.records[0].written, true, 'provider failure preserves confirmed host-write evidence');
}

console.log('PASS: centralized Gemini requests, bounded recovery, cancellation, timeout, safe diagnostics and no tool replay');
