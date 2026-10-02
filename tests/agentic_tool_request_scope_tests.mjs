import assert from 'node:assert/strict';
import { createAgenticTools } from '../src/taskpane/modules/commands/agentic-tools.js';

const oldController = new AbortController();
const newController = new AbortController();
let oldSignalReads = 0;
let newSignalReads = 0;
let trackingChanges = 0;
let releaseOldRead;
let oldReadStarted;
let firstWordRun = true;
let fetchCalls = 0;

const firstReadStarted = new Promise(resolve => { oldReadStarted = resolve; });
const hostContext = {
  document: {
    body: {
      paragraphs: {
        items: [],
        load() {}
      }
    }
  },
  async sync() {}
};

globalThis.Word = {
  async run(callback) {
    if (firstWordRun) {
      firstWordRun = false;
      oldReadStarted();
      await new Promise(resolve => { releaseOldRead = resolve; });
    }
    return callback(hostContext);
  }
};

const originalFetch = globalThis.fetch;
globalThis.fetch = async url => {
  fetchCalls++;
  assert.match(url, /key=new-key$/, 'only the newer request should reach Gemini');
  return {
    ok: true,
    async json() {
      return { candidates: [{ content: { parts: [{ text: '[]' }] } }] };
    }
  };
};

function dependencies(apiKey, getSignal) {
  return {
    getRequestSignal: getSignal,
    loadApiKey: () => apiKey,
    loadModel: () => 'gemini-2.0-flash',
    loadSystemMessage: () => '',
    loadRedlineSetting: () => false,
    loadRedlineAuthor: () => 'Tester',
    setChangeTrackingForAi: async () => { trackingChanges++; return {}; },
    restoreChangeTracking: async () => {},
    SAFETY_SETTINGS_BLOCK_NONE: []
  };
}

try {
  const oldTools = createAgenticTools(dependencies('old-key', () => {
    oldSignalReads++;
    return oldController.signal;
  }));
  const oldPending = oldTools.executeRedline('Change the text', '[P1] Original');
  await firstReadStarted;

  oldController.abort();

  const newTools = createAgenticTools(dependencies('new-key', () => {
    newSignalReads++;
    return newController.signal;
  }));
  const newResult = await newTools.executeRedline('Change the text', '[P1] Original');
  assert.match(newResult.message, /no changes to suggest/i);
  assert.equal(fetchCalls, 1);
  assert.ok(newSignalReads > 0, 'the new executor uses its own valid request signal');
  const newSignalReadsAfterNewCall = newSignalReads;

  releaseOldRead();
  const oldResult = await oldPending;
  assert.equal(oldResult.status, 'error');
  assert.equal(oldResult.error?.code, 'REQUEST_CANCELLED');
  assert.ok(oldSignalReads > 0, 'the resumed old executor still reads its original signal');
  assert.equal(newSignalReads, newSignalReadsAfterNewCall, 'resuming the old executor must not read the newer signal');
  assert.equal(fetchCalls, 1, 'the cancelled old executor must not send a provider request with the newer credentials');
  assert.equal(trackingChanges, 0, 'neither path may begin a document mutation for an empty/aborted response');

  const missingKeyTools = createAgenticTools(dependencies('', () => newController.signal));
  const missingKeyResult = await missingKeyTools.executeRedline('Change the text', '[P1] Original');
  assert.equal(missingKeyResult.error?.code, 'MISSING_API_KEY');
  assert.equal(fetchCalls, 1, 'a missing-key instance must not reuse another instance API key');

  console.log('PASS: request-scoped agentic tool dependencies remain isolated across cancellation');
} finally {
  releaseOldRead?.();
  globalThis.fetch = originalFetch;
  delete globalThis.Word;
}
