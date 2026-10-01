import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { createMutationObserver, requiresMutationInspection } from '../src/taskpane/modules/commands/mutation-outcome.js';
import { initAgenticTools, executeComment, executeInsertListItem, executeRedline } from '../src/taskpane/modules/commands/agentic-tools.js';

const observer = createMutationObserver();
observer.attempt();
await observer.sync({ async sync() {} });
observer.attempt();
const unknown = Object.assign(new Error('Unknown host failure'), { code: 'FUTURE_WORD_ERROR' });
assert.equal(observer.result().status, 'error');
assert.equal(observer.result().success, false);
assert.equal(observer.result(unknown).mutationOutcome, 'partial');
assert.equal(observer.result(unknown).written, true);
assert.equal(observer.result(unknown).error.code, 'FUTURE_WORD_ERROR');
assert.ok(requiresMutationInspection(observer.result(unknown)));
assert.deepEqual(observer.result(unknown).receipts, [], 'consumer observations are not library receipts');

initAgenticTools({
  loadApiKey: () => 'fake-test-key', loadModel: () => 'test-model', loadSystemMessage: () => '',
  loadRedlineSetting: () => true, loadRedlineAuthor: () => 'Tester',
  setChangeTrackingForAi: async () => ({}), restoreChangeTracking: async () => {},
  SAFETY_SETTINGS_BLOCK_NONE: [], API_LIMITS: {}
});
const oldFetch = globalThis.fetch;
const oldWord = globalThis.Word;
try {
  const proposals = [1, 2, 3].map(paragraphIndex => ({ paragraphIndex, textToFind: 'clause', commentContent: 'Review.' }));
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ candidates: [{ content: { parts: [{ text: JSON.stringify(proposals) }] } }] }) });
  const writes = [];
  let thirdReads = 0;
  const paragraphs = [1, 2, 3].map(index => ({
    text: `Test clause ${index}.`, font: { name: 'Calibri' }, load() {},
    getOoxml() {
      if (index === 3) thirdReads++;
      return { value: buildDocumentFragmentPackage(`<w:p><w:r><w:t>Test clause ${index}.</w:t></w:r></w:p>`, { appendTrailingParagraph: false }) };
    },
    insertOoxml() {
      writes.push(index);
      if (index === 2) throw unknown;
    }
  }));
  globalThis.Word = { run: async callback => callback({ async sync() {}, document: {
    body: {
      getOoxml: () => ({ value: buildDocumentFragmentPackage('<w:p><w:r><w:t>Plain</w:t></w:r></w:p>', { appendTrailingParagraph: false }) }),
      insertOoxml() { throw new Error('Plain-anchor routing must retain native insertion'); },
      paragraphs: { items: paragraphs, load() {} }
    }
  } }) };
  const result = await executeComment('Review clauses', '[P1] Test clause 1.');
  assert.equal(result.status, 'error');
  assert.equal(result.mutationOutcome, 'partial');
  assert.equal(result.written, true);
  assert.equal(result.confirmedHostWrites, 1);
  assert.deepEqual(writes, [1, 2]);
  assert.equal(thirdReads, 0, 'stop subsequent mutations after unconfirmed insertion');
  assert.equal(result.operationResults[1].hostError.code, 'FUTURE_WORD_ERROR');
  assert.ok(requiresMutationInspection(result));

  // A native list insertion reports an uncertain write, with the original code.
  paragraphs[0].getOoxml = () => ({ value: '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>' });
  paragraphs[0].insertParagraph = () => { throw unknown; };
  const native = await executeInsertListItem(1, 'New item', 0);
  assert.equal(native.success, false);
  assert.equal(native.writeAttempted, true);
  assert.equal(native.mutationOutcome, 'indeterminate');
  assert.equal(native.error.code, 'FUTURE_WORD_ERROR');

  let providerCalls = 0;
  globalThis.fetch = async () => {
    providerCalls++;
    return { ok: true, json: async () => ({ candidates: [{ content: { parts: [{ text: JSON.stringify([
      { operation: 'edit_paragraph', paragraphIndex: 1, anchorText: 'Stale clause.', replacements: [{ find: 'Stale', replace: 'Updated' }] }
    ]) }] } }] }) };
  };
  const actual = buildDocumentFragmentPackage('<w:p><w:r><w:t>Actual clause.</w:t></w:r></w:p>', { appendTrailingParagraph: false });
  let engineWrites = 0;
  globalThis.Word = { run: async callback => callback({ async sync() {}, document: {
    changeTrackingMode: 'Off', load() {}, body: {
      paragraphs: { items: paragraphs, load() {} },
      getOoxml: () => ({ value: actual }), insertOoxml() { engineWrites++; }
    }
  } }) };
  const refused = await executeRedline('Update the term', '[P1] Stale clause.');
  assert.equal(refused.status, 'error');
  assert.equal(refused.writeAttempted, false);
  assert.equal(engineWrites, 0);
  assert.equal(providerCalls, 1, 'engine refusal must not trigger corrective provider replay');
} finally {
  globalThis.fetch = oldFetch;
  globalThis.Word = oldWord;
}
console.log('PASS: mutation outcomes preserve host uncertainty and stop remaining writes');
