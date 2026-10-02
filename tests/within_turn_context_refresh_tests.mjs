import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { appendFunctionExchange, sanitizeHistory } from '../src/taskpane/modules/chat/chat-history.js';
import { appendRefreshedDocumentContext, isNoWriteStaleContextRefusal } from '../src/taskpane/modules/chat/refreshed-document-context.js';
import { applyRedlineChangesToWordContext } from '../src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js';
import { captureWordSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';

function makePackage() {
  return buildDocumentFragmentPackage(
    '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="AAAAB001">'
      + '<w:r><w:t>First paragraph stays unchanged.</w:t></w:r></w:p>'
      + '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="AAAAB002">'
      + '<w:r><w:t>Final paragraph stays unchanged.</w:t></w:r></w:p>',
    { appendTrailingParagraph: false }
  );
}

const oldWord = globalThis.Word;
globalThis.Word = { InsertLocation: { replace: 'Replace' } };
try {
  let currentOoxml = makePackage();
  const writes = [];
  const body = {
    getOoxml() { return { value: currentOoxml }; },
    insertOoxml(value) { writes.push(value); currentOoxml = value; }
  };
  const context = { document: { body }, async sync() {} };
  const oldBaseline = captureWordSourceBaseline(currentOoxml);

  const underline = await applyRedlineChangesToWordContext(context, [{
    operation: 'edit_paragraph',
    paragraphIndex: 1,
    replacements: [{ find: 'stays unchanged', replace: '++stays unchanged++' }]
  }], { author: 'Editor', sourceBaseline: oldBaseline, onInfo() {}, onWarn() {} });
  assert.equal(underline.written, true, 'the first formatting operation must have a confirmed host write');

  const afterUnderlineBaseline = captureWordSourceBaseline(currentOoxml);
  const table = '| Mountain | River | Forest |\n|---|---|---|\n| Ocean | Valley | Canyon |\n| Meadow | Desert | Island |';
  const appended = await applyRedlineChangesToWordContext(context, [{
    operation: 'replace_paragraph', paragraphIndex: 3, content: table
  }], { author: 'Editor', sourceBaseline: afterUnderlineBaseline, onInfo() {}, onWarn() {} });
  assert.equal(appended.written, true, 'a later model turn should append the table using the refreshed baseline');
  assert.equal((currentOoxml.match(/<w:tr(?=[ >])/g) || []).length, 3);
  assert.equal((currentOoxml.match(/<w:tc(?=[ >])/g) || []).length, 9);

  const refreshedBaseline = captureWordSourceBaseline(currentOoxml);
  const nextChange = [{
    operation: 'edit_paragraph', paragraphIndex: refreshedBaseline.find(item => item.exactText === 'Final paragraph stays unchanged.').index,
    replacements: [{ find: 'Final paragraph', replace: 'Updated paragraph' }]
  }];
  const staleAttempt = await applyRedlineChangesToWordContext(context, nextChange, {
    author: 'Editor', sourceBaseline: oldBaseline, onInfo() {}, onWarn() {}
  });
  assert.equal(staleAttempt.error?.code, 'STALE_DOCUMENT_CONTEXT');
  assert.equal(staleAttempt.written, false, 'an old baseline must still refuse stale paragraph counts');

  const currentAttempt = await applyRedlineChangesToWordContext(context, nextChange, {
    author: 'Editor', sourceBaseline: refreshedBaseline, onInfo() {}, onWarn() {}
  });
  assert.equal(currentAttempt.written, true, 'the same edit should proceed with the refreshed baseline');
  assert.equal(writes.length, 3);

  const history = [{ role: 'user', parts: [{ text: 'Original request' }] }];
  const functionResponses = [{
    functionResponse: { name: 'apply_redlines', response: { content: [{ text: 'Edit applied.' }] } }
  }];
  appendRefreshedDocumentContext(functionResponses, '[P1|Normal] Updated paragraph');
  appendFunctionExchange(history,
    { role: 'model', parts: [{ functionCall: { name: 'apply_redlines', args: { instruction: 'edit' } } }] },
    { role: 'user', parts: functionResponses });
  const validated = sanitizeHistory(history);
  assert.equal(validated.length, 3, 'the refresher text must stay beside its matched function response');
  assert.match(validated[2].parts.at(-1).text, /Updated paragraph/);

  // Simulate Undo after the original request captured its source snapshot.
  currentOoxml = makePackage();
  const refusedAfterUndo = await applyRedlineChangesToWordContext(context, nextChange, {
    author: 'Editor', sourceBaseline: refreshedBaseline, onInfo() {}, onWarn() {}
  });
  assert.equal(isNoWriteStaleContextRefusal(refusedAfterUndo), true);
  const writeCountBeforeRecovery = writes.length;
  const freshAfterUndo = captureWordSourceBaseline(currentOoxml);
  const recoveryResponses = [{ functionResponse: { name: 'apply_redlines', response: { text: 'Source stale; no write.' } } }];
  appendRefreshedDocumentContext(recoveryResponses, '[P2] Final paragraph stays unchanged.', 'stale-refusal');
  assert.equal(writes.length, writeCountBeforeRecovery, 'refreshing context does not replay or apply the refused batch');
  assert.match(recoveryResponses.at(-1).text, /Replan.*do not reuse/s);
  const mixedResponses = [];
  appendRefreshedDocumentContext(mixedResponses, '[P2] Mixed context.', 'mixed');
  assert.match(mixedResponses[0].text, /Some edits.*were applied.*refused as stale before any write/s);
  assert.match(mixedResponses[0].text, /do not reuse the refused batch.*do not redo the edits that were already applied/s);
  assert.doesNotMatch(mixedResponses[0].text, /^The previous edit was refused/);
  const replanned = await applyRedlineChangesToWordContext(context, [{
    operation: 'edit_paragraph', paragraphIndex: 2,
    replacements: [{ find: 'Final paragraph', replace: 'Replanned paragraph' }]
  }], { author: 'Editor', sourceBaseline: freshAfterUndo, onInfo() {}, onWarn() {} });
  assert.equal(replanned.written, true, 'a newly planned edit against the post-Undo source can commit');
  for (const unsafe of [
    { ...refusedAfterUndo, writeAttempted: true, mutationOutcome: 'indeterminate' },
    { ...refusedAfterUndo, written: true, mutationOutcome: 'applied_with_host_error' },
    { ...refusedAfterUndo, error: { code: 'UNSUPPORTED_TABLE_FORMATTING' } },
    { error: { code: 'STALE_DOCUMENT_CONTEXT' } }
  ]) assert.equal(isNoWriteStaleContextRefusal(unsafe), false, 'unknown or attempted writes cannot trigger stale-refusal recovery');
} finally {
  globalThis.Word = oldWord;
}

console.log('PASS: within-turn context refresh tests');
