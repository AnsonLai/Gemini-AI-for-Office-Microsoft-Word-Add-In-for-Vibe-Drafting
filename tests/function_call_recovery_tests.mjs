import assert from 'node:assert/strict';
import { recoverFunctionCall } from '../src/taskpane/modules/chat/function-call-recovery.js';

const editArgs = {
  edits: [{
    type: 'replace',
    target: { handle: 'opaque-handle' },
    content: { text: 'Replace {old} with the approved wording.' },
  }],
};

assert.deepEqual(
  recoverFunctionCall(`Diagnostic: apply_document_edits(${JSON.stringify(editArgs)})`, ['inspect_document', 'apply_document_edits']),
  { name: 'apply_document_edits', args: editArgs },
  'complete nested JSON arguments are recovered without flattening braces or guessing fields'
);

const inspectArgs = { search: 'termination', around: 1, limit: 5 };
assert.deepEqual(
  recoverFunctionCall(`Malformed call inspect_document ${JSON.stringify(inspectArgs)}`, new Set(['inspect_document'])),
  { name: 'inspect_document', args: inspectArgs }
);

assert.equal(
  recoverFunctionCall('apply_document_edits({edits: [{type: replace}]})', ['apply_document_edits']),
  null,
  'unquoted pseudo-JSON is not repaired by guessing'
);
assert.equal(
  recoverFunctionCall('apply_document_edits({"edits":[{"type":"replace"}]', ['apply_document_edits']),
  null,
  'incomplete JSON objects are not recovered'
);
assert.equal(recoverFunctionCall('legacy_edit({"instruction":"fix wording"})', ['inspect_document', 'apply_document_edits']), null);
assert.equal(recoverFunctionCall('apply_document_edits({"edits":[]})', ['inspect_document']), null);
assert.equal(recoverFunctionCall(undefined, ['inspect_document']), null);

console.log('PASS: complete JSON recovery is limited to inspect_document and apply_document_edits');
