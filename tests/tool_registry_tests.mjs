import assert from 'node:assert/strict';
import { createToolRegistry, getToolDeclarations } from '../src/taskpane/modules/chat/tool-registry.js';
import { runDocumentAgentLoop } from '../src/taskpane/modules/chat/document-agent-loop.js';

const EXPECTED_NAMES = ['inspect_document', 'apply_document_edits'];
const LEGACY_NAMES = [
  'apply_redlines', 'insert_comment', 'highlight_text', 'perform_research',
  'navigate_to_section', 'edit_list', 'insert_list_item', 'edit_table',
  'edit_section', 'convert_headers_to_list', 'get_selection_stats'
];

function testDeclarationsAndRecoveryNames() {
  const declarations = getToolDeclarations();
  assert.equal(declarations.length, 1);
  const tools = declarations[0].function_declarations;
  assert.deepEqual(tools.map(tool => tool.name), EXPECTED_NAMES);

  const inspect = tools.find(tool => tool.name === 'inspect_document');
  assert.deepEqual(Object.keys(inspect.parameters.properties), [
    'search', 'around', 'limit', 'cursor', 'includeDocumentEnd'
  ]);
  const apply = tools.find(tool => tool.name === 'apply_document_edits');
  assert.deepEqual(apply.parameters.required, ['edits']);
  assert.deepEqual(Object.keys(apply.parameters.properties), ['edits']);
  const edit = apply.parameters.properties.edits.items;
  assert.deepEqual(edit.required, ['type', 'target']);
  assert.deepEqual(edit.properties.type.enum, [
    'replace_text', 'rewrite_paragraph', 'delete_paragraph', 'insert_paragraph',
    'append_paragraph', 'insert_list_item', 'comment', 'format_text'
  ]);
  assert.deepEqual(Object.keys(edit.properties), [
    'type', 'target', 'find', 'replace', 'text', 'occurrence', 'position', 'anchor', 'properties'
  ]);
  assert.deepEqual(Object.keys(edit.properties.properties.properties), [
    'bold', 'italic', 'underline', 'strikethrough', 'highlight'
  ]);
  assert.deepEqual(edit.properties.properties.properties.highlight.enum, ['yellow', 'none']);
  assert.doesNotMatch(JSON.stringify(tools), /font|rawOp|modelIndex|transport|options/i,
    'the model protocol must not expose font styling or host/transport internals');

  const registry = createToolRegistry({});
  assert.deepEqual(registry.recoveryToolNames, EXPECTED_NAMES);
  assert.equal(Object.isFrozen(registry.recoveryToolNames), true);
  tools.pop();
  assert.equal(getToolDeclarations()[0].function_declarations.length, 2,
    'caller mutations must not alter the authoritative definitions');
  assert.equal(registry.describe('inspect_document').mutating, false);
  assert.equal(registry.describe('apply_document_edits').mutating, true);
  assert.equal(registry.describe('apply_document_edits').withCheckpoint, undefined,
    'the adapter persists a source-guarded checkpoint inside the atomic batch');
}

async function testDirectArgumentValidationAndExecutorMapping() {
  const calls = [];
  const registry = createToolRegistry({
    inspectDocument: async args => {
      calls.push({ executor: 'inspectDocument', args });
      return { status: 'ok', matches: [{ handle: 'h:1', text: 'A sentence.' }] };
    },
    applyDocumentEdits: async args => {
      calls.push({ executor: 'applyDocumentEdits', args });
      return {
        status: 'ok', success: true, completion: true, verified: true,
        written: true, writeAttempted: true, mutationOutcome: 'applied',
        requestedCount: args.edits.length, message: 'Applied and verified 1 requested edit(s).'
      };
    }
  });
  const inspectArgs = { search: 'sentence', around: 1, limit: 2 };
  const editArgs = { edits: [{ type: 'replace_text', target: 'h:1', find: 'sentence', replace: 'statement' }] };

  assert.equal(registry.validate('inspect_document', inspectArgs).valid, true);
  assert.equal(registry.validate('inspect_document', { ...inspectArgs, instruction: 'legacy' }).error.code, 'INVALID_TOOL_ARGUMENTS');
  for (const invalidInspectArgs of [
    { search: 42 }, { cursor: false }, { around: 1.5 }, { limit: '3' }, { includeDocumentEnd: 1 }
  ]) {
    assert.equal(registry.validate('inspect_document', invalidInspectArgs).error.code, 'INVALID_TOOL_ARGUMENTS',
      'inspection arguments must match their declared primitive types');
  }
  assert.equal(registry.validate('apply_document_edits', editArgs).valid, true);
  assert.equal(registry.validate('apply_document_edits', {}).error.code, 'INVALID_DOCUMENT_EDIT_BATCH');
  assert.equal(registry.validate('apply_document_edits', { edits: [] }).error.code, 'INVALID_DOCUMENT_EDIT_BATCH');
  assert.equal(registry.validate('apply_document_edits', { ...editArgs, options: {} }).error.code, 'INVALID_TOOL_ARGUMENTS');
  assert.equal(registry.validate('apply_document_edits', { edits: Array(41).fill(editArgs.edits[0]) }).error.code,
    'INVALID_DOCUMENT_EDIT_BATCH');

  const inspection = await registry.execute('inspect_document', inspectArgs);
  const applied = await registry.execute('apply_document_edits', editArgs);
  assert.deepEqual(calls, [
    { executor: 'inspectDocument', args: inspectArgs },
    { executor: 'applyDocumentEdits', args: editArgs }
  ], 'the direct protocol forwards its declared argument records without legacy normalization');
  assert.equal(inspection.success, true);
  assert.deepEqual(inspection.recordResult.matches, [{ handle: 'h:1', text: 'A sentence.' }]);
  assert.equal(applied.success, true);
  assert.equal(applied.mutation, applied.rawResult);
  assert.equal(applied.recordResult, applied.rawResult);
  assert.equal(applied.hostWriteEvidence.written, true);
  assert.equal(applied.hostWriteEvidence.mutationOutcome, 'applied');
  assert.equal(applied.completionText, applied.rawResult.message);

  const unavailable = await createToolRegistry({}).execute('inspect_document', {});
  assert.equal(unavailable.rawResult.error.code, 'TOOL_UNAVAILABLE');
  assert.equal(unavailable.success, false);
}

async function runApplyReceipt(rawReceipt) {
  const registry = createToolRegistry({ applyDocumentEdits: async () => rawReceipt });
  const state = {
    history: [{ role: 'user', parts: [{ text: 'Make the requested edit.' }] }],
    docText: '[P1] Original', docSourceBaseline: [], records: []
  };
  const events = [];
  const replies = [
    { candidates: [{ content: { parts: [{ functionCall: {
      name: 'apply_document_edits', args: { edits: [{ type: 'replace_text', target: 'h:1', find: 'Old', replace: 'New' }] }
    } }] } }] },
    { candidates: [{ content: { parts: [{ text: 'The request is complete.' }] } }] }
  ];
  let refreshes = 0;
  await runDocumentAgentLoop({
    state, registry, loadRegistry: async () => registry,
    requestModel: async () => replies.shift(),
    refreshContext: async () => {
      refreshes++;
      return { formattedText: '[P1] New', sourceBaseline: [] };
    },
    createCheckpoint: async () => assert.fail('direct apply uses its source-guarded atomic batch checkpoint'),
    requestStartTime: 0, now: () => 1, onEvent: event => events.push(event)
  });
  const toolEnd = events.find(event => event.type === 'tool_end' && event.name === 'apply_document_edits');
  assert.ok(toolEnd, 'the tool receipt reaches the taskpane event/UI boundary');
  assert.deepEqual(state.records[0].result, rawReceipt, 'recovery history preserves the structured host receipt');
  assert.equal(toolEnd.completionText, rawReceipt.message);
  assert.equal(toolEnd.display.show, true, 'the direct mutation receipt remains visible to the user');
  assert.equal(state.records[0].written, rawReceipt.written === true);
  assert.equal(state.records[0].writeAttempted, rawReceipt.writeAttempted === true);
  assert.equal(state.records[0].mutationOutcome, rawReceipt.mutationOutcome);
  return { state, events, toolEnd, refreshes };
}

async function testVerifiedNoopRefusedAndUncertainReceiptsReachUi() {
  const complete = {
    status: 'ok', success: true, completion: true, verified: true,
    written: true, writeAttempted: true, mutationOutcome: 'applied',
    message: 'Applied and verified 1 requested edit(s).'
  };
  const applied = await runApplyReceipt(complete);
  assert.equal(applied.state.records[0].success, true);
  assert.equal(applied.state.records[0].requiresInspection, false);
  assert.equal(applied.refreshes, 1, 'confirmed writes refresh document context before another provider turn');

  const noop = {
    status: 'ok', success: true, completion: true, verified: true,
    written: false, writeAttempted: false, mutationOutcome: 'noop',
    message: 'The requested state is already satisfied; no document write was needed.'
  };
  const noOpRun = await runApplyReceipt(noop);
  assert.equal(noOpRun.state.records[0].success, true, 'verified no-op counts as complete without a host write');
  assert.equal(noOpRun.state.records[0].written, false);
  assert.equal(noOpRun.refreshes, 0, 'a verified no-op does not trigger a needless source refresh');

  const refused = {
    status: 'error', success: false, completion: false, verified: false,
    written: false, writeAttempted: false, mutationOutcome: 'refused',
    message: 'No changes made. The selected text was not found.'
  };
  const refusedRun = await runApplyReceipt(refused);
  assert.equal(refusedRun.state.records[0].success, false);
  assert.equal(refusedRun.state.records[0].requiresInspection, false);
  assert.equal(refusedRun.toolEnd.completionText, refused.message);

  const uncertain = {
    status: 'error', success: false, completion: false, verified: false,
    written: false, writeAttempted: true, mutationOutcome: 'indeterminate',
    message: 'An edit was attempted but completion could not be verified. Inspect the document before continuing; do not replay it.'
  };
  const uncertainRun = await runApplyReceipt(uncertain);
  assert.equal(uncertainRun.state.records[0].success, false);
  assert.equal(uncertainRun.state.records[0].requiresInspection, true);
  assert.equal(uncertainRun.state.stopReason, 'requires_inspection');
  assert.equal(uncertainRun.refreshes, 0, 'uncertain outcomes are never presented as confirmed writes');
}

async function testLegacyCapabilitiesRefuseBeforeExecutor() {
  let executorCalls = 0;
  const registry = createToolRegistry({
    inspectDocument: async () => { executorCalls++; },
    applyDocumentEdits: async () => { executorCalls++; }
  });
  for (const name of LEGACY_NAMES) {
    const validation = registry.validate(name, {});
    assert.equal(validation.valid, false, `${name} is not in the new authoritative protocol`);
    assert.equal(validation.error.code, 'UNSUPPORTED_TOOL');
    const result = await registry.execute(name, {});
    assert.equal(result.rawResult.error.code, 'UNSUPPORTED_TOOL');
    assert.equal(result.hostWriteEvidence.writeAttempted, false);
    assert.equal(result.hostWriteEvidence.written, false);
  }
  assert.equal(executorCalls, 0, 'unsupported legacy tools are refused before reaching any document executor');
}

testDeclarationsAndRecoveryNames();
await testDirectArgumentValidationAndExecutorMapping();
await testVerifiedNoopRefusedAndUncertainReceiptsReachUi();
await testLegacyCapabilitiesRefuseBeforeExecutor();
console.log('Tool registry tests passed.');
