import assert from 'node:assert/strict';
import fs from 'node:fs';
import { parse } from 'acorn';
import { runDocumentAgentLoop } from '../src/taskpane/modules/chat/document-agent-loop.js';
import { createToolRegistry } from '../src/taskpane/modules/chat/tool-registry.js';

const taskpanePath = new URL('../src/taskpane/taskpane.js', import.meta.url);
const taskpaneSource = fs.readFileSync(taskpanePath, 'utf8');
const taskpaneAst = parse(taskpaneSource, { ecmaVersion: 'latest', sourceType: 'module' });

function visit(node, callback) {
  if (!node || typeof node !== 'object') return;
  callback(node);
  for (const value of Object.values(node)) {
    if (Array.isArray(value)) value.forEach(child => visit(child, callback));
    else if (value && typeof value === 'object') visit(value, callback);
  }
}

let sendChatMessage = null;
visit(taskpaneAst, node => {
  if (node.type === 'FunctionDeclaration' && node.id?.name === 'sendChatMessage') sendChatMessage = node;
});
assert.ok(sendChatMessage, 'sendChatMessage must remain an inspectable host adapter');

// Strong descriptors belong to the source-bound inspection session, rather than loop state.
let loopInvocation, loopStateDeclaration;
visit(sendChatMessage.body, node => {
  if (node.type === 'CallExpression' && node.callee?.name === 'runDocumentAgentLoop') loopInvocation = node;
  if (node.type === 'VariableDeclarator' && node.id?.name === 'loopState') loopStateDeclaration = node;
});
assert.ok(loopInvocation && loopStateDeclaration);
assert.deepEqual(loopStateDeclaration.init.properties.map(p => p.key.name), ['history', 'records']);
assert.doesNotMatch(taskpaneSource, /sourceBaselineFrom|docSourceBaseline|body\.clear\(/);
const state = { history: [{ role: 'user', parts: [{ text: 'Revise the source paragraph.' }] }], records: [] };

const expectedEdits = {
  edits: [{ type: 'replace_text', target: 'opaque:source-handle', find: 'Source', replace: 'Updated' }]
};
let inspectedArgs = null;
let appliedArgs = null;
let applyArgumentCount = null;
let providerTurns = 0;
const directRegistry = createToolRegistry({
  inspectDocument: async (...args) => {
    inspectedArgs = args[0];
    assert.equal(args.length, 1, 'inspection executor receives its direct model arguments only');
    return { status: 'ok', matches: [{ handle: 'opaque:source-handle', text: 'Source paragraph' }] };
  },
  applyDocumentEdits: async (...args) => {
    appliedArgs = args[0];
    applyArgumentCount = args.length;
    return {
      status: 'ok', success: true, completion: true, verified: true,
      written: true, writeAttempted: true, mutationOutcome: 'applied',
      message: 'Applied and verified 1 requested edit(s).'
    };
  }
});
const replies = [
  { candidates: [{ content: { parts: [{ functionCall: {
    name: 'inspect_document', args: { search: 'Source paragraph', around: 1, limit: 2 }
  } }] } }] },
  { candidates: [{ content: { parts: [{ functionCall: {
    name: 'apply_document_edits', args: expectedEdits
  } }] } }] },
  { candidates: [{ content: { parts: [{ text: 'Done.' }] } }] }
];
const refreshedBaseline = [{ index: 1, exactText: 'Updated paragraph', fingerprint: 'fnv1a32:updated' }];
let refreshCount = 0;
await runDocumentAgentLoop({
  state,
  registry: directRegistry,
  loadRegistry: async () => directRegistry,
  requestModel: async () => {
    providerTurns++;
    return replies.shift();
  },
  refreshContext: async () => {
    refreshCount++;
    return { formattedText: '[P1] Updated paragraph', sourceBaseline: refreshedBaseline };
  },
  createCheckpoint: async () => assert.fail('the direct edit batch checkpoints inside its atomic source-guarded adapter'),
  requestStartTime: 0,
  now: () => 1
});
assert.deepEqual(inspectedArgs, { search: 'Source paragraph', around: 1, limit: 2 });
assert.deepEqual(appliedArgs, expectedEdits,
  'the loop forwards the inspected opaque handle and direct edit argument record unchanged');
assert.equal(applyArgumentCount, 1, 'the executor no longer receives document text or a separate baseline');
assert.equal(providerTurns, 3, 'inspection, one atomic apply batch, and the final response each use one model turn');
assert.equal(state.records.filter(record => record.mutating).length, 1);
assert.ok(state.history.some(turn => turn.parts.some(part => part.text?.includes('Updated paragraph'))), 'fresh post-write context reaches the next model turn');
assert.equal(refreshCount, 1);

const unsupportedRegistry = createToolRegistry({
  inspectDocument: async () => assert.fail('legacy navigation cannot execute inspection'),
  applyDocumentEdits: async () => assert.fail('legacy navigation cannot execute a write')
});
assert.equal(unsupportedRegistry.recoveryToolNames.includes('navigate_to_section'), false,
  'the former navigation capability is not advertised or accepted for malformed-call recovery');
assert.equal(unsupportedRegistry.validate('navigate_to_section', { instruction: 'go to section' }).error.code,
  'UNSUPPORTED_TOOL');
const unsupportedNavigation = await unsupportedRegistry.execute('navigate_to_section', { instruction: 'go to section' });
assert.equal(unsupportedNavigation.rawResult.error.code, 'UNSUPPORTED_TOOL');
assert.equal(unsupportedNavigation.hostWriteEvidence.writeAttempted, false);

// Execute the taskpane's identity guards directly so refresh/new-request races
// cannot let a superseded request restore history, overwrite UI, or unlock the
// new request's controls.
const isControllerComparison = (node, operator) => node?.type === 'BinaryExpression'
  && node.operator === operator
  && node.left?.type === 'Identifier'
  && node.left.name === 'currentRequestController'
  && node.right?.type === 'Identifier'
  && node.right.name === 'requestController';
const containsNode = (root, predicate) => {
  let found = false;
  visit(root, node => { if (predicate(node)) found = true; });
  return found;
};

let initialReadSource = null;
let firstUserHistoryPush = null;
const identityConditionals = [];
visit(sendChatMessage.body, node => {
  if (node.type === 'VariableDeclarator'
    && node.id?.name === 'documentSource'
    && node.init?.type === 'AwaitExpression'
    && node.init.argument?.type === 'CallExpression'
    && node.init.argument.callee?.type === 'Identifier'
    && node.init.argument.callee.name === 'readSource') {
    initialReadSource = node.init.argument;
  }
  if (node.type === 'CallExpression'
    && node.callee?.type === 'MemberExpression'
    && node.callee.object?.name === 'chatHistory'
    && node.callee.property?.name === 'push'
    && (!firstUserHistoryPush || node.start < firstUserHistoryPush.start)) {
    firstUserHistoryPush = node;
  }
  if (node.type === 'IfStatement'
    && (isControllerComparison(node.test, '===') || isControllerComparison(node.test, '!=='))) {
    identityConditionals.push(node);
  }
});

const postContextGuard = identityConditionals.find(node =>
  isControllerComparison(node.test, '!==')
  && node.consequent.type === 'ReturnStatement'
  && initialReadSource && node.start > initialReadSource.end
  && firstUserHistoryPush && node.end < firstUserHistoryPush.start
);
assert.ok(initialReadSource && firstUserHistoryPush && postContextGuard,
  'a superseded request must exit after awaited canonical source capture and before appending its user turn');
assert.ok(taskpaneSource.slice(postContextGuard.start, postContextGuard.end).includes('return;'));
const runPostContextGuard = new Function(
  'currentRequestController', 'requestController',
  `${taskpaneSource.slice(postContextGuard.start, postContextGuard.end)}; return 'continued';`
);
const activeController = {};
assert.equal(runPostContextGuard(activeController, activeController), 'continued');
assert.equal(runPostContextGuard({}, activeController), undefined, 'superseded request returns before history writes');
assert.equal(runPostContextGuard(null, activeController), undefined, 'refresh-cleared controller also exits before history writes');

const onEventProperty = loopInvocation.arguments[0].properties.find(property => property.key?.name === 'onEvent');
assert.equal(onEventProperty?.value?.type, 'FunctionExpression', 'the loop UI callback must stay request-scoped');
// `scope` carries the request-local state the callback closes over across events.
function invokeEvent(currentController, requestController, event, requestRecords = [], scope = {}) {
  const effects = [];
  const callback = new Function(
    'scope', 'currentRequestController', 'requestController', 'requestRecords', 'loadingMsg', 'removeMessage', 'addMessageToChat',
    'updateSystemMessage', 'document', 'event',
    `let shownToolResult = scope.shownToolResult ?? null;
    try {
      ${taskpaneSource.slice(onEventProperty.value.body.start + 1, onEventProperty.value.body.end - 1)}
    } finally { scope.shownToolResult = shownToolResult; }`
  );
  callback(scope, currentController, requestController, requestRecords, 'loading', value => effects.push(['remove', value]),
    (...args) => effects.push(['add', ...args]),
    (...args) => effects.push(['update', ...args]),
    { getElementById: id => ({ appendChild: node => effects.push(['append', id, node]) }) }, event);
  return effects;
}
assert.deepEqual(invokeEvent(activeController, activeController, { type: 'response', text: 'Fresh response.' }), [
  ['remove', 'loading'], ['add', 'Gemini', 'Fresh response.']
], 'the active request may update its response UI');
assert.deepEqual(invokeEvent(activeController, activeController, { type: 'response', text: 'Inspection complete.' }, [
  { name: 'inspect_document' }
]), [
  ['remove', 'loading'], ['add', 'Gemini', 'Inspection complete.']
], 'read-only inspection does not leave a completed progress banner behind');
assert.deepEqual(invokeEvent(activeController, activeController, { type: 'response', text: 'Edit complete.' }, [
  { name: 'apply_document_edits' }
]), [
  ['add', 'Gemini', 'Edit complete.']
], 'an edit result remains visible after a direct apply tool call');
assert.deepEqual(invokeEvent(activeController, activeController, {
  type: 'tool_end', name: 'inspect_document', display: { show: false }, completionText: 'Inspection complete.'
}, [{ name: 'inspect_document' }]), [
  ['append', 'chat-messages', 'loading']
], 'successful inspection does not replace the progress text with a result banner');
assert.deepEqual(invokeEvent(activeController, activeController, {
  type: 'tool_end', name: 'apply_document_edits', display: { show: true }, completionText: 'Applied and verified.', completionWithCheckpoint: false
}, [{ name: 'apply_document_edits' }]), [
  ['update', 'loading', 'Applied and verified.', undefined], ['append', 'chat-messages', 'loading']
], 'a verified edit result is surfaced through the existing progress UI');
{
  // An inspection after an edit shows its progress, then the edit's receipt comes back.
  const scope = {};
  const records = [{ name: 'apply_document_edits' }, { name: 'inspect_document' }];
  invokeEvent(activeController, activeController, {
    type: 'tool_end', name: 'apply_document_edits', display: { show: true }, completionText: 'Applied and verified 1 requested edit(s).',
    completionWithCheckpoint: true, checkpointIndex: 4
  }, records, scope);
  assert.deepEqual(invokeEvent(activeController, activeController, {
    type: 'tool_end', name: 'inspect_document', display: { show: false }, completionText: 'Inspection complete.'
  }, records, scope), [
    ['update', 'loading', 'Applied and verified 1 requested edit(s).'], ['append', 'chat-messages', 'loading']
  ], 'a later read-only step cannot leave its progress text in place of the edit receipt');
}
assert.deepEqual(invokeEvent({}, activeController, { type: 'response', text: 'Fresh response.' }), [],
  'a superseded request cannot write response UI');
assert.deepEqual(invokeEvent(null, activeController, { type: 'response', text: 'Fresh response.' }), [],
  'a cleared request cannot write response UI');

const historySyncGuard = identityConditionals.find(node =>
  isControllerComparison(node.test, '===')
  && containsNode(node.consequent, child => child.type === 'AssignmentExpression'
    && child.left?.type === 'Identifier' && child.left.name === 'chatHistory'
    && child.right?.type === 'MemberExpression' && child.right.object?.name === 'loopState'
    && child.right.property?.name === 'history')
);
assert.ok(historySyncGuard, 'only the active request may synchronize loop history back to the taskpane');
function syncHistory(currentController, requestController, initialHistory, requestHistory) {
  const run = new Function(
    'currentRequestController', 'requestController', 'chatHistory', 'loopState',
    `${taskpaneSource.slice(historySyncGuard.start, historySyncGuard.end)}; return chatHistory;`
  );
  return run(currentController, requestController, initialHistory, { history: requestHistory });
}
const originalHistory = [{ role: 'user', parts: [{ text: 'old conversation' }] }];
const requestHistory = [{ role: 'user', parts: [{ text: 'request-local conversation' }] }];
assert.equal(syncHistory(activeController, activeController, originalHistory, requestHistory), requestHistory);
assert.equal(syncHistory({}, activeController, originalHistory, requestHistory), originalHistory,
  'a superseded request cannot restore history after refresh or overwrite its successor');
assert.equal(syncHistory(null, activeController, originalHistory, requestHistory), originalHistory,
  'a refresh-cleared conversation stays cleared');

const lifecycleGuard = identityConditionals.find(node =>
  isControllerComparison(node.test, '===')
  && containsNode(node.consequent, child => child.type === 'AssignmentExpression'
    && child.left?.type === 'Identifier' && child.left.name === 'currentRequestController'
    && child.right?.type === 'Literal' && child.right.value === null)
);
assert.ok(lifecycleGuard, 'only the active request may clear the controller and unlock the UI');
function cleanUpRequest(currentController, requestController) {
  let currentRequestController = currentController;
  const effects = [];
  const chatInput = { disabled: true, focus: () => effects.push('focus') };
  const sendButton = { disabled: true };
  const thinkButton = { disabled: true };
  const run = new Function(
    'currentRequestController', 'requestController', 'chatInput', 'sendButton', 'thinkButton',
    `${taskpaneSource.slice(lifecycleGuard.start, lifecycleGuard.end)}; return { currentRequestController, chatInput, sendButton, thinkButton };`
  );
  const result = run(currentRequestController, requestController, chatInput, sendButton, thinkButton);
  return { ...result, effects };
}
const activeCleanup = cleanUpRequest(activeController, activeController);
assert.equal(activeCleanup.currentRequestController, null);
assert.equal(activeCleanup.chatInput.disabled, false);
assert.equal(activeCleanup.sendButton.disabled, false);
assert.equal(activeCleanup.thinkButton.disabled, false);
assert.deepEqual(activeCleanup.effects, ['focus']);
for (const staleController of [{}, null]) {
  const staleCleanup = cleanUpRequest(staleController, activeController);
  assert.equal(staleCleanup.currentRequestController, staleController);
  assert.equal(staleCleanup.chatInput.disabled, true);
  assert.equal(staleCleanup.sendButton.disabled, true);
  assert.equal(staleCleanup.thinkButton.disabled, true);
  assert.deepEqual(staleCleanup.effects, []);
}

// Evaluate the actual checkpoint UI adapter with host/storage dependencies
// stubbed. A failed lookup or unverified write must retain both the model
// history and visible transcript; only the current, verified restore may reset
// them and remove stale retry/revert controls with the old message nodes.
let restoreCheckpointNode = null;
visit(taskpaneAst, node => {
  if (node.type === 'FunctionDeclaration' && node.id?.name === 'restoreCheckpoint') restoreCheckpointNode = node;
});
assert.ok(restoreCheckpointNode, 'checkpoint restore must remain an inspectable taskpane adapter');
const restoreFunctionSource = taskpaneSource.slice(restoreCheckpointNode.start, restoreCheckpointNode.end)
  .replace("await import('./modules/chat/checkpoint-restoration.js')", 'await loadCheckpointRestoration()');
assert.notEqual(restoreFunctionSource, taskpaneSource.slice(restoreCheckpointNode.start, restoreCheckpointNode.end),
  'the restore adapter keeps its production import seam for this isolated test');
const createRestoreHarness = new Function(
  'getCheckpoint', 'loadCheckpointRestoration', 'getDocumentId',
  'document', 'addMessageToChat', 'updateSystemMessage', 'console', 'initialController', 'initialHistory',
  `${restoreFunctionSource};
   let currentRequestController = initialController;
   let chatHistory = initialHistory;
   return {
     restoreCheckpoint,
     getCurrent: () => currentRequestController,
     setCurrent: value => { currentRequestController = value; },
     getHistory: () => chatHistory
   };`
);
function makeRestoreHarness({ record = { ooxml: '<checkpoint/>' }, restoreResult, deferredRestore,
  checkpointError, restoreError } = {}) {
  const oldHistory = [{ role: 'user', parts: [{ text: 'old conversation' }] }];
  const messages = { innerHTML: 'old transcript including revert and retry controls' };
  const controls = new Map(['chat-input', 'send-button', 'think-button'].map(id => [id, { disabled: false }]));
  const effects = [];
  const restore = createRestoreHarness(
    async () => {
      if (checkpointError) throw checkpointError;
      return record;
    },
    async () => ({ restoreBodyCheckpoint: async (...args) => {
      effects.push(['restore', ...args]);
      if (restoreError) throw restoreError;
      if (deferredRestore) {
        deferredRestore.started();
        return deferredRestore.promise;
      }
      return restoreResult;
    } }),
    () => 'document-id',
    { getElementById: id => id === 'chat-messages' ? messages : controls.get(id) || null },
    (...args) => { effects.push(['add', ...args]); return { id: `message-${effects.length}` }; },
    (...args) => effects.push(['update', ...args]),
    { error() {} },
    null,
    oldHistory
  );
  return { ...restore, oldHistory, messages, controls, effects };
}

let restoreHarness = makeRestoreHarness({ restoreResult: {
  success: false, verified: false, message: 'The checkpoint could not be verified.'
} });
await restoreHarness.restoreCheckpoint('checkpoint-id');
assert.equal(restoreHarness.getHistory(), restoreHarness.oldHistory,
  'an unverified restore preserves model history');
assert.equal(restoreHarness.messages.innerHTML, 'old transcript including revert and retry controls',
  'an unverified restore preserves the visible transcript and its controls');
assert.ok(restoreHarness.effects.some(effect => effect[0] === 'update' && effect[2] === 'The checkpoint could not be verified.'));

restoreHarness = makeRestoreHarness({ record: null, restoreResult: null });
await restoreHarness.restoreCheckpoint('missing-checkpoint');
assert.equal(restoreHarness.getHistory(), restoreHarness.oldHistory,
  'a missing checkpoint preserves model history');
assert.equal(restoreHarness.messages.innerHTML, 'old transcript including revert and retry controls',
  'a missing checkpoint leaves the existing transcript intact');

restoreHarness = makeRestoreHarness({ checkpointError: new Error('Checkpoint storage failed.') });
await restoreHarness.restoreCheckpoint('checkpoint-id');
assert.equal(restoreHarness.getHistory(), restoreHarness.oldHistory,
  'a checkpoint lookup error preserves model history');
assert.equal(restoreHarness.messages.innerHTML, 'old transcript including revert and retry controls',
  'a checkpoint lookup error preserves the existing transcript');

restoreHarness = makeRestoreHarness({ restoreError: new Error('Word restore failed.') });
await restoreHarness.restoreCheckpoint('checkpoint-id');
assert.equal(restoreHarness.getHistory(), restoreHarness.oldHistory,
  'a restore exception preserves model history');
assert.equal(restoreHarness.messages.innerHTML, 'old transcript including revert and retry controls',
  'a restore exception preserves the existing transcript');

restoreHarness = makeRestoreHarness({ restoreResult: {
  success: true, verified: true, message: 'Reverted and verified the body checkpoint.'
} });
await restoreHarness.restoreCheckpoint('checkpoint-id');
assert.deepEqual(restoreHarness.getHistory(), [], 'a verified restore starts a fresh model conversation');
assert.equal(restoreHarness.messages.innerHTML, '', 'a verified restore removes the old transcript and retry/revert controls');
assert.ok(restoreHarness.effects.some(effect => effect[0] === 'add'
  && effect[1] === 'System' && effect[2] === 'Reverted and verified the body checkpoint.'),
'the visible transcript records the verified restoration');

let signalRestoreResolve, signalRestoreStarted;
const deferredRestore = {
  promise: new Promise(resolve => { signalRestoreResolve = resolve; }),
  started: () => signalRestoreStarted?.()
};
const restoreStarted = new Promise(resolve => { signalRestoreStarted = resolve; });
restoreHarness = makeRestoreHarness({ deferredRestore });
const staleRestore = restoreHarness.restoreCheckpoint('checkpoint-id');
await restoreStarted;
const newerController = {};
restoreHarness.setCurrent(newerController);
signalRestoreResolve({ success: true, verified: true, message: 'Reverted and verified the body checkpoint.' });
await staleRestore;
assert.equal(restoreHarness.getCurrent(), newerController, 'a superseded restore cannot clear its replacement controller');
assert.equal(restoreHarness.getHistory(), restoreHarness.oldHistory,
  'a superseded restore cannot reset the successor conversation history');
assert.equal(restoreHarness.messages.innerHTML, 'old transcript including revert and retry controls',
  'a superseded restore cannot clear the successor transcript');

console.log('PASS: direct tool dispatch and request-scoped UI/history guards');
