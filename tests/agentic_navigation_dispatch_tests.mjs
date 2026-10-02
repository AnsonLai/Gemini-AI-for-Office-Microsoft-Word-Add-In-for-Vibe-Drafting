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

async function runNavigationRoundTrip(rawResult, expectedSuccess, expectedResponseText) {
  const registry = createToolRegistry({
    executeNavigate: async (instruction, documentText) => {
      assert.equal(instruction, 'go to the requested section');
      assert.equal(documentText, '[P1] Section heading');
      return rawResult;
    }
  });
  const state = {
    history: [{ role: 'user', parts: [{ text: 'Go to the requested section.' }] }],
    docText: '[P1] Section heading',
    docSourceBaseline: [],
    records: []
  };
  const replies = [
    { candidates: [{ content: { parts: [{ functionCall: {
      name: 'navigate_to_section', args: { instruction: 'go to the requested section' }
    } }] } }] },
    { candidates: [{ content: { parts: [{ text: 'Finished.' }] } }] }
  ];

  await runDocumentAgentLoop({
    state,
    registry,
    loadRegistry: async () => registry,
    requestModel: async () => replies.shift(),
    refreshContext: async () => assert.fail('navigation must not refresh document context'),
    createCheckpoint: async () => assert.fail('navigation must not create an edit checkpoint'),
    requestStartTime: 0,
    now: () => 1
  });

  assert.equal(state.records[0].success, expectedSuccess);
  assert.deepEqual(state.records[0].result, rawResult, 'the recovery record retains the structured navigation result');

  const responseTurn = state.history.find(turn => turn.role === 'user'
    && turn.parts.some(part => part.functionResponse?.name === 'navigate_to_section'));
  assert.ok(responseTurn, 'the model tool call must have a paired function response');
  const responseText = responseTurn.parts.find(part => part.functionResponse)
    .functionResponse.response.content[0].text;
  assert.equal(typeof responseText, 'string', 'Gemini function response text must be a string');
  assert.equal(responseText, expectedResponseText);
}

await runNavigationRoundTrip(
  { status: 'ok', success: true, message: 'Selected the target section.' },
  true,
  'Selected the target section.'
);
await runNavigationRoundTrip(
  {
    status: 'error',
    success: false,
    error: { code: 'MISSING_API_KEY', message: 'Error: Please set your Gemini API key in the Settings.' },
    message: 'Error: Please set your Gemini API key in the Settings.'
  },
  false,
  'Error: Please set your Gemini API key in the Settings.'
);

// The exact canonical source baseline is derived from the same library
// inspection that formats model context, then handed to the loop before any
// registry executor runs.
let baselineDeclaration = null;
let baselineAssignment = null;
let sourceBaselineFromDeclaration = null;
let loopStateDeclaration = null;
let loopInvocation = null;
visit(sendChatMessage.body, node => {
  if (node.type === 'VariableDeclarator' && node.id?.name === 'docSourceBaseline') {
    baselineDeclaration = node;
    if (node.init?.type === 'CallExpression'
      && node.init.callee?.type === 'Identifier'
      && node.init.callee.name === 'sourceBaselineFrom'
      && node.init.arguments[0]?.type === 'MemberExpression'
      && node.init.arguments[0].object?.name === 'documentSource'
      && node.init.arguments[0].property?.name === 'inspection') {
      baselineAssignment = node;
    }
  }
  if (node.type === 'VariableDeclarator' && node.id?.name === 'loopState') loopStateDeclaration = node;
  if (node.type === 'CallExpression' && node.callee?.type === 'Identifier'
    && node.callee.name === 'runDocumentAgentLoop') loopInvocation = node;
});
visit(taskpaneAst, node => {
  if (node.type === 'VariableDeclarator' && node.id?.name === 'sourceBaselineFrom') {
    sourceBaselineFromDeclaration = node;
  }
});

assert.ok(baselineDeclaration, 'sendChatMessage must declare the source baseline locally');
assert.ok(baselineAssignment, 'sendChatMessage must derive the baseline from documentSource.inspection');
assert.ok(sourceBaselineFromDeclaration?.init, 'the taskpane must define its baseline projection');
assert.ok(loopStateDeclaration?.init?.type === 'ObjectExpression', 'sendChatMessage must construct loop state');
assert.ok(loopInvocation, 'sendChatMessage must hand work to the portable loop');
const baselineStateProperty = loopStateDeclaration.init.properties.find(property =>
  property.key?.name === 'docSourceBaseline'
  && property.value?.type === 'Identifier'
  && property.value.name === 'docSourceBaseline'
);
const loopStateArgument = loopInvocation.arguments[0]?.properties?.find(property => property.key?.name === 'state');
assert.ok(baselineStateProperty, 'loop state must contain the locally captured docSourceBaseline');
assert.equal(loopStateArgument?.value?.name, 'loopState', 'the captured baseline state must be passed into the loop');
assert.ok(
  baselineAssignment.start < loopInvocation.start
    && loopInvocation.end < sendChatMessage.body.end,
  'baseline derivation and loop invocation must share sendChatMessage scope and order'
);

const baselineSource = [{ index: 1, exactText: 'Source paragraph', fingerprint: 'fnv1a32:test', paragraphId: 'ABCDEF01', inTable: false }];
const sourceBaselineFrom = new Function(
  `return (${taskpaneSource.slice(sourceBaselineFromDeclaration.init.start, sourceBaselineFromDeclaration.init.end)});`
)();
const deriveBaseline = new Function(
  'sourceBaselineFrom', 'documentSource',
  `return ${taskpaneSource.slice(baselineAssignment.init.start, baselineAssignment.init.end)};`
);
const capturedBaseline = deriveBaseline(sourceBaselineFrom, { inspection: { paragraphs: baselineSource } });
assert.deepEqual(capturedBaseline, baselineSource,
  'the actual taskpane sourceBaselineFrom projection must preserve canonical targeting fields');

const makeLoopState = new Function(
  'chatHistory', 'docText', 'docSourceBaseline', 'requestRecords',
  `return (${taskpaneSource.slice(loopStateDeclaration.init.start, loopStateDeclaration.init.end)});`
);
const state = makeLoopState(
  [{ role: 'user', parts: [{ text: 'Revise the source paragraph.' }] }],
  '[P1] Source paragraph',
  capturedBaseline,
  []
);
assert.deepEqual(state.docSourceBaseline, capturedBaseline);

let executeRedlineArguments = null;
const redlineRegistry = createToolRegistry({
  executeRedline: async (...args) => {
    executeRedlineArguments = args;
    return {
      status: 'ok', success: true, showToUser: true, message: 'Edit applied.',
      written: true, writeAttempted: true, mutationOutcome: 'applied'
    };
  }
});
const replies = [
  { candidates: [{ content: { parts: [{ functionCall: {
    name: 'apply_redlines', args: { instruction: 'revise the source paragraph' }
  } }] } }] },
  { candidates: [{ content: { parts: [{ text: 'Done.' }] } }] }
];
await runDocumentAgentLoop({
  state,
  registry: redlineRegistry,
  loadRegistry: async () => redlineRegistry,
  requestModel: async () => replies.shift(),
  refreshContext: async () => ({ formattedText: '[P1] Updated paragraph', sourceBaseline: [] }),
  createCheckpoint: async () => 'checkpoint',
  requestStartTime: 0,
  now: () => 1
});
assert.deepEqual(executeRedlineArguments[2], capturedBaseline,
  'the exact baseline projected from canonical inspection must reach executeRedline through loop state and registry');

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
function invokeResponseEvent(currentController, requestController) {
  const effects = [];
  const callback = new Function(
    'currentRequestController', 'requestController', 'requestRecords', 'loadingMsg', 'removeMessage', 'addMessageToChat', 'event',
    `${taskpaneSource.slice(onEventProperty.value.body.start + 1, onEventProperty.value.body.end - 1)}`
  );
  callback(currentController, requestController, [], 'loading', value => effects.push(['remove', value]),
    (...args) => effects.push(['add', ...args]), { type: 'response', text: 'Fresh response.' });
  return effects;
}
assert.deepEqual(invokeResponseEvent(activeController, activeController), [
  ['remove', 'loading'], ['add', 'Gemini', 'Fresh response.']
], 'the active request may update its response UI');
assert.deepEqual(invokeResponseEvent({}, activeController), [], 'a superseded request cannot write response UI');
assert.deepEqual(invokeResponseEvent(null, activeController), [], 'a cleared request cannot write response UI');

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

console.log('PASS: agentic navigation dispatch and loop-bound source baseline');
