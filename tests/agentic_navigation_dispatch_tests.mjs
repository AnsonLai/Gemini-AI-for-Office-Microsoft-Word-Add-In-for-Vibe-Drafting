import assert from 'node:assert/strict';
import fs from 'node:fs';
import { parse } from 'acorn';

const taskpanePath = new URL('../src/taskpane/taskpane.js', import.meta.url);
const source = fs.readFileSync(taskpanePath, 'utf8');
const ast = parse(source, { ecmaVersion: 'latest', sourceType: 'module' });

function visit(node, callback) {
  if (!node || typeof node !== 'object') return;
  callback(node);
  for (const value of Object.values(node)) {
    if (Array.isArray(value)) value.forEach(child => visit(child, callback));
    else if (value && typeof value === 'object') visit(value, callback);
  }
}

let navigationBranch = null;
let responsePush = null;
let sendChatMessage = null;
visit(ast, node => {
  if (node.type === 'FunctionDeclaration' && node.id?.name === 'sendChatMessage') {
    sendChatMessage = node;
  }

  if (node.type === 'IfStatement'
    && node.test?.type === 'BinaryExpression'
    && node.test.operator === '==='
    && node.test.left?.type === 'MemberExpression'
    && node.test.left.object?.name === 'functionCall'
    && node.test.left.property?.name === 'name'
    && node.test.right?.value === 'navigate_to_section') {
    navigationBranch = node.consequent;
  }

  if (node.type === 'CallExpression'
    && node.callee?.type === 'MemberExpression'
    && node.callee.object?.name === 'functionResponses'
    && node.callee.property?.name === 'push'
    && source.slice(node.start, node.end).includes('functionResponse:')) {
    responsePush = node;
  }
});

assert.ok(navigationBranch?.type === 'BlockStatement', 'navigation dispatch branch must exist');
assert.ok(responsePush, 'shared Gemini function response builder must exist');
assert.ok(sendChatMessage, 'sendChatMessage must remain an inspectable function');

const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
const runNavigationBranch = new AsyncFunction(
  'functionCall', 'instruction', 'loadingMsg', 'docText', 'agenticTools',
  'updateSystemMessage', 'toolsExecutedInCurrentRequest',
  `let toolResult = '';\nlet toolSucceeded = false;\n${source.slice(navigationBranch.start + 1, navigationBranch.end - 1)}\nreturn { toolResult, toolSucceeded, toolsExecutedInCurrentRequest };`
);
const emitFunctionResponse = new Function(
  'functionResponses', 'functionCall', 'toolResult',
  `${source.slice(responsePush.start, responsePush.end)};\nreturn functionResponses;`
);

async function verifyNavigationResponse(executeNavigate, expectedSuccess, expectedMessage) {
  const updates = [];
  const toolRecords = [];
  const functionCall = { name: 'navigate_to_section' };
  const result = await runNavigationBranch(
    functionCall,
    'go to the requested section',
    {},
    '[P1] Section heading',
    { executeNavigate },
    (_loading, message) => updates.push(message),
    toolRecords
  );

  assert.equal(result.toolSucceeded, expectedSuccess);
  assert.equal(result.toolResult, expectedMessage);
  assert.equal(toolRecords[0].result.message, expectedMessage, 'recovery record keeps the structured tool result');

  const functionResponses = emitFunctionResponse([], functionCall, result.toolResult);
  const responseText = functionResponses[0].functionResponse.response.content[0].text;
  assert.equal(typeof responseText, 'string', 'Gemini function response text must remain a string');
  assert.equal(responseText, expectedMessage);
}

await verifyNavigationResponse(
  async () => ({ status: 'ok', success: true, message: 'Selected the target section.' }),
  true,
  'Selected the target section.'
);

await verifyNavigationResponse(
  async () => ({
    status: 'error',
    success: false,
    error: { code: 'MISSING_API_KEY', message: 'Error: Please set your Gemini API key in the Settings.' },
    message: 'Error: Please set your Gemini API key in the Settings.'
  }),
  false,
  'Error: Please set your Gemini API key in the Settings.'
);

let baselineDeclaration = null;
let baselineAssignment = null;
let redlineDispatch = null;
visit(sendChatMessage.body, node => {
  if (node.type === 'VariableDeclarator' && node.id?.name === 'docSourceBaseline') {
    baselineDeclaration = node;
  }
  if (node.type === 'AssignmentExpression'
    && node.left?.type === 'Identifier'
    && node.left.name === 'docSourceBaseline'
    && node.right?.type === 'CallExpression'
    && node.right.callee?.type === 'MemberExpression'
    && node.right.callee.object?.name === 'wordOperationRunner'
    && node.right.callee.property?.name === 'captureWordSourceBaseline') {
    baselineAssignment = node;
  }
  if (node.type === 'CallExpression'
    && node.callee?.type === 'MemberExpression'
    && node.callee.object?.name === 'agenticTools'
    && node.callee.property?.name === 'executeRedline'
    && node.arguments[2]?.type === 'Identifier'
    && node.arguments[2].name === 'docSourceBaseline') {
    redlineDispatch = node;
  }
});

assert.ok(baselineDeclaration, 'sendChatMessage must declare the baseline locally');
assert.ok(baselineAssignment, 'sendChatMessage must capture the canonical source baseline');
assert.ok(redlineDispatch, 'apply_redlines must pass the captured baseline to executeRedline');
assert.ok(
  baselineDeclaration.start < baselineAssignment.start
    && baselineAssignment.end < sendChatMessage.body.end
    && baselineDeclaration.start < redlineDispatch.start
    && redlineDispatch.end < sendChatMessage.body.end,
  'declaration, capture, and redline dispatch must share the sendChatMessage lexical scope'
);

const baselineSource = [{ index: 1, exactText: 'Source paragraph', fingerprint: 'fnv1a32:test' }];
const captureAndStoreBaseline = new Function(
  'wordOperationRunner', 'sourceOoxml',
  `let docSourceBaseline = [];\n${source.slice(baselineAssignment.start, baselineAssignment.end)};\nreturn docSourceBaseline;`
);
const capturedBaseline = captureAndStoreBaseline(
  {
    captureWordSourceBaseline: ooxml => {
      assert.equal(ooxml, '<pkg:package/>');
      return baselineSource;
    }
  },
  { value: '<pkg:package/>' }
);
let executeRedlineArguments = null;
const invokeRedlineDispatch = new Function(
  'agenticTools', 'instruction', 'docText', 'docSourceBaseline',
  `return ${source.slice(redlineDispatch.start, redlineDispatch.end)};`
);
await invokeRedlineDispatch(
  { executeRedline: async (...args) => { executeRedlineArguments = args; } },
  'revise the source paragraph',
  '[P1] Source paragraph',
  capturedBaseline
);
assert.equal(executeRedlineArguments[2], baselineSource,
  'the baseline produced by capture must reach the actual executeRedline call');

console.log('PASS: agentic navigation dispatch tests');
