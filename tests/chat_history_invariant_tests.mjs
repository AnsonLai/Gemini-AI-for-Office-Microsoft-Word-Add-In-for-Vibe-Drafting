import assert from 'assert';
import {
  appendFunctionExchange,
  maintainHistoryWindow,
  validateHistoryPairs
} from '../src/taskpane/modules/chat/chat-history.js';

// Helpers to build well-formed turns.
function modelTurn(callNames, extraParts = []) {
  const parts = callNames.map((name) => ({ functionCall: { name, args: {} } }));
  return { role: 'model', parts: [...extraParts, ...parts] };
}
function userTurn(responseNames) {
  return {
    role: 'user',
    parts: responseNames.map((name) => ({
      functionResponse: { name, response: { name, content: [{ text: 'ok' }] } }
    }))
  };
}

function testValidPairAppendsBoth() {
  const history = [{ role: 'user', parts: [{ text: 'hi' }] }];
  const result = appendFunctionExchange(history, modelTurn(['apply_redlines']), userTurn(['apply_redlines']));
  assert.strictEqual(result, history, 'returns the same array');
  assert.strictEqual(history.length, 3);
  assert.strictEqual(history[1].role, 'model');
  assert.strictEqual(history[2].role, 'user');
}

function testMixedModelPartsAllowed() {
  // A model turn may include thought/text parts alongside the functionCall.
  const history = [];
  const mTurn = modelTurn(['edit_list'], [{ text: 'Let me fix that.', thought: true }]);
  appendFunctionExchange(history, mTurn, userTurn(['edit_list']));
  assert.strictEqual(history.length, 2);
}

function testMultipleToolsMatched() {
  const history = [];
  appendFunctionExchange(
    history,
    modelTurn(['apply_redlines', 'insert_comment']),
    userTurn(['apply_redlines', 'insert_comment'])
  );
  assert.strictEqual(history.length, 2);
}

function testMismatchedNameThrowsAndLeavesHistoryUnchanged() {
  const history = [{ role: 'user', parts: [{ text: 'hi' }] }];
  assert.throws(
    () => appendFunctionExchange(history, modelTurn(['apply_redlines']), userTurn(['insert_comment'])),
    /mismatch|uncalled/
  );
  assert.strictEqual(history.length, 1, 'history must be untouched on throw');
}

function testCountMismatchThrows() {
  const history = [];
  // 2 calls to the same tool, only 1 response.
  assert.throws(
    () => appendFunctionExchange(history, modelTurn(['apply_redlines', 'apply_redlines']), userTurn(['apply_redlines'])),
    /mismatch/
  );
  assert.strictEqual(history.length, 0);
}

function testNoFunctionCallThrows() {
  const history = [];
  assert.throws(
    () => appendFunctionExchange(history, { role: 'model', parts: [{ text: 'hello' }] }, userTurn([])),
    /no functionCall/
  );
  assert.strictEqual(history.length, 0);
}

function testBadShapesThrow() {
  assert.throws(() => appendFunctionExchange(null, modelTurn(['x']), userTurn(['x'])), /history must be an array/);
  assert.throws(() => appendFunctionExchange([], { role: 'user', parts: [] }, userTurn([])), /modelTurn must be/);
  assert.throws(() => appendFunctionExchange([], modelTurn(['x']), { role: 'model', parts: [] }), /userTurn must be/);
}

function testValidateHistoryPairsLeavesBuiltHistoryUnchanged() {
  // A history assembled solely via appendFunctionExchange must survive validation.
  let history = [{ role: 'user', parts: [{ text: 'do two things' }] }];
  appendFunctionExchange(history, modelTurn(['apply_redlines']), userTurn(['apply_redlines']));
  appendFunctionExchange(history, modelTurn(['insert_comment']), userTurn(['insert_comment']));
  const validated = validateHistoryPairs(history);
  assert.deepStrictEqual(validated, history, 'validation should not drop any turns');
}

function buildLongToolHistory() {
  const history = [{ role: 'user', parts: [{ text: 'Please make the requested edits.' }] }];
  for (let i = 0; i < 7; i++) {
    const name = `tool_${i}`;
    appendFunctionExchange(history, modelTurn([name]), userTurn([name]));
  }
  return history;
}

function functionCallNames(history) {
  return history.flatMap((turn) => (turn.parts || [])
    .filter((part) => part.functionCall)
    .map((part) => part.functionCall.name));
}

function testHistoryWindowKeepsRequestAndRecentPairsAfterLongToolTurn() {
  const history = buildLongToolHistory();
  const originalHistory = structuredClone(history);

  const windowed = maintainHistoryWindow(history, 10);

  assert.ok(windowed.length <= 10, 'window must remain within the configured message cap');
  assert.deepStrictEqual(windowed[0], history[0], 'the first retained turn must be the actual user request');
  assert.deepStrictEqual(functionCallNames(windowed), ['tool_3', 'tool_4', 'tool_5', 'tool_6']);
  assert.deepStrictEqual(validateHistoryPairs(windowed), windowed, 'retained exchanges must remain valid');
  assert.deepStrictEqual(history, originalHistory, 'windowing must not mutate the input history');
}

function testHistoryWindowKeepsNextRequestAndRecentPairs() {
  const history = buildLongToolHistory();
  const nextRequest = { role: 'user', parts: [{ text: 'Now summarize the changes.' }] };
  history.push(nextRequest);
  const originalHistory = structuredClone(history);

  const windowed = maintainHistoryWindow(history, 10);

  assert.ok(windowed.length <= 10, 'window must remain within the configured message cap');
  assert.deepStrictEqual(windowed[0], history[0], 'tool exchanges need their original user context');
  assert.deepStrictEqual(windowed[windowed.length - 1], nextRequest, 'the newest user request must survive');
  assert.deepStrictEqual(functionCallNames(windowed), ['tool_3', 'tool_4', 'tool_5', 'tool_6']);
  assert.deepStrictEqual(validateHistoryPairs(windowed), windowed, 'retained exchanges must remain valid');
  assert.deepStrictEqual(history, originalHistory, 'windowing must not mutate the input history');
}

function testSmallHistoryWindowPrioritizesRequestAndWholePairs() {
  const history = buildLongToolHistory();

  const windowed = maintainHistoryWindow(history, 4);

  assert.ok(windowed.length <= 4);
  assert.deepStrictEqual(windowed[0], history[0], 'the actual request takes priority');
  assert.deepStrictEqual(functionCallNames(windowed), ['tool_6'], 'only a complete recent pair should be retained');
  assert.deepStrictEqual(validateHistoryPairs(windowed), windowed);
}

function testHistoryWindowKeepsEachSelectedPairWithItsUserContext() {
  const history = [{ role: 'user', parts: [{ text: 'Request 1' }] }];
  for (let i = 1; i <= 4; i++) {
    const name = `context_tool_${i}`;
    appendFunctionExchange(history, modelTurn([name]), userTurn([name]));
    history.push({ role: 'user', parts: [{ text: `Request ${i + 1}` }] });
  }
  const originalHistory = structuredClone(history);

  const windowed = maintainHistoryWindow(history, 10);
  const contextByTool = {};
  for (let i = 0; i < windowed.length; i++) {
    const call = (windowed[i].parts || []).find((part) => part.functionCall);
    if (!call) continue;
    const context = windowed.slice(0, i).reverse().find((turn) =>
      turn.role === 'user' &&
      !(turn.parts || []).some((part) => part.functionResponse) &&
      (turn.parts || []).some((part) => typeof part.text === 'string' && part.text.trim())
    );
    contextByTool[call.functionCall.name] = context.parts.find((part) => part.text).text;
  }

  assert.ok(windowed.length <= 10);
  assert.deepStrictEqual(functionCallNames(windowed), ['context_tool_2', 'context_tool_3', 'context_tool_4']);
  assert.deepStrictEqual(contextByTool, {
    context_tool_2: 'Request 2',
    context_tool_3: 'Request 3',
    context_tool_4: 'Request 4'
  }, 'each retained tool exchange must keep the real request that preceded it');
  assert.deepStrictEqual(windowed[windowed.length - 1], history[history.length - 1], 'latest request must survive');
  assert.deepStrictEqual(validateHistoryPairs(windowed), windowed);
  assert.deepStrictEqual(history, originalHistory, 'windowing must not mutate the input history');
}

testValidPairAppendsBoth();
testMixedModelPartsAllowed();
testMultipleToolsMatched();
testMismatchedNameThrowsAndLeavesHistoryUnchanged();
testCountMismatchThrows();
testNoFunctionCallThrows();
testBadShapesThrow();
testValidateHistoryPairsLeavesBuiltHistoryUnchanged();
testHistoryWindowKeepsRequestAndRecentPairsAfterLongToolTurn();
testHistoryWindowKeepsNextRequestAndRecentPairs();
testSmallHistoryWindowPrioritizesRequestAndWholePairs();
testHistoryWindowKeepsEachSelectedPairWithItsUserContext();

console.log('chat_history_invariant_tests passed');
