import assert from 'node:assert/strict';
import { recoverFunctionCall } from '../src/taskpane/modules/chat/function-call-recovery.js';

function testTolerantObjectParsing() {
  const result = recoverFunctionCall(
    "insert_comment({paragraphIndex: 4, text: 'Please review this provision',})",
    ['insert_comment']
  );

  assert.deepEqual(result, {
    name: 'insert_comment',
    args: { paragraphIndex: 4, text: 'Please review this provision' }
  });
}

function testMalformedEditListRecovery() {
  const result = recoverFunctionCall(
    'edit_list({startParagraphIndex: 2, endParagraphIndex: 4, newItems: [Protect confidential information. , Notify the other party promptly.], listType: bullet, numberingStyle: lowerLetter})',
    ['edit_list']
  );

  assert.deepEqual(result, {
    name: 'edit_list',
    args: {
      startParagraphIndex: 2,
      endParagraphIndex: 4,
      listType: 'bullet',
      numberingStyle: 'lowerLetter',
      newItems: [
        'Protect confidential information.',
        'Notify the other party promptly.'
      ]
    }
  });
}

function testMalformedHeaderConversionRecovery() {
  const result = recoverFunctionCall(
    'convert_headers_to_list({paragraphIndices: [1, 3, malformed, 5], numberingFormat: upperRoman})',
    ['convert_headers_to_list']
  );

  assert.deepEqual(result, {
    name: 'convert_headers_to_list',
    args: { paragraphIndices: [1, 3, 5], numberingFormat: 'upperRoman' }
  });
}

function testOnlyKnownToolNamesAreScanned() {
  assert.equal(
    recoverFunctionCall('unknown_tool({instruction: "do something"})', ['apply_redlines']),
    null
  );
  assert.equal(recoverFunctionCall('apply_redlines({instruction: "fix this"})', []), null);
}

function testLegacyRedlineFallbackRequiresAuthorization() {
  const finishMessage = 'apply_redlines { instruction: Replace the misspelled clause with the approved wording. }';

  assert.equal(recoverFunctionCall(finishMessage, ['insert_comment']), null);
  assert.deepEqual(recoverFunctionCall(finishMessage, ['apply_redlines']), {
    name: 'apply_redlines',
    args: { instruction: 'Replace the misspelled clause with the approved wording.' }
  });
}

function testValidCallAndUnparseableInput() {
  assert.deepEqual(
    recoverFunctionCall('apply_redlines({"instruction":"correct the typo"})', new Set(['apply_redlines'])),
    { name: 'apply_redlines', args: { instruction: 'correct the typo' } }
  );

  assert.equal(recoverFunctionCall(undefined, ['apply_redlines']), null);
  assert.equal(recoverFunctionCall('edit_list({startParagraphIndex: nope})', ['edit_list']), null);
}

testTolerantObjectParsing();
testMalformedEditListRecovery();
testMalformedHeaderConversionRecovery();
testOnlyKnownToolNamesAreScanned();
testLegacyRedlineFallbackRequiresAuthorization();
testValidCallAndUnparseableInput();
console.log('PASS: function call recovery');
