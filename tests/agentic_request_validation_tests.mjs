import assert from 'assert';
import { validateListRequest } from '../src/taskpane/modules/commands/agentic-request-validation.js';

function expectInvalid(request, paragraphCount, message) {
  const result = validateListRequest(request, paragraphCount);
  assert.equal(result.valid, false, message);
  assert.equal(result.error.code, 'INVALID_LIST_REQUEST');
  assert.equal(typeof result.error.message, 'string');
}

function testEditListValidationAndPreservation() {
  const items = ['Top level', '    * custom indented child'];
  const result = validateListRequest({
    startParagraphIndex: 2,
    endParagraphIndex: 5,
    newItems: items,
    listType: 'bullet'
  }, 8);

  assert.deepStrictEqual(result, {
    valid: true,
    request: {
      kind: 'edit_list',
      startParagraphIndex: 2,
      endParagraphIndex: 5,
      newItems: ['Top level', '    * custom indented child'],
      listType: 'bullet',
      numberingStyle: 'decimal'
    }
  });
  assert.notStrictEqual(result.request.newItems, items, 'normalized request should own its array');
  assert.deepStrictEqual(items, ['Top level', '    * custom indented child'], 'input item strings must remain untouched');

  const numbered = validateListRequest({
    startParagraphIndex: 1,
    endParagraphIndex: 1,
    newItems: ['Item'],
    listType: 'numbered',
    numberingStyle: 'lowerRoman'
  }, 1);
  assert.equal(numbered.valid, true);
  assert.equal(numbered.request.numberingStyle, 'lowerRoman');
}

function testEditListRejectsInvalidIndexesAndArguments() {
  const validRemainder = { newItems: ['One'], listType: 'numbered' };
  for (const [field, value] of [
    ['startParagraphIndex', '1'],
    ['startParagraphIndex', 0],
    ['startParagraphIndex', 1.5],
    ['endParagraphIndex', -1],
    ['endParagraphIndex', 9]
  ]) {
    const request = { startParagraphIndex: 1, endParagraphIndex: 2, ...validRemainder, [field]: value };
    expectInvalid(request, 8, `${field}=${String(value)} must fail closed`);
  }

  expectInvalid({ startParagraphIndex: 4, endParagraphIndex: 3, ...validRemainder }, 8, 'reversed list range must fail');
  expectInvalid({ startParagraphIndex: 1, endParagraphIndex: 2, newItems: [], listType: 'bullet' }, 8, 'empty list items must fail');
  expectInvalid({ startParagraphIndex: 1, endParagraphIndex: 2, newItems: ['   '], listType: 'bullet' }, 8, 'blank list item must fail');
  expectInvalid({ startParagraphIndex: 1, endParagraphIndex: 2, newItems: ['One', 2], listType: 'bullet' }, 8, 'non-string list item must fail');
  expectInvalid({ startParagraphIndex: 1, endParagraphIndex: 2, newItems: ['One'], listType: 'Bullet' }, 8, 'list type enum is case-sensitive');
  expectInvalid({ startParagraphIndex: 1, endParagraphIndex: 2, newItems: ['One'], listType: 'bullet', numberingStyle: 'roman' }, 8, 'numbering style must match the declared enum');
  expectInvalid({ startParagraphIndex: 1, endParagraphIndex: 2, newItems: ['One'], listType: 'bullet', numberingStyle: '' }, 8, 'an explicitly empty numbering style is invalid');
  expectInvalid({ startParagraphIndex: 1, endParagraphIndex: 2, newItems: ['One'], listType: 'bullet' }, '8', 'paragraph count must be an integer when provided');
  expectInvalid(null, 8, 'null request must fail');
}

function testInsertListItemValidation() {
  const result = validateListRequest({ afterParagraphIndex: 4, text: '  Child item  ' }, 4);
  assert.deepStrictEqual(result, {
    valid: true,
    request: {
      kind: 'insert_list_item',
      afterParagraphIndex: 4,
      text: '  Child item  ',
      indentLevel: 0
    }
  });

  for (const value of [-1, 0, '4', 1.2, null]) {
    expectInvalid({ afterParagraphIndex: value, text: 'Item' }, 4, `afterParagraphIndex=${String(value)} must fail`);
  }
  expectInvalid({ afterParagraphIndex: 5, text: 'Item' }, 4, 'out-of-range insertion anchor must fail');
  expectInvalid({ afterParagraphIndex: 1, text: '  ' }, 4, 'blank inserted text must fail');
  expectInvalid({ afterParagraphIndex: 1, text: 'First\nSecond' }, 4, 'multi-paragraph inserted text must fail');
  expectInvalid({ afterParagraphIndex: 1, text: 'First\u2028Second' }, 4, 'Unicode line separator must fail');
  for (const value of [-2, 2, '1', 0.5, null]) {
    expectInvalid({ afterParagraphIndex: 1, text: 'Item', indentLevel: value }, 4, `indentLevel=${String(value)} must fail`);
  }
  for (const value of [-1, 0, 1]) {
    const valid = validateListRequest({ afterParagraphIndex: 2, text: 'Item', indentLevel: value }, 4);
    assert.equal(valid.valid, true);
    assert.equal(valid.request.indentLevel, value);
  }
}

function testHeaderTextPairsStayBoundWhenSorted() {
  const source = {
    paragraphIndices: [9, 3, 6],
    newHeaderTexts: ['Ninth header', 'Third header', 'Sixth header'],
    numberingFormat: 'upperRoman'
  };
  const result = validateListRequest(source, 10);

  assert.deepStrictEqual(result, {
    valid: true,
    request: {
      kind: 'convert_headers_to_list',
      paragraphIndices: [3, 6, 9],
      newHeaderTexts: ['Third header', 'Sixth header', 'Ninth header'],
      numberingFormat: 'upperRoman',
      headerRecords: [
        { paragraphIndex: 3, text: 'Third header' },
        { paragraphIndex: 6, text: 'Sixth header' },
        { paragraphIndex: 9, text: 'Ninth header' }
      ]
    }
  });
  assert.deepStrictEqual(source.paragraphIndices, [9, 3, 6], 'validation must not sort caller-owned indexes in place');
  assert.deepStrictEqual(source.newHeaderTexts, ['Ninth header', 'Third header', 'Sixth header']);

  const noReplacementText = validateListRequest({ paragraphIndices: [4, 2] }, 5);
  assert.deepStrictEqual(noReplacementText.request.headerRecords, [
    { paragraphIndex: 2 },
    { paragraphIndex: 4 }
  ]);
  assert.deepStrictEqual(noReplacementText.request.paragraphIndices, [2, 4]);
  assert.equal(noReplacementText.request.numberingFormat, 'arabic');
  assert.equal(Object.hasOwn(noReplacementText.request, 'newHeaderTexts'), false);
}

function testHeaderIndexesAndTextConflictsFailClosed() {
  expectInvalid({ paragraphIndices: [] }, 10, 'empty header indexes must fail');
  expectInvalid({ paragraphIndices: ['3'] }, 10, 'numeric string header index must fail');
  expectInvalid({ paragraphIndices: [0] }, 10, 'zero header index must fail');
  expectInvalid({ paragraphIndices: [11] }, 10, 'out-of-range header index must fail');
  expectInvalid({ paragraphIndices: [3, 3] }, 10, 'duplicate header indexes must fail');
  expectInvalid({ paragraphIndices: [3, 7], newHeaderTexts: ['Only one'] }, 10, 'header text count must match index count');
  expectInvalid({ paragraphIndices: [3], newHeaderTexts: [] }, 10, 'provided empty header texts must fail');
  expectInvalid({ paragraphIndices: [3], newHeaderTexts: [''] }, 10, 'empty header text must fail');
  expectInvalid({ paragraphIndices: [3], newHeaderTexts: ['Header', null] }, 10, 'all header texts must be strings');
  expectInvalid({ paragraphIndices: [3], numberingFormat: 'lowercase' }, 10, 'numbering format must match the declared enum');
  expectInvalid({ paragraphIndices: [3], numberingFormat: '' }, 10, 'an explicitly empty numbering format is invalid');
  expectInvalid({ paragraphIndices: [3], startParagraphIndex: 1 }, 10, 'mixed list request kinds must fail');
}

function run() {
  testEditListValidationAndPreservation();
  testEditListRejectsInvalidIndexesAndArguments();
  testInsertListItemValidation();
  testHeaderTextPairsStayBoundWhenSorted();
  testHeaderIndexesAndTextConflictsFailClosed();
}

try {
  run();
  console.log('PASS: agentic request validation tests');
} catch (error) {
  console.error('FAIL:', error?.message || error);
  process.exit(1);
}
