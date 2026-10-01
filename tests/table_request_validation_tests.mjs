import assert from 'node:assert/strict';
import { validateTableRequest } from '../src/taskpane/modules/commands/table-request-validation.js';

function expectInvalid(request, live, explanation) {
  const result = validateTableRequest(request, live);
  assert.equal(result.valid, false, explanation);
  assert.equal(result.error.code, 'INVALID_TABLE_REQUEST');
  assert.equal(typeof result.error.message, 'string');
}

function testReplaceContentNormalizesAndChecksPerRowBounds() {
  const source = [['Header', ''], ['A', 'B']];
  const result = validateTableRequest({ paragraphIndex: 3, action: 'replace_content', content: source }, {
    paragraphCount: 8,
    rowCount: 3,
    columnCount: 2,
    rowCellCounts: [2, 2, 2]
  });

  assert.deepEqual(result, {
    valid: true,
    request: {
      kind: 'edit_table',
      paragraphIndex: 3,
      action: 'replace_content',
      content: [['Header', ''], ['A', 'B']]
    }
  });
  assert.notEqual(result.request.content, source, 'normalized rows must be independent arrays');
  assert.deepEqual(source, [['Header', ''], ['A', 'B']], 'caller data must not be mutated');

  // The executor overlays the supplied cells; a shorter/ragged matrix is valid.
  assert.equal(validateTableRequest({ paragraphIndex: 1, action: 'replace_content', content: [['x'], ['y', 'z']] }, {
    rowCount: 2,
    rowCellCounts: [2, 2]
  }).valid, true);

  expectInvalid({ paragraphIndex: 1, action: 'replace_content', content: [] }, {}, 'empty replacement matrix must fail');
  expectInvalid({ paragraphIndex: 1, action: 'replace_content', content: [[]] }, {}, 'empty replacement row must fail');
  expectInvalid({ paragraphIndex: 1, action: 'replace_content', content: [['x', 2]] }, {}, 'replacement cells must be strings');
  expectInvalid({ paragraphIndex: 1, action: 'replace_content', content: [['x'], ['y']] }, { rowCount: 1 }, 'extra replacement row must fail bounds check');
  expectInvalid({ paragraphIndex: 1, action: 'replace_content', content: [['x', 'y']] }, { rowCellCounts: [1] }, 'extra replacement cell must fail per-row bounds check');
}

function testAddRowAcceptsCompatibleFormsAndRejectsTruncation() {
  for (const content of [['A', 'B'], [['A', 'B']]]) {
    const result = validateTableRequest({ paragraphIndex: 2, action: 'add_row', content }, {
      rowCount: 2,
      rowCellCounts: [2, 2]
    });
    assert.deepEqual(result.request.content, [['A', 'B']]);
  }

  expectInvalid({ paragraphIndex: 1, action: 'add_row', content: [] }, {}, 'empty row must fail');
  expectInvalid({ paragraphIndex: 1, action: 'add_row', content: [['A'], ['B']] }, {}, 'extra nested rows must not be silently discarded');
  expectInvalid({ paragraphIndex: 1, action: 'add_row', content: ['A', null] }, {}, 'row cells must be strings');
  expectInvalid({ paragraphIndex: 1, action: 'add_row', content: ['A', 'B', 'C'] }, { rowCellCounts: [2] }, 'row wider than append template must fail');
  expectInvalid({ paragraphIndex: 1, action: 'add_row', targetRow: 0, content: ['A'] }, {}, 'add_row always appends and cannot honor targetRow');
}

function testDeleteRowRequiresStrictInRangeCoordinate() {
  const result = validateTableRequest({ paragraphIndex: 4, action: 'delete_row', targetRow: 1 }, { rowCount: 2 });
  assert.deepEqual(result.request, {
    kind: 'edit_table', paragraphIndex: 4, action: 'delete_row', targetRow: 1
  });

  for (const targetRow of [undefined, null, -1, 2, 1.5, '1']) {
    expectInvalid({ paragraphIndex: 1, action: 'delete_row', targetRow }, { rowCount: 2 }, `delete row coordinate ${String(targetRow)} must fail`);
  }
  expectInvalid({ paragraphIndex: 1, action: 'delete_row' }, {}, 'missing delete row coordinate must fail');
  expectInvalid({ paragraphIndex: 1, action: 'delete_row', targetRow: 0, content: ['ignored'] }, {}, 'delete row must not silently ignore content');
}

function testUpdateCellCanonicalNestedAndLegacyShapes() {
  for (const content of ['value', ['value'], [['value']], '']) {
    const result = validateTableRequest({
      paragraphIndex: 6,
      action: 'update_cell',
      targetRow: 1,
      targetColumn: 2,
      content
    }, { rowCount: 2, rowCellCounts: [2, 3] });
    assert.equal(result.valid, true);
    assert.deepEqual(result.request.content, [[typeof content === 'string' ? content : 'value']]);
  }

  for (const request of [
    { paragraphIndex: 1, action: 'update_cell', targetRow: 0, targetColumn: 0 },
    { paragraphIndex: 1, action: 'update_cell', targetColumn: 0, content: 'x' },
    { paragraphIndex: 1, action: 'update_cell', targetRow: 0, content: 'x' }
  ]) expectInvalid(request, {}, 'update cell requires content and both coordinates');

  for (const value of [-1, 1, 0.5, '0', null]) {
    expectInvalid({ paragraphIndex: 1, action: 'update_cell', targetRow: value, targetColumn: 0, content: 'x' }, { rowCount: 1, columnCount: 1 }, `update row ${String(value)} must fail`);
    expectInvalid({ paragraphIndex: 1, action: 'update_cell', targetRow: 0, targetColumn: value, content: 'x' }, { rowCount: 1, columnCount: 1 }, `update column ${String(value)} must fail`);
  }
  expectInvalid({ paragraphIndex: 1, action: 'update_cell', targetRow: 0, targetColumn: 0, content: ['A', 'B'] }, {}, 'multi-cell legacy array must not silently truncate');
  expectInvalid({ paragraphIndex: 1, action: 'update_cell', targetRow: 0, targetColumn: 0, content: [['A', 'B']] }, {}, 'multi-cell nested array must not silently truncate');
  expectInvalid({ paragraphIndex: 1, action: 'update_cell', targetRow: 1, targetColumn: 0, content: 'x' }, { rowCount: 1, columnCount: 1 }, 'update row must be in range');
  expectInvalid({ paragraphIndex: 1, action: 'update_cell', targetRow: 0, targetColumn: 1, content: 'x' }, { rowCellCounts: [1] }, 'update column must respect row-specific bounds');
}

function testStrictSourceIndexesActionsAndLiveDimensions() {
  for (const paragraphIndex of [undefined, null, 0, -1, 1.1, '1']) {
    expectInvalid({ paragraphIndex, action: 'delete_row', targetRow: 0 }, {}, `paragraph index ${String(paragraphIndex)} must fail`);
  }
  expectInvalid({ paragraphIndex: 4, action: 'delete_row', targetRow: 0 }, { paragraphCount: 3 }, 'out-of-range source paragraph must fail');
  expectInvalid({ paragraphIndex: 1, action: 'add_column', content: ['x'] }, {}, 'unsupported column action must fail');
  expectInvalid({ paragraphIndex: 1, action: 'delete_row', targetRow: 0 }, { rowCount: 2, rowCellCounts: [2] }, 'inconsistent live table dimensions must fail');
  expectInvalid({ paragraphIndex: 1, action: 'delete_row', targetRow: 0 }, { columnCount: 0 }, 'zero live column count is not a valid table dimension');

  const normalized = validateTableRequest({ paragraphIndex: 1, action: ' UPDATE_CELL ', targetRow: 0, targetColumn: 0, content: 'x' }, {});
  assert.equal(normalized.valid, true, 'action casing/spacing matches the existing executor normalizer');
  assert.equal(normalized.request.action, 'update_cell');
}

testReplaceContentNormalizesAndChecksPerRowBounds();
testAddRowAcceptsCompatibleFormsAndRejectsTruncation();
testDeleteRowRequiresStrictInRangeCoordinate();
testUpdateCellCanonicalNestedAndLegacyShapes();
testStrictSourceIndexesActionsAndLiveDimensions();
console.log('PASS: table request validation tests');
