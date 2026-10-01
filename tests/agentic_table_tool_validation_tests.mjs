import assert from 'node:assert/strict';
import fs from 'node:fs';
import { parse } from 'acorn';
import {
  initAgenticTools,
  executeEditTable,
  executeRedline,
  executeComment,
  executeHighlight,
  executeNavigate
} from '../src/taskpane/modules/commands/agentic-tools.js';

let apiKey = 'table-test-key';
initAgenticTools({
  getRequestSignal: () => null,
  loadApiKey: () => apiKey,
  loadModel: () => 'test-model',
  loadSystemMessage: () => '',
  loadRedlineSetting: () => false,
  loadRedlineAuthor: () => 'Test Editor',
  setChangeTrackingForAi: async () => ({}),
  restoreChangeTracking: async () => {},
  SAFETY_SETTINGS_BLOCK_NONE: [],
  API_LIMITS: {}
});

const previousWord = globalThis.Word;
const previousFetch = globalThis.fetch;

function createTableHarness({ rowCellCounts = [2, 2] } = {}) {
  const events = [];
  const rows = rowCellCounts.map((cellCount, index) => ({
    cellCount,
    delete() { events.push({ type: 'deleteRow', index }); }
  }));
  const table = {
    rowCount: rows.length,
    columnCount: rowCellCounts.length ? Math.max(...rowCellCounts) : 0,
    load(properties) { events.push({ type: 'table.load', properties }); },
    rows: {
      items: rows,
      load(properties) { events.push({ type: 'rows.load', properties }); }
    },
    getCell(row, column) {
      events.push({ type: 'getCell', row, column });
      if (!rows[row] || column >= rows[row].cellCount) throw Object.assign(new Error('Cell not found'), { code: 'ItemNotFound' });
      return { body: { insertText(value, location) { events.push({ type: 'insertText', row, column, value, location }); } } };
    },
    addRows(location, count, values) { events.push({ type: 'addRows', location, count, values }); }
  };
  const paragraph = {
    text: 'Cell paragraph',
    parentTableOrNullObject: { isNullObject: false },
    select() { events.push({ type: 'selectParagraph' }); },
    get parentTable() { return table; }
  };
  // Office.js exposes parentTableOrNullObject itself as the table object with its null flag.
  paragraph.parentTableOrNullObject = Object.assign(table, { isNullObject: false });
  const paragraphs = {
    items: [paragraph],
    load(properties) { events.push({ type: 'paragraphs.load', properties }); }
  };
  let runCount = 0;
  globalThis.Word = {
    InsertLocation: { replace: 'Replace', end: 'End' },
    run: async callback => {
      runCount += 1;
      return callback({
        async sync() { events.push({ type: 'sync' }); },
        document: { body: { paragraphs } }
      });
    }
  };
  return { events, get runCount() { return runCount; } };
}

function findProperty(node, name) {
  if (!node || node.type !== 'ObjectExpression') return null;
  return node.properties.find(property => property.type === 'Property'
    && (property.key.name ?? property.key.value) === name) || null;
}

function literalValue(node) {
  return node?.type === 'Literal' ? node.value : undefined;
}

function testTaskpaneSchemaUsesNestedArraysAndOnlySupportedActions() {
  const taskpaneSource = fs.readFileSync(new URL('../src/taskpane/taskpane.js', import.meta.url), 'utf8');
  const ast = parse(taskpaneSource, { ecmaVersion: 'latest', sourceType: 'module' });
  const stack = [...ast.body];
  let declaration = null;
  while (stack.length > 0) {
    const node = stack.pop();
    if (!node || typeof node !== 'object') continue;
    if (node.type === 'ObjectExpression'
      && literalValue(findProperty(node, 'name')?.value) === 'edit_table') {
      declaration = node;
      break;
    }
    for (const value of Object.values(node)) {
      if (Array.isArray(value)) stack.push(...value);
      else if (value && typeof value === 'object') stack.push(value);
    }
  }
  assert.ok(declaration, 'edit_table function declaration must exist in taskpane schema');

  const parameters = findProperty(declaration, 'parameters')?.value;
  const properties = findProperty(parameters, 'properties')?.value;
  const contentSchema = findProperty(properties, 'content')?.value;
  const actionSchema = findProperty(properties, 'action')?.value;
  assert.equal(literalValue(findProperty(contentSchema, 'type')?.value), 'ARRAY');
  const rowSchema = findProperty(contentSchema, 'items')?.value;
  assert.equal(literalValue(findProperty(rowSchema, 'type')?.value), 'ARRAY');
  const cellSchema = findProperty(rowSchema, 'items')?.value;
  assert.equal(literalValue(findProperty(cellSchema, 'type')?.value), 'STRING');
  assert.deepEqual(findProperty(actionSchema, 'enum')?.value.elements.map(literalValue), [
    'replace_content', 'add_row', 'delete_row', 'update_cell'
  ]);
  assert.doesNotMatch(literalValue(findProperty(declaration, 'description')?.value), /columns/i,
    'tool description must not advertise unavailable column operations');
}

async function testTablePreflightRejectsInvalidRequestsBeforeWord() {
  const fractional = createTableHarness();
  const fractionalResult = await executeEditTable(1, 'delete_row', undefined, 0.5);
  assert.equal(fractionalResult.success, false);
  assert.equal(fractionalResult.writeAttempted, false);
  assert.equal(fractionalResult.error.code, 'INVALID_TABLE_REQUEST');
  assert.equal(fractional.runCount, 0, 'fractional row coordinate must fail before opening Word');

  const multipleRows = createTableHarness();
  const addResult = await executeEditTable(1, 'add_row', [['first'], ['silently ignored']]);
  assert.equal(addResult.success, false);
  assert.equal(addResult.writeAttempted, false);
  assert.equal(addResult.error.code, 'INVALID_TABLE_REQUEST');
  assert.equal(multipleRows.runCount, 0, 'multi-row add request must fail before opening Word');
}

async function testLiveTableBoundsRefuseOversizedOverlayBeforeFirstWrite() {
  const harness = createTableHarness({ rowCellCounts: [1, 2] });
  const result = await executeEditTable(1, 'replace_content', [['valid'], ['too many', 'cells', 'here']]);
  assert.equal(result.success, false);
  assert.equal(result.writeAttempted, false);
  assert.equal(result.written, false);
  assert.equal(result.error.code, 'INVALID_TABLE_REQUEST');
  assert.equal(harness.runCount, 1, 'live row shape is loaded from Word before validating the overlay');
  assert.equal(harness.events.some(event => event.type === 'insertText'), false,
    'all row and cell bounds must validate before the executor stages its first cell write');
  assert.equal(harness.events.some(event => event.type === 'getCell'), false,
    'an oversized overlay must be rejected before resolving any target cell');
}

async function testNestedCanonicalInputsReachWordWithoutTruncation() {
  const updateHarness = createTableHarness();
  const update = await executeEditTable(1, 'update_cell', [['Nested cell text']], 0, 1);
  assert.equal(update.success, true);
  assert.equal(update.written, true);
  assert.deepEqual(updateHarness.events.filter(event => event.type === 'insertText').map(event => ({
    row: event.row, column: event.column, value: event.value
  })), [{ row: 0, column: 1, value: 'Nested cell text' }]);

  const rowHarness = createTableHarness();
  const addRow = await executeEditTable(1, 'add_row', [['New A', 'New B']]);
  assert.equal(addRow.success, true);
  assert.equal(addRow.written, true);
  const addRowsEvent = rowHarness.events.find(event => event.type === 'addRows');
  assert.deepEqual(addRowsEvent.values, [['New A', 'New B']]);
  assert.equal(addRowsEvent.count, 1);
}

async function testMissingApiKeyReturnsStructuredFailures() {
  apiKey = '';
  const harness = createTableHarness();
  const results = await Promise.all([
    executeRedline('change text', '[P1] Source'),
    executeComment('add a comment', '[P1] Source'),
    executeHighlight('highlight text', '[P1] Source'),
    executeNavigate('go to section', '[P1] Source')
  ]);

  for (const result of results) {
    assert.equal(typeof result, 'object');
    assert.equal(result.success, false);
    assert.equal(result.status, 'error');
    assert.equal(result.error.code, 'MISSING_API_KEY');
    assert.match(result.message, /Gemini API key/);
  }
  for (const result of results.slice(0, 3)) {
    assert.equal(result.written, false);
    assert.equal(result.writeAttempted, false);
    assert.equal(result.mutationOutcome, 'refused');
  }
  assert.equal(results[3].showToUser, true, 'navigation missing-key error is user-visible');
  assert.equal(harness.runCount, 0, 'missing API key must refuse without Word access');
  apiKey = 'table-test-key';
}

async function testNavigationSuccessAndFailureAreStructured() {
  apiKey = 'table-test-key';
  const successHarness = createTableHarness();
  globalThis.fetch = async () => ({
    ok: true,
    json: async () => ({ candidates: [{ content: { parts: [{ text: JSON.stringify({
      paragraphIndex: 1,
      navigationDescription: 'Selected the first paragraph.'
    }) }] } }] })
  });
  const success = await executeNavigate('go to the first paragraph', '[P1] Cell paragraph');
  assert.equal(success.status, 'ok');
  assert.equal(success.success, true);
  assert.equal(successHarness.events.some(event => event.type === 'selectParagraph'), true);
  assert.ok(successHarness.events.filter(event => event.type === 'sync').length >= 2,
    'navigation reports success only after loading and syncing the selection');

  const invalidHarness = createTableHarness();
  globalThis.fetch = async () => ({
    ok: true,
    json: async () => ({ candidates: [{ content: { parts: [{ text: JSON.stringify({ paragraphIndex: 2 }) }] } }] })
  });
  const invalid = await executeNavigate('go to missing paragraph', '[P1] Cell paragraph');
  assert.equal(invalid.status, 'error');
  assert.equal(invalid.success, false);
  assert.equal(invalid.error.code, 'NAVIGATION_FAILED');
  assert.equal(invalidHarness.events.some(event => event.type === 'selectParagraph'), false);
}

try {
  testTaskpaneSchemaUsesNestedArraysAndOnlySupportedActions();
  await testTablePreflightRejectsInvalidRequestsBeforeWord();
  await testLiveTableBoundsRefuseOversizedOverlayBeforeFirstWrite();
  await testNestedCanonicalInputsReachWordWithoutTruncation();
  await testMissingApiKeyReturnsStructuredFailures();
  await testNavigationSuccessAndFailureAreStructured();
  console.log('PASS: agentic table tool validation tests');
} catch (error) {
  console.error('FAIL:', error?.stack || error?.message || error);
  process.exitCode = 1;
} finally {
  globalThis.Word = previousWord;
  globalThis.fetch = previousFetch;
  apiKey = 'table-test-key';
}
