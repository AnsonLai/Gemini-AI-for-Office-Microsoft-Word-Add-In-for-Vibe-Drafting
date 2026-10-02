import assert from 'node:assert/strict';
import {
  createToolRegistry,
  getToolDeclarations
} from '../src/taskpane/modules/chat/tool-registry.js';

const expectedNames = [
  'apply_redlines',
  'insert_comment',
  'highlight_text',
  'perform_research',
  'navigate_to_section',
  'edit_list',
  'insert_list_item',
  'edit_table',
  'edit_section',
  'convert_headers_to_list',
  'get_selection_stats'
];

function testDeclarationsAndRecoveryNames() {
  const declarations = getToolDeclarations();
  assert.equal(declarations.length, 1);
  assert.deepEqual(declarations[0].function_declarations.map(tool => tool.name), expectedNames);
  assert.equal(declarations[0].function_declarations.length, 11);
  assert.equal(declarations[0].function_declarations[0].parameters.required[0], 'instruction');

  const registry = createToolRegistry({});
  assert.deepEqual(registry.recoveryToolNames, expectedNames);
  assert.equal(Object.isFrozen(registry.recoveryToolNames), true);
  declarations[0].function_declarations.pop();
  assert.equal(getToolDeclarations()[0].function_declarations.length, 11, 'caller mutations must not alter static definitions');
  assert.equal(registry.describe('edit_table', { paragraphIndex: 1, action: 'add_row' }).withCheckpoint, true);
  assert.equal(registry.describe('get_selection_stats', {}).withCheckpoint, false);
}

function testStructuralValidationKeepsExecutorNormalization() {
  const registry = createToolRegistry({});
  assert.equal(registry.validate('apply_redlines', { instruction: 'Fix a typo' }).valid, true);
  assert.equal(registry.validate('apply_redlines', {}).valid, false);
  assert.equal(registry.validate('edit_list', {
    startParagraphIndex: '2', endParagraphIndex: 3, newItems: ['one'], listType: 'bullet'
  }).valid, false);
  assert.equal(registry.validate('edit_table', {
    paragraphIndex: 4, action: 'add_row', content: ['cell 1', 'cell 2']
  }).valid, true, 'flat add-row content must reach the existing table validator');
  assert.equal(registry.validate('edit_table', {
    paragraphIndex: 4, action: 'update_cell', content: 'replacement', targetRow: 0, targetColumn: 1
  }).valid, true, 'scalar update-cell content must reach the existing table validator');
  assert.equal(registry.validate('edit_list', {
    startParagraphIndex: 1, endParagraphIndex: 1, newItems: [3], listType: 'bullet'
  }).valid, false, 'ordinary array fields still receive structural validation');
  assert.equal(registry.validate('edit_table', {
    paragraphIndex: 4, action: 'unknown_action'
  }).valid, false);
  assert.equal(registry.validate('edit_list', {
    startParagraphIndex: 1, endParagraphIndex: 1, newItems: ['x'], listType: 'bullet', extra: 'legacy extra'
  }).valid, true, 'undeclared optional arguments remain tolerated');
  assert.equal(registry.validate('missing_tool', {}).error.code, 'UNKNOWN_TOOL');
}

async function testExecutorArgumentMappingAndNormalization() {
  const calls = [];
  const partialMutation = {
    status: 'error', success: false, message: 'Check the document.',
    written: false, writeAttempted: true, mutationOutcome: 'partial',
    receipts: [{ id: 'r1' }], operationResults: [{ status: 'error' }],
    error: { code: 'HOST_WRITE_PARTIAL' }, showToUser: true
  };
  const results = {
    executeRedline: { status: 'ok', message: 'Redline complete.', written: true, writeAttempted: true, mutationOutcome: 'applied', showToUser: true },
    executeComment: { status: 'ok', message: 'Comment result.', written: true, showToUser: false },
    executeHighlight: { status: 'error', message: 'Highlight refused.', written: false, showToUser: false },
    executeResearch: 'Research results',
    executeNavigate: { status: 'ok', success: true, message: 'Found the section.' },
    executeEditList: partialMutation,
    executeInsertListItem: { success: true, message: 'List item inserted.' },
    executeEditTable: { success: true, message: 'Table updated.' },
    executeEditSection: { success: false, message: 'Section refused.' },
    executeConvertHeadersToList: { success: true, message: 'Headers converted.' },
    executeGetSelectionStats: { success: true, message: 'Statistics calculated.' }
  };
  const executors = Object.fromEntries(Object.entries(results).map(([name, result]) => [
    name,
    async (...args) => {
      calls.push({ name, args });
      return result;
    }
  ]));
  const registry = createToolRegistry(executors);
  const sourceBaseline = [{ index: 1, exactText: 'source' }];
  const context = { docText: 'DOCUMENT', docSourceBaseline: sourceBaseline };
  const input = [
    ['apply_redlines', { instruction: 'Edit it' }],
    ['insert_comment', { instruction: 'Comment on it' }],
    ['highlight_text', { instruction: 'Highlight it' }],
    ['perform_research', { instruction: 'Find facts' }],
    ['navigate_to_section', { instruction: 'Go to the end' }],
    ['edit_list', { startParagraphIndex: 2, endParagraphIndex: 3, newItems: ['a'], listType: 'bullet' }],
    ['insert_list_item', { afterParagraphIndex: 5, text: 'new item' }],
    ['edit_table', { paragraphIndex: 6, action: 'add_row', content: ['a', 'b'] }],
    ['edit_section', { sectionHeaderIndex: 7 }],
    ['convert_headers_to_list', { paragraphIndices: [8] }],
    ['get_selection_stats', { text: 'count this' }]
  ];
  const normalized = [];
  for (const [name, args] of input) normalized.push(await registry.execute(name, args, context));

  assert.deepEqual(calls, [
    { name: 'executeRedline', args: ['Edit it', 'DOCUMENT', sourceBaseline] },
    { name: 'executeComment', args: ['Comment on it', 'DOCUMENT'] },
    { name: 'executeHighlight', args: ['Highlight it', 'DOCUMENT', 'yellow'] },
    { name: 'executeResearch', args: ['Find facts'] },
    { name: 'executeNavigate', args: ['Go to the end', 'DOCUMENT'] },
    { name: 'executeEditList', args: [2, 3, ['a'], 'bullet', undefined] },
    { name: 'executeInsertListItem', args: [5, 'new item', undefined, sourceBaseline] },
    { name: 'executeEditTable', args: [6, 'add_row', ['a', 'b'], undefined, undefined] },
    { name: 'executeEditSection', args: [7, undefined, undefined, undefined] },
    { name: 'executeConvertHeadersToList', args: [[8], undefined, undefined] },
    { name: 'executeGetSelectionStats', args: [{ text: 'count this' }, 'DOCUMENT'] }
  ]);

  assert.equal(normalized[0].success, true);
  assert.equal(registry.describe('apply_redlines', { instruction: 'Edit it' }).withCheckpoint, true);
  assert.equal(normalized[0].hostWriteEvidence.written, true);
  assert.equal(normalized[0].mutation, results.executeRedline);
  assert.equal(normalized[0].completionText, 'Redline complete.');
  assert.equal(normalized[0].completionWithCheckpoint, true);

  assert.equal(normalized[1].success, true, 'comment success follows status plus confirmed write');
  assert.equal(normalized[1].recordSuccess, false, 'legacy comment recovery record follows showToUser');
  assert.equal(normalized[1].display.show, false);
  assert.equal(normalized[1].hostWriteEvidence.written, true);
  assert.equal(normalized[2].success, false);
  assert.equal(normalized[3].success, true, 'research dispatch historically treats returned text as success');
  assert.equal(normalized[3].completionText, 'Found search results for: "Find facts"');
  assert.equal(normalized[3].display.show, true);
  assert.equal(normalized[4].recordResult, results.executeNavigate, 'navigation preserves its raw result in recovery records');
  assert.equal(normalized[4].completionText, 'Navigated to: "Go to the end"');
  assert.equal(normalized[5].requiresInspection, true);
  assert.equal(normalized[5].hostWriteEvidence.mutationOutcome, 'partial');
  assert.deepEqual(normalized[5].hostWriteEvidence.receipts, [{ id: 'r1' }]);
  assert.equal(normalized[5].recordResult, 'Check the document.');
  assert.equal(normalized[6].success, true);
  assert.equal(normalized[7].success, true);
  assert.equal(normalized[8].success, false);
  assert.equal(normalized[9].success, true);
  assert.equal(normalized[10].success, true);
  assert.equal(normalized[10].completionText, 'Calculated statistics.');

  const description = registry.describe('edit_table', { paragraphIndex: 6, action: 'add_row' });
  assert.equal(description.withCheckpoint, true);
  assert.equal(description.progressText, 'Editing table (add_row)...');
}

async function testStructuredFailuresAndThrownExecutors() {
  const registry = createToolRegistry({});
  const unknown = await registry.execute('invented_tool', {}, {});
  assert.equal(unknown.success, false);
  assert.equal(unknown.rawResult.error.code, 'UNKNOWN_TOOL');
  assert.match(unknown.resultText, /Unknown tool/);

  const invalid = await registry.execute('edit_list', { startParagraphIndex: 'bad' }, {});
  assert.equal(invalid.rawResult.error.code, 'INVALID_TOOL_ARGUMENTS');
  assert.equal(invalid.mutating, true);

  const malformedStats = await registry.execute('get_selection_stats', { text: 42 }, {});
  assert.equal(malformedStats.rawResult.error.code, 'INVALID_TOOL_ARGUMENTS');

  const unavailable = await registry.execute('apply_redlines', { instruction: 'edit' }, {});
  assert.equal(unavailable.rawResult.error.code, 'TOOL_EXECUTOR_UNAVAILABLE');

  const failing = createToolRegistry({ executeRedline: async () => { throw new Error('executor transport failure'); } });
  await assert.rejects(
    failing.execute('apply_redlines', { instruction: 'edit' }, {}),
    /executor transport failure/
  );
}

await testDeclarationsAndRecoveryNames();
testStructuralValidationKeepsExecutorNormalization();
await testExecutorArgumentMappingAndNormalization();
await testStructuredFailuresAndThrownExecutors();
console.log('Tool registry tests passed.');
