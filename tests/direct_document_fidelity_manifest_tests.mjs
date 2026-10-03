import assert from 'node:assert/strict';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { openDocx } from '@ansonlai/docx-redline-js';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fixtureDirectory = path.join(root, 'tests/fixtures/agentic-lists');
const sourceObservations = JSON.parse(readFileSync(path.join(fixtureDirectory, 'source-observations.json'), 'utf8'));
const sourceParagraphs = sourceObservations.paragraphs;
const sourceTexts = sourceParagraphs.map(paragraph => paragraph.text);
assert.equal(sourceObservations.provenance.authoredBy, 'Microsoft Word desktop COM');
assert.equal(sourceObservations.provenance.verifiedAfterSaveReopen, true);
assert.equal(sourceObservations.paragraphCount, sourceTexts.length);

const sourceText = sourceTexts.join('\n');
const observation = text => {
  const item = sourceParagraphs.find(paragraph => paragraph.text === text);
  assert.ok(item, `Missing source observation: ${text}`);
  return item;
};
const groupFor = text => text.startsWith('Bullet ')
  ? 'bullet-chain'
  : ['Number Root A', 'Number Nested Anchor', 'Number Root B', 'Number Continued Item'].includes(text)
    ? 'number-continuation'
    : text.startsWith('Number Restart ')
      ? 'number-restart'
      : null;
function numberingRow(text, listLevel = observation(text).listLevel, listString = observation(text).listString) {
  return { text, isList: true, listType: observation(text).listType, listLevel,
    ...(listString !== null && listString !== undefined ? { listString } : {}), group: groupFor(text) };
}
const sourceNumbering = sourceTexts.filter(text => observation(text).isList).map(text => numberingRow(text));
const insertParagraph = (paragraphs, anchorText, text) => {
  const anchorIndex = paragraphs.indexOf(anchorText);
  assert.notEqual(anchorIndex, -1, `Fixture is missing paragraph: ${anchorText}`);
  return [...paragraphs.slice(0, anchorIndex + 1), text, ...paragraphs.slice(anchorIndex + 1)];
};
const replaceParagraph = (paragraphs, oldText, newText) => {
  const matches = paragraphs.filter(text => text === oldText).length;
  assert.equal(matches, 1, `Expected one source paragraph: ${oldText}`);
  return paragraphs.map(text => text === oldText ? newText : text);
};
const insertListItem = (name, targetText, text, expectedLabels = {}, options = {}) => {
  const target = observation(targetText);
  const acceptedParagraphs = insertParagraph(sourceTexts, targetText, text);
  const targetListIndex = sourceNumbering.findIndex(item => item.text === targetText);
  assert.notEqual(targetListIndex, -1, `List insertion anchor is not a list item: ${targetText}`);
  const listString = expectedLabels[text] || (target.listString || '');
  const inserted = { text, isList: true, listType: target.listType, listLevel: target.listLevel,
    ...(listString ? { listString } : {}), group: groupFor(targetText) };
  const acceptedRows = sourceNumbering.map(item => expectedLabels[item.text]
    ? { ...item, listString: expectedLabels[item.text] }
    : item);
  acceptedRows.splice(targetListIndex + 1, 0, inserted);
  return {
    name,
    source: 'nested-lists-source.docx',
    agenticRequest: {
      type: 'insert_list_item', targetText, text,
      // Word COM's observed index is fixture metadata only; the collector resolves targetText to an opaque handle.
      afterParagraphIndex: target.paragraphIndex
    },
    productionOptions: {
      generateRedlines: options.generateRedlines !== false,
      priorTrackingMode: options.priorTrackingMode || 'off'
    },
    expectedSourceText: sourceText,
    expectedAcceptedText: acceptedParagraphs.join('\n'),
    expectedRejectedText: options.expectedRejectedText || sourceText,
    expectedComments: 0,
    ...(Number.isInteger(options.expectedExactBodyRevisions)
      ? { expectedExactBodyRevisions: options.expectedExactBodyRevisions }
      : { expectedMinimumBodyRevisions: 1 }),
    expectedFormatting: [{ sourceText: 'Untouched bold sentinel', acceptedText: 'Untouched bold sentinel', bold: true, italic: false }],
    expectedNumbering: {
      source: sourceNumbering, accepted: acceptedRows,
      rejected: options.generateRedlines === false ? acceptedRows : sourceNumbering
    },
    expectedEngineReferencePackages: false
  };
};

function textCase(name, edits, acceptedParagraphs, extras = {}) {
  return {
    name,
    source: 'nested-lists-source.docx',
    edits,
    productionOptions: { generateRedlines: true, priorTrackingMode: 'off' },
    expectedSourceText: sourceText,
    expectedAcceptedText: acceptedParagraphs.join('\n'),
    expectedRejectedText: sourceText,
    expectedComments: 0,
    expectedMinimumBodyRevisions: 1,
    expectedFormatting: [{ sourceText: 'Untouched bold sentinel', acceptedText: 'Untouched bold sentinel', bold: true, italic: false }],
    expectedNumbering: { source: sourceNumbering, accepted: sourceNumbering, rejected: sourceNumbering },
    expectedEngineReferencePackages: false,
    ...extras
  };
}

const listCases = [
  insertListItem('same-level-insert-root-bullet', 'Bullet Root B', 'Inserted root bullet'),
  insertListItem('same-level-insert-level1-decimal', 'Number Root B', 'Inserted level-one decimal item', {
    'Inserted level-one decimal item': '3.', 'Number Continued Item': '4.'
  }),
  insertListItem('same-level-insert-continued-decimal', 'Number Continued Item', 'Inserted continued decimal item', {
    'Inserted continued decimal item': '4.'
  }),
  insertListItem('same-level-insert-separate-restart', 'Number Restart Root', 'Inserted separate restart item', {
    'Inserted separate restart item': '2.', 'Number Restart Nested': '2.1.'
  })
];
listCases[3].productionOptions.priorTrackingMode = 'trackAll';
const noRedlineListCase = insertListItem(
  'same-level-insert-redlines-disabled-restores-tracking',
  'Bullet Root B',
  'Inserted bullet with redlines disabled',
  {},
  { generateRedlines: false, priorTrackingMode: 'trackAll', expectedExactBodyRevisions: 0,
    expectedRejectedText: insertParagraph(sourceTexts, 'Bullet Root B', 'Inserted bullet with redlines disabled').join('\n') }
);

const commentSource = 'tests/fixtures/wp6/threaded-comments.docx';
const commentSourceTexts = [
  'The Supplier shall deliver the goods within thirty days.',
  'Payment is due upon receipt of invoice.',
  'This agreement is governed by local law.'
];
const commentSourceInspection = openDocx(new Uint8Array(readFileSync(path.join(root, commentSource)))).inspect();
assert.equal(commentSourceInspection.status, 'ok');
assert.equal(commentSourceInspection.comments.length, 3, 'the Word-authored restore checkpoint source contains three existing comments');
assert.deepEqual(commentSourceInspection.paragraphs.map(paragraph => paragraph.exactText), commentSourceTexts);
const commentRevisionSentinel = 'Existing tracked checkpoint revision sentinel.';
const commentPreparedTexts = insertParagraph(commentSourceTexts, commentSourceTexts[2], commentRevisionSentinel);
const commentLiteralSource = commentSourceTexts[2];
const commentLiteralAccepted = 'This agreement is governed by applicable local law.';
const literalCase = textCase('direct-literal-replacement-preserves-comment-threads', [
  { type: 'replace_text', targetText: commentLiteralSource, find: 'governed by local law', replace: 'governed by applicable local law' }
], insertParagraph(replaceParagraph(commentSourceTexts, commentLiteralSource, commentLiteralAccepted), commentLiteralAccepted, commentRevisionSentinel), {
  source: commentSource,
  usesPreparedSource: true,
  sourcePreparation: [{ type: 'tracked-append', afterText: commentSourceTexts[2], text: commentRevisionSentinel,
    trackingMode: 'trackAll', author: 'Other reviewer' }],
  productionOptions: { generateRedlines: false, priorTrackingMode: 'trackAll' },
  expectedSourceText: commentPreparedTexts.join('\n'),
  expectedRejectedText: replaceParagraph(commentSourceTexts, commentLiteralSource, commentLiteralAccepted).join('\n'),
  expectedRestoredAcceptedText: commentPreparedTexts.join('\n'),
  expectedRestoredRejectedText: commentSourceTexts.join('\n'),
  expectedComments: 3,
  expectedExactBodyRevisions: 1,
  expectedRestoredBodyRevisions: 1,
  expectedFormatting: [],
  expectedNumbering: { source: [], accepted: [], rejected: [] },
  revisionIdentityCheck: { author: 'Other reviewer' },
  restorationHandleText: commentLiteralAccepted
});
assert.equal(literalCase.expectedPrewriteRefusal, undefined,
  'the default manifest keeps the untracked literal-replacement capability gate as an applied edit expectation');

const supportedRevisionReplacement = textCase('direct-untracked-replacement-preserves-pending-revision', [
  { type: 'replace_text', targetText: commentLiteralSource, find: 'local', replace: 'national' }
], insertParagraph(replaceParagraph(commentSourceTexts, commentLiteralSource,
  'This agreement is governed by national law.'), 'This agreement is governed by national law.', commentRevisionSentinel), {
  source: commentSource,
  usesPreparedSource: true,
  sourcePreparation: [{ type: 'tracked-append', afterText: commentSourceTexts[2], text: commentRevisionSentinel,
    trackingMode: 'trackAll', author: 'Other reviewer' }],
  productionOptions: { generateRedlines: false, priorTrackingMode: 'trackAll' },
  expectedSourceText: commentPreparedTexts.join('\n'),
  expectedRejectedText: replaceParagraph(commentSourceTexts, commentLiteralSource,
    'This agreement is governed by national law.').join('\n'),
  expectedRestoredAcceptedText: commentPreparedTexts.join('\n'),
  expectedRestoredRejectedText: commentSourceTexts.join('\n'),
  expectedComments: 3,
  expectedExactBodyRevisions: 1,
  expectedRestoredBodyRevisions: 1,
  expectedEngineReferencePackages: false,
  expectedNumbering: { source: [], accepted: [], rejected: [] },
  expectedFormatting: [],
  revisionIdentityCheck: { author: 'Other reviewer' },
  restorationHandleText: 'This agreement is governed by national law.'
});

const untrackedFormatCase = {
  name: 'direct-untracked-format-with-foreign-pending-revision',
  source: commentSource,
  usesPreparedSource: true,
  sourcePreparation: [{ type: 'tracked-append', afterText: commentSourceTexts[2], text: commentRevisionSentinel,
    trackingMode: 'trackAll', author: 'Other reviewer' }],
  edits: [{ type: 'format_text', targetText: commentSourceTexts[2], find: 'governed by local law', properties: { underline: true } }],
  productionOptions: { generateRedlines: false, priorTrackingMode: 'trackAll' },
  expectedSourceText: commentPreparedTexts.join('\n'),
  expectedAcceptedText: commentPreparedTexts.join('\n'),
  expectedRejectedText: commentSourceTexts.join('\n'),
  expectedComments: 3,
  expectedExactBodyRevisions: 1,
  expectedEngineReferencePackages: false,
  expectedNumbering: { source: [], accepted: [], rejected: [] },
  expectedFormatting: [{ sourceText: 'governed by local law', acceptedText: 'governed by local law',
    bold: false, italic: false, underlineSource: false, underlineRejected: true, underlineAccepted: true }],
  revisionIdentityCheck: { author: 'Other reviewer' },
  formatReadback: { text: 'governed by local law', before: 'None', after: 'Single' }
};

const sourceThreadIdentities = commentSourceInspection.comments.map(comment => {
  const parent = comment.parentCommentId == null ? null
    : commentSourceInspection.comments.find(candidate => candidate.id === comment.parentCommentId);
  return {
    author: comment.author,
    text: comment.text,
    parent: parent ? { author: parent.author, text: parent.text } : null,
    done: comment.done === true
  };
});
const firstWorkflowComment = 'Please confirm the requested delivery timing.';
const secondWorkflowComment = 'Please confirm this revised governing-law sentence.';
const workflowFinalParagraph = 'This agreement is governed by applicable local law.';
const workflowExpectedThreadIdentities = [
  ...sourceThreadIdentities,
  { author: 'WP4 reviewer', text: firstWorkflowComment, parent: null, done: false },
  { author: 'WP4 reviewer', text: secondWorkflowComment, parent: null, done: false }
];
const workflowRestoredThreadIdentities = [
  ...sourceThreadIdentities,
  { author: 'WP4 reviewer', text: firstWorkflowComment, parent: null, done: false }
];
const commentWorkflowCase = textCase('direct-threaded-comments-sequential-merge-and-restore', [
  { type: 'rewrite_paragraph', targetText: commentSourceTexts[2], text: workflowFinalParagraph },
  { type: 'comment', targetText: commentSourceTexts[2], anchor: 'This agreement is governed by', text: secondWorkflowComment }
], replaceParagraph(commentSourceTexts, commentSourceTexts[2], workflowFinalParagraph), {
  source: commentSource,
  expectedSourceText: commentSourceTexts.join('\n'),
  expectedRejectedText: commentSourceTexts.join('\n'),
  expectedComments: 5,
  expectedMinimumBodyRevisions: 1,
  expectedNumbering: { source: [], accepted: [], rejected: [] },
  expectedFormatting: [],
  expectedThreadIdentities: workflowExpectedThreadIdentities,
  expectedRestoredThreadIdentities: workflowRestoredThreadIdentities,
  expectedRestoredAcceptedText: commentSourceTexts.join('\n'),
  expectedRestoredRejectedText: commentSourceTexts.join('\n'),
  expectedRestoredBodyRevisions: 0,
  restorationHandleText: workflowFinalParagraph,
  commentWorkflow: {
    expectedOriginalComments: 3,
    expectedFirstAnchorAfterRewrite: 'applicable local law',
    first: { targetText: commentSourceTexts[2], anchor: 'local law', text: firstWorkflowComment },
    secondBatch: [
      { type: 'rewrite_paragraph', targetText: commentSourceTexts[2], text: workflowFinalParagraph },
      { type: 'comment', targetText: commentSourceTexts[2], anchor: 'This agreement is governed by', text: secondWorkflowComment }
    ],
    secondCommentText: secondWorkflowComment,
    expectedFinalParagraph: workflowFinalParagraph
  }
});

const rewriteSource = 'Plain paragraph between headers.';
const rewriteAccepted = 'Plain paragraph between the two header candidates.';
const rewriteCommentText = 'Confirm this transition remains intentionally between the headings.';
const rewriteCommentCase = textCase('direct-paragraph-rewrite-and-comment', [
  { type: 'rewrite_paragraph', targetText: rewriteSource, text: rewriteAccepted },
  { type: 'comment', targetText: rewriteSource, anchor: 'Plain paragraph', text: rewriteCommentText }
], replaceParagraph(sourceTexts, rewriteSource, rewriteAccepted), { expectedComments: 1 });

const insertAnchor = 'Plain transition after bullet list.';
const plainInsertion = 'Inserted ordinary paragraph.';
const insertionCase = textCase('direct-plain-paragraph-insertion', [
  { type: 'insert_paragraph', targetText: insertAnchor, position: 'after', text: plainInsertion }
], insertParagraph(sourceTexts, insertAnchor, plainInsertion));

const deleteTarget = 'Plain paragraph between supported headers.';
const deleteParagraphs = sourceTexts.filter(text => text !== deleteTarget);
assert.equal(deleteParagraphs.length, sourceTexts.length - 1);
const deletionCase = textCase('direct-plain-paragraph-deletion', [
  { type: 'delete_paragraph', targetText: deleteTarget }
], deleteParagraphs);

const tableSource = 'tests/fixtures/golden/sample-nda.docx';
const sourceTableCells = [
  ['DISCLOSING PARTY:', 'RECEIVING PARTY:'],
  ['_________________________', '_________________________'],
  ['By: [Name]', 'By: [Name]'],
  ['Title:', 'Title:']
];
const acceptedTableCells = sourceTableCells.map(row => [...row]);
acceptedTableCells[0][1] = 'RECEIVING PARTY (reviewed):';
const formatTarget = 'This Non-Disclosure Agreement (the “Agreement”) is effective as of the date of last signature below (the “Effective Date”) by and between:';
const formatPrefix = 'This Non-Disclosure Agreement';
const formatSuffix = ' (the “Agreement”) is effective as of the date of last signature below (the “Effective Date”) by and between:';
const unchangedTableStructure = {
  rows: 4, columns: 2,
  source: sourceTableCells, accepted: sourceTableCells, rejected: sourceTableCells
};
const formatCase = {
  name: 'direct-basic-underline-format',
  source: tableSource,
  edits: [{ type: 'format_text', targetText: formatTarget, find: formatPrefix, properties: { underline: true } }],
  tableStructure: unchangedTableStructure,
  expectedSourceText: null,
  expectedAcceptedText: null,
  expectedRejectedText: null,
  expectedComments: 0,
  expectedMinimumBodyRevisions: 1,
  expectedEngineReferencePackages: false,
  productionOptions: { generateRedlines: true, priorTrackingMode: 'off' },
  expectedFormatting: [
    { sourceText: formatPrefix, acceptedText: formatPrefix, bold: false, italic: false,
      underlineRejected: false, underlineAccepted: true },
    { sourceText: formatSuffix, acceptedText: formatSuffix, bold: false, italic: false,
      underlineRejected: false, underlineAccepted: false }
  ]
};
const tableCase = {
  name: 'direct-table-cell-literal-replacement',
  source: tableSource,
  edits: [{ type: 'replace_text', targetText: 'RECEIVING PARTY:', find: 'RECEIVING PARTY:', replace: 'RECEIVING PARTY (reviewed):' }],
  tableStructure: { ...unchangedTableStructure, accepted: acceptedTableCells },
  expectedSourceText: null,
  expectedAcceptedText: null,
  expectedRejectedText: null,
  expectedComments: 0,
  expectedMinimumBodyRevisions: 1,
  expectedEngineReferencePackages: false,
  productionOptions: { generateRedlines: true, priorTrackingMode: 'off' }
};

const cases = [...listCases, noRedlineListCase, literalCase, supportedRevisionReplacement, untrackedFormatCase, commentWorkflowCase,
  rewriteCommentCase, insertionCase, deletionCase, formatCase, tableCase];
assert.equal(listCases.length, 4);
assert.deepEqual(listCases.map(testCase => testCase.agenticRequest.afterParagraphIndex), [6, 10, 12, 14]);
assert.ok(listCases.every(testCase => testCase.agenticRequest.type === 'insert_list_item'));
assert.ok(listCases.every(testCase => testCase.expectedSourceText && testCase.expectedAcceptedText && testCase.expectedRejectedText));
assert.equal(rewriteCommentCase.edits.length, 2);
assert.equal(commentWorkflowCase.commentWorkflow.secondBatch.length, 2);
assert.equal(untrackedFormatCase.sourcePreparation[0].author, 'Other reviewer');
assert.equal(commentWorkflowCase.expectedThreadIdentities.length, 5);
assert.equal(commentWorkflowCase.expectedRestoredThreadIdentities.length, 4);
assert.equal(tableCase.tableStructure.source.length, tableCase.tableStructure.rows);
assert.equal(tableCase.tableStructure.source[0].length, tableCase.tableStructure.columns);

const outputArg = process.argv.indexOf('--export-dir');
const outputDirectory = outputArg >= 0 ? path.resolve(process.argv[outputArg + 1])
  : path.join(root, '.cache/reliability/direct-word-document-tools');
assert.ok(outputDirectory && !String(process.argv[outputArg + 1] || '').startsWith('--'), '--export-dir requires a directory');
mkdirSync(outputDirectory, { recursive: true });
const sourcePaths = new Map([
  ['nested-lists-source.docx', path.join(fixtureDirectory, 'nested-lists-source.docx')],
  [commentSource, path.join(root, commentSource)],
  [tableSource, path.join(root, tableSource)]
]);
const manifest = {
  schemaVersion: 2,
  provenance: {
    source: 'Word-authored and save/reopened source fixtures; list and text expectations are independently specified from source observations and requested edits.',
    sourceObservations: 'tests/fixtures/agentic-lists/source-observations.json',
    routeAssertion: 'The collector must resolve exact inspected anchor text to opaque handles and invoke createDocumentTools directly. List indices are oracle metadata only.',
    resolutionOracle: 'Independent Word validation checks the exported source, tracked, Accept All, and Reject All states for text, comments, formatting, list identity, continuation, restart, and table cell preservation.'
  },
  cases: cases.map(testCase => ({
    ...testCase,
    source: path.relative(outputDirectory, sourcePaths.get(testCase.source)).replaceAll('\\', '/')
  }))
};
const manifestPath = path.join(outputDirectory, 'manifest.json');
writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`PASS: exported ${cases.length} direct Word document-tool cases (${listCases.length} retained list cases) to ${manifestPath}`);
