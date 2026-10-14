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

// ---- List actions (convert_to_list, change_list_format, continue_list) -------------------
// Sources are Word-authored (scripts/create-agentic-list-action-fixtures.ps1) and every expectation below is
// written from the requested edit and the user-visible result (Word's ListString), not from engine output.
const listActionObservations = JSON.parse(readFileSync(path.join(fixtureDirectory, 'list-actions-source-observations.json'), 'utf8'));
assert.equal(listActionObservations.provenance.authoredBy, 'Microsoft Word desktop COM');
assert.equal(listActionObservations.provenance.verifiedAfterSaveReopen, true);
const listActionSource = caseId => {
  const source = listActionObservations.sources.find(item => item.case === caseId);
  assert.ok(source, `Missing list-action source observation: ${caseId}`);
  return { file: source.file, rows: source.paragraphs, texts: source.paragraphs.map(item => item.text) };
};
const plainRow = text => ({ text, isList: false });
const listRow = (text, listString, group, listLevel = 1) => ({ text, isList: true, listLevel, listString, ...(group ? { group } : {}) });
const plainRows = source => source.rows.map(item => plainRow(item.text));
// Source and Reject All rows come from what Word itself observed in the authored file.
const observedRows = (source, groups = {}) => source.rows.map(item => item.isList
  ? { text: item.text, isList: true, listType: item.listType, listLevel: item.listLevel, listString: item.listString,
    ...(groups[item.text] ? { group: groups[item.text] } : {}) }
  : plainRow(item.text));
function listActionCase(name, caseId, edits, acceptedTexts, acceptedRows, rejectedTexts, rejectedRows, sourceRows = null) {
  const source = listActionSource(caseId);
  return {
    name,
    source: `tests/fixtures/agentic-lists/${source.file}`,
    edits,
    productionOptions: { generateRedlines: true, priorTrackingMode: 'off' },
    expectedSourceText: source.texts.join('\n'),
    expectedAcceptedText: acceptedTexts.join('\n'),
    expectedRejectedText: rejectedTexts.join('\n'),
    expectedComments: 0,
    expectedMinimumBodyRevisions: 1,
    expectedFormatting: [],
    expectedNumbering: { source: sourceRows || observedRows(source), accepted: acceptedRows, rejected: rejectedRows },
    expectedEngineReferencePackages: false
  };
}

// L1: three typed headers separated by body paragraphs become one decimal list with one convert_to_list per header.
const l1 = listActionSource('L1');
assert.deepEqual(l1.texts, ['Lead-in paragraph.', '1. Alpha', 'Alpha body paragraph.', '2. Beta', 'Beta body paragraph.', '3. Gamma', 'Closing paragraph.']);
assert.ok(l1.rows.every(item => !item.isList), 'L1 typed headers are plain Word paragraphs');
const listHeadersCase = listActionCase('list-convert-noncontiguous-typed-headers', 'L1', [
  { type: 'convert_to_list', targetText: '1. Alpha' },
  { type: 'convert_to_list', targetText: '2. Beta' },
  { type: 'convert_to_list', targetText: '3. Gamma' }
], ['Lead-in paragraph.', 'Alpha', 'Alpha body paragraph.', 'Beta', 'Beta body paragraph.', 'Gamma', 'Closing paragraph.'], [
  plainRow('Lead-in paragraph.'), listRow('Alpha', '1.', 'headers'), plainRow('Alpha body paragraph.'),
  listRow('Beta', '2.', 'headers'), plainRow('Beta body paragraph.'), listRow('Gamma', '3.', 'headers'), plainRow('Closing paragraph.')
], l1.texts, plainRows(l1));

// L2: typed letter headers starting at C. keep their typed start value.
const l2 = listActionSource('L2');
assert.deepEqual(l2.texts, ['Preamble paragraph.', 'C. Scope', 'Scope body paragraph.', 'D. Fees', 'Fees body paragraph.']);
assert.ok(l2.rows.every(item => !item.isList));
const listLetterCase = listActionCase('list-convert-letter-headers-start-at-c', 'L2', [
  { type: 'convert_to_list', targetText: 'C. Scope', format: 'upperLetter' },
  { type: 'convert_to_list', targetText: 'D. Fees', format: 'upperLetter' }
], ['Preamble paragraph.', 'Scope', 'Scope body paragraph.', 'Fees', 'Fees body paragraph.'], [
  plainRow('Preamble paragraph.'), listRow('Scope', 'C.', 'letters'), plainRow('Scope body paragraph.'),
  listRow('Fees', 'D.', 'letters'), plainRow('Fees body paragraph.')
], l2.texts, plainRows(l2));

// L3: change_list_format on a Word bullet list makes the whole list decimal; Reject All restores the bullets.
const l3 = listActionSource('L3');
assert.deepEqual(l3.texts, ['Lead-in paragraph.', 'Passport', 'Driver licence', 'Health card', 'Trailing paragraph.']);
assert.deepEqual(l3.rows.filter(item => item.isList).map(item => item.listType), [2, 2, 2], 'L3 source items are Word bullets');
const bulletGroups = { Passport: 'bullets', 'Driver licence': 'bullets', 'Health card': 'bullets' };
const listBulletsCase = listActionCase('list-change-format-bullets-to-decimal', 'L3', [
  { type: 'change_list_format', targetText: 'Driver licence', format: 'decimal' }
], l3.texts, [
  plainRow('Lead-in paragraph.'), listRow('Passport', '1.', 'decimal'), listRow('Driver licence', '2.', 'decimal'),
  listRow('Health card', '3.', 'decimal'), plainRow('Trailing paragraph.')
], l3.texts, observedRows(l3, bulletGroups), observedRows(l3, bulletGroups));

// L4: a second numbered list (restarted at 1) continues the first.
const l4 = listActionSource('L4');
assert.deepEqual(l4.rows.filter(item => item.isList).map(item => item.listString), ['1.', '2.', '1.', '2.'], 'L4 source has two separate decimal lists');
const twoListGroups = { 'Alpha item': 'first', 'Beta item': 'first', 'Gamma item': 'second', 'Delta item': 'second' };
const listContinueCase = listActionCase('list-continue-second-list-after-first', 'L4', [
  { type: 'continue_list', targetText: 'Gamma item', fromText: 'Beta item' }
], l4.texts, [
  plainRow('Opening paragraph.'), listRow('Alpha item', '1.', 'joined'), listRow('Beta item', '2.', 'joined'), plainRow('Interlude paragraph.'),
  listRow('Gamma item', '3.', 'joined'), listRow('Delta item', '4.', 'joined'), plainRow('Closing paragraph.')
], l4.texts, observedRows(l4, twoListGroups), observedRows(l4, twoListGroups));

// L5: a changed-text header plus conversion in one batch: the new wording takes the label, Reject All restores the typed header.
const l5 = listActionSource('L5');
assert.deepEqual(l5.texts, ['Lead-in paragraph.', '1. Old title', 'Header body paragraph.']);
const listChangedHeaderCase = listActionCase('list-rewrite-and-convert-changed-header', 'L5', [
  { type: 'rewrite_paragraph', targetText: '1. Old title', text: '1. New title' },
  { type: 'convert_to_list', targetText: '1. Old title' }
], ['Lead-in paragraph.', 'New title', 'Header body paragraph.'], [
  plainRow('Lead-in paragraph.'), listRow('New title', '1.'), plainRow('Header body paragraph.')
], l5.texts, plainRows(l5));

// L6 (safety control, not a capability pass): "I." without a format is ambiguous (Roman numeral or letter). The library
// refuses the whole batch, so Word must see no content write, no checkpoint and an unchanged document.
const l6 = listActionSource('L6');
assert.deepEqual(l6.texts, ['Lead-in paragraph.', 'I. Introduction', 'Introduction body paragraph.']);
const listAmbiguousRefusal = {
  ...listActionCase('list-convert-ambiguous-roman-marker-refusal', 'L6', [
    { type: 'convert_to_list', targetText: 'I. Introduction' }
  ], l6.texts, plainRows(l6), l6.texts, plainRows(l6)),
  expectedPrewriteRefusal: 'BATCH_OPERATION_FAILED',
  expectedRefusalOperation: { editIndex: 0, code: 'LIST_MARKER_AMBIGUOUS', allowedFormats: ['upperLetter', 'upperRoman'] },
  refusalCapability: 'list-marker-ambiguity-no-write-refusal'
};
delete listAmbiguousRefusal.expectedMinimumBodyRevisions;
const listActionCases = [listHeadersCase, listLetterCase, listBulletsCase, listContinueCase, listChangedHeaderCase];

const existingCases = [...listCases, noRedlineListCase, literalCase, supportedRevisionReplacement, untrackedFormatCase, commentWorkflowCase,
  rewriteCommentCase, insertionCase, deletionCase, formatCase, tableCase];
const cases = [...existingCases, ...listActionCases, listAmbiguousRefusal];
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
assert.equal(existingCases.length, 14, 'the fourteen existing cases are unchanged and stay first');
assert.equal(listActionCases.length, 5, 'five required list-action capability cases');
assert.equal(cases.length, 20, 'fourteen existing cases, five list-action capability cases and one no-write refusal control');
assert.deepEqual(cases.filter(testCase => testCase.expectedPrewriteRefusal).map(testCase => testCase.name), [listAmbiguousRefusal.name],
  'the ambiguous-marker refusal is the only no-write control; it is not a capability pass');
assert.deepEqual(listActionCases.map(testCase => testCase.edits.map(edit => edit.type)), [
  ['convert_to_list', 'convert_to_list', 'convert_to_list'], ['convert_to_list', 'convert_to_list'], ['change_list_format'],
  ['continue_list'], ['rewrite_paragraph', 'convert_to_list']]);
assert.ok(listActionCases.every(testCase => testCase.expectedNumbering.accepted.some(item => item.isList)));

const outputArg = process.argv.indexOf('--export-dir');
const outputDirectory = outputArg >= 0 ? path.resolve(process.argv[outputArg + 1])
  : path.join(root, '.cache/reliability/direct-word-document-tools');
assert.ok(outputDirectory && !String(process.argv[outputArg + 1] || '').startsWith('--'), '--export-dir requires a directory');
mkdirSync(outputDirectory, { recursive: true });
const sourcePaths = new Map([
  ['nested-lists-source.docx', path.join(fixtureDirectory, 'nested-lists-source.docx')],
  ...[...listActionCases, listAmbiguousRefusal].map(testCase => [testCase.source, path.join(root, testCase.source)]),
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
console.log(`PASS: exported ${cases.length} direct Word document-tool cases (${listCases.length} retained list cases, ${listActionCases.length} list-action capability cases, 1 no-write refusal control) to ${manifestPath}`);
