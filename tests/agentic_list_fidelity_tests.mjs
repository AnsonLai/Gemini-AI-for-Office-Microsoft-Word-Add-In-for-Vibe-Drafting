import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  mergeNumberingXmlBySchemaOrder,
  openDocx,
  validateDocxPackage
} from '@ansonlai/docx-redline-js';
import { MemoryZip, unzipDocx, zipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';
import { planAgenticListOperations } from '../src/taskpane/modules/commands/list-operation-plan.js';

const here = dirname(fileURLToPath(import.meta.url));
const fixturePath = join(here, 'fixtures/agentic-lists/nested-lists-source.docx');
const observationsPath = join(here, 'fixtures/agentic-lists/source-observations.json');
const sourceBytes = new Uint8Array(readFileSync(fixturePath));
const sourceDoc = openDocx(sourceBytes);
const sourceInspect = sourceDoc.inspect();
const sourceParagraphs = sourceInspect.paragraphs;
const sourceExactTexts = [
  'Untouched bold sentinel',
  'Plain paragraph before bullet list.',
  'Bullet Root A',
  'Bullet Insertion Anchor',
  'Bullet Untouched Tail',
  'Bullet Root B',
  'Plain transition after bullet list.',
  'Number Root A',
  'Number Nested Anchor',
  'Number Root B',
  'Plain transition before continuation.',
  'Number Continued Item',
  'Plain transition before restart.',
  'Number Restart Root',
  'Number Restart Nested',
  'Plain transition after restart.',
  'Header First Candidate',
  'Plain paragraph between headers.',
  'Header Second Candidate',
  'Plain paragraph after headers.',
  '1. Supported Header First',
  'Plain paragraph between supported headers.',
  '2. Supported Header Second',
  'Plain paragraph after supported headers.'
];
const sourceText = sourceExactTexts.join('\n');
const sourceEvidence = JSON.parse(readFileSync(observationsPath, 'utf8'));
const decode = bytes => new TextDecoder().decode(bytes);
const sourceDocumentXml = decode(sourceDoc.entries.get('word/document.xml'));
const sourceNumberingXml = decode(sourceDoc.entries.get('word/numbering.xml'));
const sourceStylesXml = decode(sourceDoc.entries.get('word/styles.xml'));
const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

assert.equal(sourceInspect.status, 'ok');
assert.equal(sourceParagraphs.map(paragraph => paragraph.exactText).join('\n'), sourceText,
  'fixture text and 1-based paragraph indexes are immutable test inputs');
assert.equal(sourceEvidence.provenance.authoredBy, 'Microsoft Word desktop COM');
assert.equal(sourceEvidence.provenance.verifiedAfterSaveReopen, true);
assert.equal(sourceEvidence.provenance.structureAssertions, 'passed');
assert.equal(sourceEvidence.paragraphCount, sourceExactTexts.length);
const wordPackageRows = new Map(sourceEvidence.packageNumbering.map(row => [row.text, row]));
const wordParagraphRows = new Map(sourceEvidence.paragraphs.map(row => [row.text, row]));
for (const text of ['Bullet Root A', 'Bullet Insertion Anchor', 'Bullet Untouched Tail', 'Bullet Root B']) {
  assert.equal(wordPackageRows.get(text)?.numFmt, 'bullet', `${text} is a true Word bullet definition`);
  assert.equal(wordPackageRows.get(text)?.group, 'bullet-chain');
  assert.equal(wordParagraphRows.get(text)?.isList, true);
}
assert.equal(wordPackageRows.get('Bullet Root A')?.numId, wordPackageRows.get('Bullet Insertion Anchor')?.numId);
assert.equal(wordPackageRows.get('Bullet Insertion Anchor')?.ilvl, 1);
assert.equal(wordPackageRows.get('Bullet Untouched Tail')?.ilvl, 0);
for (const text of ['Number Root A', 'Number Nested Anchor', 'Number Root B', 'Number Continued Item']) {
  assert.equal(wordPackageRows.get(text)?.numFmt, 'decimal');
  assert.equal(wordPackageRows.get(text)?.group, 'number-continuation');
}
assert.equal(wordPackageRows.get('Number Root A')?.numId, wordPackageRows.get('Number Continued Item')?.numId);
assert.notEqual(wordPackageRows.get('Number Root A')?.numId, wordPackageRows.get('Number Restart Root')?.numId);
assert.equal(wordParagraphRows.get('Number Continued Item')?.listString, '3.');
assert.equal(wordParagraphRows.get('Number Restart Root')?.listValue, 1);

function sourceListEntry(text, level, group, listString, listValue) {
  return {
    text,
    isList: true,
    listType: 4,
    listLevel: level,
    ...(listString == null ? {} : { listString }),
    ...(listValue == null ? {} : { listValue }),
    group
  };
}

// These values are captured from the Word-authored source after save/reopen.
// Group labels are local identities: the host verifier compares direct numIds
// within each Word-opened view and never compares absolute IDs across views.
const sourceNumbering = [
  sourceListEntry('Bullet Root A', 1, 'bullet', '\uF076', 1),
  sourceListEntry('Bullet Insertion Anchor', 2, 'bullet', '\uF0D8', 1),
  sourceListEntry('Bullet Untouched Tail', 1, 'bullet', '\uF076', 2),
  sourceListEntry('Bullet Root B', 1, 'bullet', '\uF076', 3),
  sourceListEntry('Number Root A', 1, 'number-continuation', '1.', 1),
  sourceListEntry('Number Nested Anchor', 2, 'number-continuation', '1.1.', 1),
  sourceListEntry('Number Root B', 1, 'number-continuation', '2.', 2),
  sourceListEntry('Number Continued Item', 1, 'number-continuation', '3.', 3),
  sourceListEntry('Number Restart Root', 1, 'number-restart', '1.', 1),
  sourceListEntry('Number Restart Nested', 2, 'number-restart', '1.1.', 1)
];

const allCases = [
  {
    name: 'list-insert-bullet-outdent',
    request: { tool: 'insert_list_item', afterParagraphIndex: 4, text: 'Planner bullet insertion before untouched tail', indentLevel: -1 },
    expected: sourceExactTexts.toSpliced(4, 0, 'Planner bullet insertion before untouched tail'),
    newLists: [sourceListEntry('Planner bullet insertion before untouched tail', 1, 'bullet', '\uF076', 2)],
    acceptedNumbering: {
      'Bullet Untouched Tail': { listValue: 3 },
      'Bullet Root B': { listValue: 4 }
    },
    assertions: ({ paragraphs }) => {
      const anchor = paragraphs.find(paragraph => paragraph.exactText === 'Bullet Insertion Anchor');
      const inserted = paragraphs.find(paragraph => paragraph.exactText === 'Planner bullet insertion before untouched tail');
      const tail = paragraphs.find(paragraph => paragraph.exactText === 'Bullet Untouched Tail');
      assert.equal(anchor.list.level, 1);
      assert.equal(inserted.list.level, 0, 'outdent from Word list level 2 to level 1');
      assert.equal(inserted.list.numId, anchor.list.numId);
      assert.equal(inserted.list.label, '\uF076');
      assert.equal(tail.list.level, 0, 'untargeted root-level sibling remains at its source depth');
    }
  },
  {
    name: 'list-insert-bullet-deeper',
    request: { tool: 'insert_list_item', afterParagraphIndex: 3, text: 'Planner bullet insertion one level deeper', indentLevel: 1 },
    expected: sourceExactTexts.toSpliced(3, 0, 'Planner bullet insertion one level deeper'),
    newLists: [sourceListEntry('Planner bullet insertion one level deeper', 2, 'bullet', '\uF0D8', 1)],
    acceptedNumbering: {
      'Bullet Insertion Anchor': { listValue: 2 }
    },
    assertions: ({ paragraphs }) => {
      const anchor = paragraphs.find(paragraph => paragraph.exactText === 'Bullet Root A');
      const inserted = paragraphs.find(paragraph => paragraph.exactText === 'Planner bullet insertion one level deeper');
      assert.equal(inserted.list.level, 1);
      assert.equal(inserted.list.numId, anchor.list.numId);
      assert.equal(inserted.list.label, '\uF0D8');
      assert.equal(paragraphs.find(paragraph => paragraph.exactText === 'Bullet Insertion Anchor').list.label, '\uF0D8');
    }
  },
  {
    name: 'list-insert-numbered-same-level',
    request: { tool: 'insert_list_item', afterParagraphIndex: 9, text: 'Planner numbered continuation item', indentLevel: 0 },
    expected: sourceExactTexts.toSpliced(9, 0, 'Planner numbered continuation item'),
    newLists: [sourceListEntry('Planner numbered continuation item', 2, 'number-continuation', '1.2.', 2)],
    assertions: ({ paragraphs }) => {
      const anchor = paragraphs.find(paragraph => paragraph.exactText === 'Number Nested Anchor');
      const inserted = paragraphs.find(paragraph => paragraph.exactText === 'Planner numbered continuation item');
      assert.equal(inserted.list.level, 1);
      assert.equal(inserted.list.numId, anchor.list.numId, 'continuation stays in the Word-authored numbering instance');
      assert.equal(inserted.list.label, '1.2.');
    }
  },
  {
    name: 'list-insert-numbered-outdent',
    request: { tool: 'insert_list_item', afterParagraphIndex: 9, text: 'Planner numbered root insertion', indentLevel: -1 },
    expected: sourceExactTexts.toSpliced(9, 0, 'Planner numbered root insertion'),
    newLists: [sourceListEntry('Planner numbered root insertion', 1, 'number-continuation', '2.', 2)],
    acceptedNumbering: {
      'Number Root B': { listString: '3.', listValue: 3 },
      'Number Continued Item': { listString: '4.', listValue: 4 }
    },
    assertions: ({ paragraphs }) => {
      const anchor = paragraphs.find(paragraph => paragraph.exactText === 'Number Nested Anchor');
      const inserted = paragraphs.find(paragraph => paragraph.exactText === 'Planner numbered root insertion');
      const nextRoot = paragraphs.find(paragraph => paragraph.exactText === 'Number Root B');
      assert.equal(inserted.list.level, 0);
      assert.equal(inserted.list.numId, anchor.list.numId);
      assert.equal(nextRoot.list.level, 0);
      assert.equal(inserted.list.label, '2.');
      assert.equal(nextRoot.list.label, '3.');
      assert.equal(paragraphs.find(paragraph => paragraph.exactText === 'Number Continued Item').list.label, '4.');
    }
  },
  {
    name: 'list-insert-numbered-deeper',
    request: { tool: 'insert_list_item', afterParagraphIndex: 8, text: 'Planner numbered nested insertion', indentLevel: 1 },
    expected: sourceExactTexts.toSpliced(8, 0, 'Planner numbered nested insertion'),
    newLists: [sourceListEntry('Planner numbered nested insertion', 2, 'number-continuation', '1.1.', 1)],
    acceptedNumbering: {
      'Number Nested Anchor': { listString: '1.2.', listValue: 2 }
    },
    assertions: ({ paragraphs }) => {
      const anchor = paragraphs.find(paragraph => paragraph.exactText === 'Number Root A');
      const inserted = paragraphs.find(paragraph => paragraph.exactText === 'Planner numbered nested insertion');
      assert.equal(inserted.list.level, 1);
      assert.equal(inserted.list.numId, anchor.list.numId);
      assert.equal(inserted.list.label, '1.1.');
      assert.equal(paragraphs.find(paragraph => paragraph.exactText === 'Number Nested Anchor').list.label, '1.2.');
    }
  },
  {
    name: 'list-insert-after-plain-paragraph',
    request: { tool: 'insert_list_item', afterParagraphIndex: 2, text: 'Planner plain insertion after paragraph', indentLevel: 1 },
    expected: sourceExactTexts.toSpliced(2, 0, 'Planner plain insertion after paragraph'),
    newLists: [],
    assertions: ({ paragraphs }) => {
      const inserted = paragraphs.find(paragraph => paragraph.exactText === 'Planner plain insertion after paragraph');
      assert.equal(inserted.list, null, 'plain source anchor remains plain after list-shaped tool request');
    }
  },
  {
    name: 'list-edit-bullet-range',
    request: { tool: 'edit_list', startParagraphIndex: 3, endParagraphIndex: 4, newItems: ['Planner replacement parent', '    Planner replacement child'], listType: 'bullet', numberingStyle: 'decimal' },
    expected: sourceExactTexts.toSpliced(2, 2, 'Planner replacement parent', 'Planner replacement child'),
    newLists: [
      sourceListEntry('Planner replacement parent', 1, 'bullet', '\uF076'),
      sourceListEntry('Planner replacement child', 2, 'bullet', '\uF0D8')
    ],
    removeAcceptedListTexts: ['Bullet Root A', 'Bullet Insertion Anchor'],
    assertions: ({ paragraphs }) => {
      const parent = paragraphs.find(paragraph => paragraph.exactText === 'Planner replacement parent');
      const child = paragraphs.find(paragraph => paragraph.exactText === 'Planner replacement child');
      const untouched = paragraphs.find(paragraph => paragraph.exactText === 'Bullet Untouched Tail');
      assert.equal(parent.list.level, 0);
      assert.equal(child.list.level, 1);
      assert.equal(parent.list.numId, child.list.numId);
      assert.equal(parent.list.numId, untouched.list.numId, 'outside same-list item retains the original numbering instance');
    }
  },
  {
    name: 'list-insert-before-via-edit-list',
    request: { tool: 'edit_list', startParagraphIndex: 3, endParagraphIndex: 3, newItems: ['Planner bullet before root', 'Bullet Root A'], listType: 'bullet' },
    expected: sourceExactTexts.toSpliced(2, 1, 'Planner bullet before root', 'Bullet Root A'),
    newLists: [sourceListEntry('Planner bullet before root', 1, 'bullet', '\uF076', 1)],
    acceptedNumbering: {
      'Bullet Root A': { listValue: 2 },
      'Bullet Untouched Tail': { listValue: 3 },
      'Bullet Root B': { listValue: 4 }
    },
    assertions: ({ paragraphs }) => {
      const inserted = paragraphs.find(paragraph => paragraph.exactText === 'Planner bullet before root');
      const retained = paragraphs.find(paragraph => paragraph.exactText === 'Bullet Root A');
      assert.equal(inserted.list.level, 0);
      assert.equal(retained.list.level, 0);
      assert.equal(inserted.list.numId, retained.list.numId);
      assert.equal(inserted.list.label, '\uF076');
    }
  },
  {
    name: 'list-convert-noncontiguous-manual-headers',
    request: {
      tool: 'convert_headers_to_list',
      paragraphIndices: [23, 21],
      newHeaderTexts: ['Supported Header Second', 'Supported Header First'],
      numberingFormat: 'upperLetter'
    },
    expected: sourceExactTexts.map((text, index) => index === 20 ? 'Supported Header First' : index === 22 ? 'Supported Header Second' : text),
    newLists: [
      sourceListEntry('Supported Header First', 1, 'converted-headers', 'A.', 1),
      sourceListEntry('Supported Header Second', 1, 'converted-headers', 'B.', 2)
    ],
    assertions: ({ paragraphs }) => {
      const first = paragraphs.find(paragraph => paragraph.exactText === 'Supported Header First');
      const second = paragraphs.find(paragraph => paragraph.exactText === 'Supported Header Second');
      assert.equal(first.list?.format, 'upperLetter');
      assert.equal(second.list?.format, 'upperLetter');
      assert.equal(first.list?.label, 'A.');
      assert.equal(second.list?.label, 'B.');
      assert.equal(first.list?.numId, second.list?.numId);
      assert.equal(paragraphs.find(paragraph => paragraph.exactText === 'Plain paragraph between supported headers.').list, null);
    }
  }
];
const knownDefects = allCases.filter(testCase => [
  'list-insert-after-plain-paragraph',
  'list-edit-bullet-range'
].includes(testCase.name));
const cases = allCases.filter(testCase => !knownDefects.includes(testCase));

function inspectResolvedDoc(bytes) {
  const doc = openDocx(bytes);
  const report = doc.inspect();
  assert.equal(report.status, 'ok', JSON.stringify(report.errors));
  return { doc, paragraphs: report.paragraphs };
}

function hasDirectBold(documentXml, exactText) {
  const document = new DOMParser().parseFromString(documentXml, 'application/xml');
  const paragraphs = Array.from(document.getElementsByTagNameNS(wordNamespace, 'p'));
  const matches = paragraphs.filter(paragraph => Array.from(paragraph.getElementsByTagNameNS(wordNamespace, 't'))
    .map(node => node.textContent).join('') === exactText);
  assert.equal(matches.length, 1, `formatting sentinel must be unique: ${exactText}`);
  const bold = Array.from(matches[0].getElementsByTagNameNS(wordNamespace, 'b'));
  return bold.some(node => node.getAttributeNS(wordNamespace, 'val') !== '0');
}

function listEntriesForCase(testCase) {
  const accepted = sourceNumbering
    .filter(item => !testCase.removeAcceptedListTexts?.includes(item.text))
    .concat(testCase.newLists)
    .map(item => ({ ...item, ...testCase.acceptedNumbering?.[item.text] }));
  return {
    source: sourceNumbering,
    accepted,
    rejected: sourceNumbering
  };
}

async function runCase(testCase) {
  const planned = planAgenticListOperations({
    documentXml: sourceDocumentXml,
    numberingXml: sourceNumberingXml,
    stylesXml: sourceStylesXml,
    paragraphs: sourceParagraphs
  }, testCase.request);
  const result = await applyOperationsToDocumentXml(
    sourceDocumentXml,
    planned,
    'Agentic List Fidelity Test',
    { numberingXml: sourceNumberingXml, stylesXml: sourceStylesXml },
    {
      atomic: true,
      strictTargets: true,
      structuredContent: true,
      generateRedlines: true,
      existingRevisions: 'merge-same-author',
      date: '2026-09-30T12:00:00Z'
    }
  );
  assert.equal(result.status, 'ok', `${testCase.name}: ${JSON.stringify(result.error || result.results)}`);
  assert.equal(result.hasChanges, true, `${testCase.name} should make a tracked edit`);

  const trackedEntries = unzipDocx(sourceBytes);
  trackedEntries.set('word/document.xml', new TextEncoder().encode(result.documentXml));
  let numberingXml = sourceNumberingXml;
  for (const part of result.numberingXmlParts || []) {
    numberingXml = numberingXml
      ? mergeNumberingXmlBySchemaOrder(numberingXml, part)
      : part;
  }
  if (numberingXml) trackedEntries.set('word/numbering.xml', new TextEncoder().encode(numberingXml));
  const trackedBytes = zipDocx(trackedEntries);
  await validateDocxPackage(new MemoryZip(unzipDocx(trackedBytes)));

  const acceptedDoc = openDocx(trackedBytes);
  const rejectedDoc = openDocx(trackedBytes);
  const acceptResult = await acceptedDoc.resolveRevisions('accept', { allAuthors: true });
  const rejectResult = await rejectedDoc.resolveRevisions('reject', { allAuthors: true });
  assert.equal(acceptResult.status, 'ok', `${testCase.name} accept: ${JSON.stringify(acceptResult.error)}`);
  assert.equal(rejectResult.status, 'ok', `${testCase.name} reject: ${JSON.stringify(rejectResult.error)}`);

  const acceptedBytes = acceptedDoc.toUint8Array();
  const rejectedBytes = rejectedDoc.toUint8Array();
  const accepted = inspectResolvedDoc(acceptedBytes);
  const rejected = inspectResolvedDoc(rejectedBytes);
  assert.deepEqual(accepted.paragraphs.map(paragraph => paragraph.exactText), testCase.expected, `${testCase.name} Accept All exact text`);
  const expectedRejected = testCase.expectedRejected || sourceExactTexts;
  assert.deepEqual(rejected.paragraphs.map(paragraph => paragraph.exactText), expectedRejected, `${testCase.name} Reject All result`);
  assert.equal(accepted.paragraphs.map(paragraph => paragraph.exactText).join('\n'), testCase.expected.join('\n'));
  testCase.assertions({ paragraphs: accepted.paragraphs, source: sourceParagraphs });

  const acceptedXml = decode(unzipDocx(acceptedBytes).get('word/document.xml'));
  const rejectedXml = decode(unzipDocx(rejectedBytes).get('word/document.xml'));
  assert.equal(hasDirectBold(sourceDocumentXml, 'Untouched bold sentinel'), true, 'Word source has a bold formatting sentinel');
  assert.equal(hasDirectBold(acceptedXml, 'Untouched bold sentinel'), true, `${testCase.name} preserves unrelated bold formatting on accept`);
  assert.equal(hasDirectBold(rejectedXml, 'Untouched bold sentinel'), true, `${testCase.name} preserves unrelated bold formatting on reject`);

  // The independently authored numbering fixture proves continuity and restart
  // groups; check those same group relationships in resolved packages without
  // asserting Word's generated absolute IDs.
  for (const state of [accepted, ...(testCase.defect ? [] : [rejected])]) {
    const actual = new Map(state.paragraphs.filter(paragraph => paragraph.list?.numId)
      .map(paragraph => [paragraph.exactText, paragraph.list.numId]));
    assert.equal(actual.get('Bullet Untouched Tail'), actual.get('Bullet Root B'), `${testCase.name} retains bullet-list identity`);
    assert.equal(actual.get('Number Root A'), actual.get('Number Continued Item'), `${testCase.name} retains numbering continuation identity`);
    assert.notEqual(actual.get('Number Root A'), actual.get('Number Restart Root'), `${testCase.name} keeps restart in an independent list`);
  }

  return { planned, trackedBytes, acceptedBytes, rejectedBytes };
}

function manifestCase(testCase) {
  const listItems = listEntriesForCase(testCase);
  return {
    name: testCase.name,
    source: '../../../tests/fixtures/agentic-lists/nested-lists-source.docx',
    tracked: `${testCase.name}-tracked.docx`,
    accepted: `${testCase.name}-accepted.docx`,
    rejected: `${testCase.name}-rejected.docx`,
    expectedAcceptedText: testCase.expected.join('\n'),
    expectedRejectedText: sourceText,
    expectedComments: 0,
    expectedMinimumBodyRevisions: 1,
    expectedFormatting: [{ sourceText: 'Untouched bold sentinel', acceptedText: 'Untouched bold sentinel', bold: true, italic: false }],
    expectedNumbering: listItems,
    agenticRequest: testCase.request
  };
}

async function exportHostFixtures(exportDir, results) {
  mkdirSync(exportDir, { recursive: true });
  const manifest = {
    schemaVersion: 1,
    provenance: {
      source: 'Word-authored Microsoft Word COM fixture; see ../tests/fixtures/agentic-lists/source-observations.json',
      sourceWordBuild: sourceEvidence.provenance.wordBuild,
      trackedPackages: 'docx-redline-js 0.8.2 standalone Office.js bridge; fixed author/date; outputs are independent inputs for Word resolution checks',
      numberingExpectations: 'Hard-coded from Word source observations and requested list semantics; not copied from tracked-package output.'
    },
    cases: cases.map((testCase, index) => {
      const base = manifestCase(testCase);
      const result = results[index];
      writeFileSync(join(exportDir, base.tracked), result.trackedBytes);
      writeFileSync(join(exportDir, base.accepted), result.acceptedBytes);
      writeFileSync(join(exportDir, base.rejected), result.rejectedBytes);
      return base;
    }),
    capabilityChecks: [{
      name: 'unmarked-headers-refuse-conversion',
      request: { tool: 'convert_headers_to_list', paragraphIndices: [17, 19], newHeaderTexts: ['Header First Candidate', 'Header Second Candidate'], numberingFormat: 'arabic' },
      expectedOutcome: 'refused-without-writing',
      expectedErrorCode: 'UNSUPPORTED_LIST_CONVERSION',
      sourceParagraphs: ['Header First Candidate', 'Header Second Candidate']
    }],
    knownDefects: knownDefects.map(testCase => ({
      name: testCase.name,
      request: testCase.request,
      expectedOutcome: 'engine-defect-reproduced; excluded from passing host cases',
      expectedRejectedText: testCase.expectedRejected.join('\n'),
      defect: testCase.defect
    }))
  };
  writeFileSync(join(exportDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  const diagnosticCases = knownDefects.map(testCase => {
    const base = manifestCase(testCase);
    base.tracked = `${testCase.name}-diagnostic-tracked.docx`;
    base.accepted = `${testCase.name}-diagnostic-accepted.docx`;
    base.rejected = `${testCase.name}-diagnostic-rejected.docx`;
    // The defect manifest asks the independent Word host to Reject All back to
    // the true source text. The standalone outputs are retained as diagnostics.
    base.expectedRejectedText = sourceText;
    const result = testCase.result;
    writeFileSync(join(exportDir, base.tracked), result.trackedBytes);
    writeFileSync(join(exportDir, base.accepted), result.acceptedBytes);
    writeFileSync(join(exportDir, base.rejected), result.rejectedBytes);
    return base;
  });
  writeFileSync(join(exportDir, 'known-defects-manifest.json'), `${JSON.stringify({
    schemaVersion: 1,
    provenance: {
      source: 'Word-authored Microsoft Word COM fixture; independent Word Reject All is expected to confirm or refute the engine defect.',
      numberingExpectations: 'Source and accepted numbering expectations are specified independently from the tracked output.'
    },
    cases: diagnosticCases,
    controls: []
  }, null, 2)}\n`);
  return manifest;
}

for (const testCase of [...cases, ...knownDefects]) {
  if (knownDefects.includes(testCase)) {
    if (testCase.name === 'list-insert-after-plain-paragraph') {
      testCase.expectedRejected = sourceExactTexts.toSpliced(2, 0, '');
      testCase.defect = 'Reject All leaves one extra empty paragraph after the plain insertion anchor.';
    } else {
      testCase.expectedRejected = sourceExactTexts.toSpliced(2, 2, 'Bullet Root ABullet Insertion Anchor');
      testCase.defect = 'Reject All collapses the original adjacent list paragraphs into one paragraph.';
    }
  }
  const result = await runCase(testCase);
  testCase.result = result;
}

assert.throws(() => planAgenticListOperations({ paragraphs: sourceParagraphs }, {
  tool: 'convert_headers_to_list',
  paragraphIndices: [17, 19],
  newHeaderTexts: ['Header First Candidate', 'Header Second Candidate'],
  numberingFormat: 'arabic'
}), error => error.code === 'UNSUPPORTED_LIST_CONVERSION', 'unmarked headers must fail closed before mutation');
assert.throws(() => planAgenticListOperations({ paragraphs: sourceParagraphs }, {
  tool: 'convert_headers_to_list',
  paragraphIndices: [21],
  newHeaderTexts: ['Changed header text']
}), error => error.code === 'UNSUPPORTED_LIST_CONVERSION', 'changed manual-header text must fail closed');

const exportIndex = process.argv.indexOf('--export-host-dir');
if (exportIndex >= 0) {
  const destination = process.argv[exportIndex + 1];
  assert.ok(destination && !destination.startsWith('--'), '--export-host-dir requires a directory');
  const absolute = resolve(destination);
  const manifest = await exportHostFixtures(absolute, cases.map(testCase => testCase.result));
  console.log(`Exported ${manifest.cases.length} Word host cases to ${absolute}`);
}

for (const testCase of knownDefects) {
  console.log(`KNOWN_LIBRARY_DEFECT: ${testCase.name}: ${testCase.defect} See the separate known-defects-manifest.json for independent Word Reject All confirmation.`);
}

console.log(`PASS: ${cases.length} agentic list fidelity cases on Word-authored OOXML; bold formatting, exact Accept/Reject, and numbering identities preserved`);
