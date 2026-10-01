import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  inspectDocumentParts,
  openDocx,
  rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';
import { unzipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';

const fixturePath = fileURLToPath(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url));
const sourceBytes = new Uint8Array(readFileSync(fixturePath));
const decoder = new TextDecoder();
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function emptyListDocumentXml(count) {
  const paragraphs = Array.from({ length: count }, () =>
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>'
  ).join('');
  return `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Intro</w:t></w:r></w:p>${paragraphs}<w:sectPr/></w:body></w:document>`;
}

async function testEmptyOnlyListRangeRejectsWithEverySourceParagraph() {
  for (const scenario of [
    { label: 'single empty source paragraph', sourceEmptyCount: 1, target: 2, targetEnd: undefined, replacement: '- One' },
    { label: 'all-empty source range', sourceEmptyCount: 2, target: 2, targetEnd: 3, replacement: '- One\n- Two' }
  ]) {
    const sourceXml = emptyListDocumentXml(scenario.sourceEmptyCount);
    const operation = {
      type: 'redline',
      target: { index: scenario.target, exactText: '' },
      ...(scenario.targetEnd ? { targetEnd: { index: scenario.targetEnd, exactText: '' } } : {}),
      modified: scenario.replacement
    };
    const result = await applyOperationsToDocumentXml(
      sourceXml,
      [operation],
      '0.8.3 Empty List Compatibility Test',
      {},
      {
        atomic: true,
        strictTargets: true,
        structuredContent: true,
        generateRedlines: true,
        date: '2026-09-30T12:00:00Z'
      }
    );
    assert.equal(result.status, 'ok', `${scenario.label}: ${JSON.stringify(result.error || result.results)}`);

    const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true });
    assert.equal(rejected.status, undefined, `${scenario.label}: ${JSON.stringify(rejected.error)}`);
    const inspected = inspectDocumentParts({ documentXml: rejected.oxml });
    assert.equal(inspected.status, 'ok', `${scenario.label}: ${JSON.stringify(inspected.errors)}`);
    assert.deepEqual(inspected.paragraphs.map(paragraph => paragraph.exactText), [
      'Intro', ...Array(scenario.sourceEmptyCount).fill('')
    ], `${scenario.label}: Reject All restores each empty source paragraph boundary`);
  }
}

function part(bytes, name) {
  const data = unzipDocx(bytes).get(name);
  assert.ok(data, `DOCX package is missing ${name}`);
  return data;
}

async function resolvePackage(bytes, resolution) {
  const doc = openDocx(bytes);
  const result = await doc.resolveRevisions(resolution, { allAuthors: true });
  assert.equal(result.status, 'ok', `${resolution}: ${JSON.stringify(result.error)}`);
  return doc;
}

async function testPublicFacadePreservesSameKindListNumbering() {
  const sourceDoc = openDocx(sourceBytes);
  const sourceReport = sourceDoc.inspect();
  assert.equal(sourceReport.status, 'ok');
  const sourceParagraphs = sourceReport.paragraphs;
  const sourceNumbering = new Uint8Array(part(sourceBytes, 'word/numbering.xml'));
  const sourceListItems = [
    { text: 'Bullet Root A', numId: sourceParagraphs.find(p => p.exactText === 'Bullet Root A')?.list?.numId },
    { text: 'Bullet Insertion Anchor', numId: sourceParagraphs.find(p => p.exactText === 'Bullet Insertion Anchor')?.list?.numId },
    { text: 'Bullet Untouched Tail', numId: sourceParagraphs.find(p => p.exactText === 'Bullet Untouched Tail')?.list?.numId }
  ];
  assert.ok(sourceListItems.every(item => item.numId), 'source fixture exposes the three related list bindings');
  assert.equal(new Set(sourceListItems.map(item => item.numId)).size, 1, 'the replaced range and untouched tail share one source numbering instance');

  const result = await sourceDoc.applyOperations([{
    type: 'redline',
    target: { index: 3, exactText: 'Bullet Root A' },
    targetEnd: { index: 4, exactText: 'Bullet Insertion Anchor' },
    modified: '- Planner replacement parent\n    - Planner replacement child'
  }], {
    author: '0.8.3 Public Facade Compatibility Test',
    atomic: true,
    strictTargets: true,
    structuredContent: true,
    generateRedlines: true,
    date: '2026-09-30T12:00:00Z'
  });
  assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
  const trackedBytes = sourceDoc.toUint8Array();
  assert.deepEqual(part(trackedBytes, 'word/numbering.xml'), sourceNumbering,
    'same-kind list replacement must not add or remap numbering definitions through openDocx');

  const accepted = await resolvePackage(trackedBytes, 'accept');
  const rejected = await resolvePackage(trackedBytes, 'reject');
  const acceptedParagraphs = accepted.inspect().paragraphs;
  const rejectedParagraphs = rejected.inspect().paragraphs;
  assert.deepEqual(acceptedParagraphs.map(paragraph => paragraph.exactText), sourceParagraphs.map(paragraph => paragraph.exactText)
    .toSpliced(2, 2, 'Planner replacement parent', 'Planner replacement child'));
  assert.deepEqual(rejectedParagraphs.map(paragraph => paragraph.exactText), sourceParagraphs.map(paragraph => paragraph.exactText),
    'Reject All restores the source text and both replaced paragraph boundaries');

  const acceptedByText = new Map(acceptedParagraphs.map(paragraph => [paragraph.exactText, paragraph]));
  assert.equal(acceptedByText.get('Planner replacement parent')?.list?.format, 'bullet');
  assert.equal(acceptedByText.get('Planner replacement child')?.list?.format, 'bullet');
  assert.equal(acceptedByText.get('Planner replacement parent')?.list?.numId, sourceListItems[0].numId,
    'replacement items reuse the source list numbering instance');
  assert.equal(acceptedByText.get('Planner replacement child')?.list?.numId, sourceListItems[0].numId);
  assert.equal(acceptedByText.get('Bullet Untouched Tail')?.list?.numId, sourceListItems[2].numId,
    'the untouched tail stays on the same list after the range replacement');

  const rejectedByText = new Map(rejectedParagraphs.map(paragraph => [paragraph.exactText, paragraph]));
  for (const sourceItem of sourceListItems) {
    assert.equal(rejectedByText.get(sourceItem.text)?.list?.format, 'bullet', `${sourceItem.text} restores as a bullet`);
    assert.equal(rejectedByText.get(sourceItem.text)?.list?.numId, sourceItem.numId, `${sourceItem.text} restores to its original list`);
  }
}

try {
  await testEmptyOnlyListRangeRejectsWithEverySourceParagraph();
  await testPublicFacadePreservesSameKindListNumbering();
  console.log('PASS: docx-redline-js 0.8.3 empty-range rejection and public facade list numbering compatibility tests');
} catch (error) {
  console.error('FAIL:', error?.stack || error?.message || error);
  process.exitCode = 1;
}
