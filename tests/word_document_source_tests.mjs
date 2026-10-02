import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { inspectWordSource } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import {
  createWordDocumentSourceReader
} from '../src/taskpane/modules/chat/word-document-source.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REVISION = 'w:id="1" w:author="Review Bot" w:date="2026-10-02T00:00:00Z"';
const bodyXml = `
  <w:p>
    <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
    <w:commentRangeStart w:id="4"/>
    <w:r><w:t>Prior retained </w:t></w:r>
    <w:del ${REVISION}><w:r><w:delText>deleted wording </w:delText></w:r></w:del>
    <w:ins ${REVISION}><w:r><w:t>accepted wording</w:t></w:r></w:ins>
    <w:r><w:t> remains.</w:t></w:r>
    <w:commentRangeEnd w:id="4"/>
    <w:r><w:commentReference w:id="4"/></w:r>
  </w:p>
  <w:tbl>
    <w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
    <w:tr>
      <w:tc><w:p><w:r><w:t>Party A</w:t></w:r></w:p></w:tc>
      <w:tc><w:p><w:r><w:t>Party B</w:t></w:r></w:p></w:tc>
    </w:tr>
    <w:tr>
      <w:tc><w:p><w:pPr><w:pStyle w:val="Normal"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="77"/></w:numPr></w:pPr><w:r><w:t>Listed cell</w:t></w:r></w:p></w:tc>
      <w:tc><w:p><w:r><w:t>Second row, second cell</w:t></w:r></w:p></w:tc>
    </w:tr>
  </w:tbl>`;
const numberingXml = `<w:numbering xmlns:w="${W}">
  <w:abstractNum w:abstractNumId="12"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
  <w:num w:numId="77"><w:abstractNumId w:val="12"/></w:num>
</w:numbering>`;
const commentsPart = `<pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">
  <pkg:xmlData><w:comments xmlns:w="${W}"><w:comment w:id="4" w:author="Reviewer" w:date="2026-10-02T00:00:00Z"><w:p><w:r><w:t>Confirm this party name.</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData>
</pkg:part>`;
const basePackage = buildDocumentFragmentPackage(bodyXml, { numberingXml, appendTrailingParagraph: false });
const sourceOoxml = basePackage.replace('</pkg:package>', `${commentsPart}</pkg:package>`);

let supportLoads = 0;
let integrationInspections = 0;
const support = {
  inspectWordSource(ooxml) {
    integrationInspections++;
    return inspectWordSource(ooxml);
  }
};
let runCount = 0;
let sourceReadCount = 0;
let selectionReadCount = 0;
let selectionLoadCount = 0;
let syncCount = 0;
const selection = {
  text: 'Party A',
  load(property) {
    selectionLoadCount++;
    assert.equal(property, 'text', 'the source reader only requests the selected text it consumes');
  }
};
const run = async callback => {
  runCount++;
  await callback({
    document: {
      body: {
        getOoxml() {
          sourceReadCount++;
          return { value: sourceOoxml };
        }
      },
      getSelection() {
        selectionReadCount++;
        return selection;
      }
    },
    async sync() { syncCount++; }
  });
};
const reader = createWordDocumentSourceReader({
  loadSupport: async () => { supportLoads++; return [support]; },
  getDocumentId: () => 'word-document:one',
  run
});
const source = await reader();
const expected = inspectWordSource(sourceOoxml);

assert.deepEqual(source.parts, expected.parts, 'the host source must retain the library integration parts');
assert.deepEqual(source.inspection, expected.inspection, 'the reader must return the real library inspection without reinterpreting it');
assert.equal(source.documentId, 'word-document:one');
assert.equal(source.revisionView, 'accepted');
assert.equal(source.selectionText, 'Party A');
assert.equal(source.sourceOoxml, sourceOoxml);
assert.equal(supportLoads, 1);
assert.equal(integrationInspections, 1);
assert.equal(runCount, 1);
assert.equal(sourceReadCount, 1);
assert.equal(selectionReadCount, 1);
assert.equal(selectionLoadCount, 1);
assert.equal(syncCount, 1, 'a complete inspection uses one Office.js sync and no paragraph metadata round trips');

const acceptedParagraph = source.inspection.paragraphs[0];
assert.equal(acceptedParagraph.exactText, 'Prior retained accepted wording remains.');
assert.doesNotMatch(acceptedParagraph.exactText, /deleted wording/);
assert.equal(acceptedParagraph.hasRevisions, true);
assert.deepEqual(acceptedParagraph.revisionAuthors, ['Review Bot']);

const listedCell = source.inspection.paragraphs.find(paragraph => paragraph.exactText === 'Listed cell');
assert.ok(listedCell, 'the library inspection must include table-cell paragraphs');
assert.deepEqual(listedCell.table, { tableIndex: 1, rowIndex: 2, cellIndex: 1 });
assert.equal(listedCell.list.numId, '77');
assert.equal(listedCell.list.level, 0);
assert.equal(listedCell.list.format, 'upperLetter');
assert.equal(listedCell.list.label, 'C.');

assert.deepEqual(source.parts.numberingXml, expected.parts.numberingXml);
assert.deepEqual(source.parts.commentsXml, expected.parts.commentsXml);
const comment = source.inspection.comments.find(item => item.id === '4');
assert.ok(comment);
assert.equal(comment.author, 'Reviewer');
assert.equal(comment.text, 'Confirm this party name.');
assert.equal(comment.paragraphIndex, 1);
assert.equal(comment.anchoredText, 'Prior retained accepted wording remains.');
assert.deepEqual(acceptedParagraph.commentIds, ['4']);

// Cancellation before the host call does not even load the integration bridge.
const alreadyCancelled = new AbortController();
alreadyCancelled.abort();
let cancelledSupportLoads = 0;
let cancelledRunCount = 0;
const cancelledReader = createWordDocumentSourceReader({
  loadSupport: async () => { cancelledSupportLoads++; return [support]; },
  getDocumentId: () => 'word-document:one',
  signal: alreadyCancelled.signal,
  run: async () => { cancelledRunCount++; }
});
await assert.rejects(cancelledReader(), error => error.name === 'AbortError');
assert.equal(cancelledSupportLoads, 0);
assert.equal(cancelledRunCount, 0);

// Cancellation while the single host sync is settling is surfaced before any
// canonical inspection result can be used by a caller.
const cancelledDuringRead = new AbortController();
let cancelledDuringInspections = 0;
const cancellationReader = createWordDocumentSourceReader({
  loadSupport: async () => [{
    inspectWordSource(ooxml) { cancelledDuringInspections++; return inspectWordSource(ooxml); }
  }],
  getDocumentId: () => 'word-document:one',
  signal: cancelledDuringRead.signal,
  run: async callback => callback({
    document: {
      body: { getOoxml: () => ({ value: sourceOoxml }) },
      getSelection: () => selection
    },
    async sync() { cancelledDuringRead.abort(); }
  })
});
await assert.rejects(cancellationReader(), error => error.name === 'AbortError');
assert.equal(cancelledDuringInspections, 0);

// A document switch during the host read invalidates the snapshot before the
// library's result can escape as the active document's source.
let activeDocumentId = 'word-document:one';
let switchedDocumentInspections = 0;
const switchedDocumentReader = createWordDocumentSourceReader({
  loadSupport: async () => [{
    inspectWordSource(ooxml) { switchedDocumentInspections++; return inspectWordSource(ooxml); }
  }],
  getDocumentId: () => activeDocumentId,
  run: async callback => callback({
    document: {
      body: { getOoxml: () => ({ value: sourceOoxml }) },
      getSelection: () => selection
    },
    async sync() { activeDocumentId = 'word-document:two'; }
  })
});
await assert.rejects(switchedDocumentReader(), error => error.code === 'DOCUMENT_CHANGED');
assert.equal(switchedDocumentInspections, 0);

console.log('PASS: Word source capture uses one host sync and library-canonical accepted text, table/list metadata and comments');
