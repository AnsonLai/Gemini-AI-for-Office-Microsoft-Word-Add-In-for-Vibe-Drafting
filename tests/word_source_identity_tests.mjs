import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
  computeDocumentPartsRevisionToken,
  inspectDocumentParts
} from '@ansonlai/docx-redline-js';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { inspectWordSource } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { createWordDocumentSourceReader } from '../src/taskpane/modules/chat/word-document-source.js';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';
import {
  computeWordSourceRevisionToken,
  detectExportTailIndex
} from '../src/taskpane/modules/docx-redline-js-integration/word-source-identity.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const tracked = 'w:ins w:id="71" w:author="Tracked Author" w:date="2026-10-02T00:00:00Z"';
const bodyXml = `
  <w:p xmlns:w14="${W14}" w14:paraId="1111AAAA" w:rsidR="00000001" w:rsidRDefault="00000001">
    <w:r><w:t>First real paragraph.</w:t></w:r>
  </w:p>
  <w:p xmlns:w14="${W14}" w14:paraId="2222BBBB" w:rsidR="00000002" w:rsidP="00000002">
    <${tracked}><w:r><w:t>Tracked accepted text.</w:t></w:r></w:ins>
  </w:p>
  <w:p xmlns:w14="${W14}" w14:paraId="3333CCCC" w:rsidR="00000003"/>
`;
const settingsA = `<w:settings xmlns:w="${W}"><w:rsids><w:rsidRoot w:val="AAAA0001"/><w:rsid w:val="00000001"/></w:rsids><w:doNotTrackMoves/></w:settings>`;
const settingsB = `<w:settings xmlns:w="${W}"><w:rsids><w:rsidRoot w:val="BBBB0002"/><w:rsid w:val="00000002"/><w:rsid w:val="00000003"/></w:rsids><w:doNotTrackMoves/></w:settings>`;

function fixtureParts(paragraphXml = bodyXml, settingsXml = settingsA) {
  const flatOpc = buildDocumentFragmentPackage(paragraphXml, { appendTrailingParagraph: false });
  const parts = inspectWordSource(flatOpc).parts;
  return {
    ...parts,
    additionalParts: { ...parts.additionalParts, '/word/settings.xml': settingsXml }
  };
}

function replaceOnce(source, before, after) {
  assert.ok(source.includes(before), `fixture contains ${before}`);
  return source.replace(before, after);
}

const originalParts = fixtureParts();
const originalPartsCopy = JSON.parse(JSON.stringify(originalParts));
const originalInspection = inspectDocumentParts(originalParts);
assert.equal(originalInspection.status, 'ok');
assert.equal(originalInspection.paragraphs.length, 3);
assert.equal(originalInspection.paragraphs[0].exactText, 'First real paragraph.');
assert.equal(originalInspection.paragraphs[1].exactText, 'Tracked accepted text.');
assert.equal(originalInspection.paragraphs[1].hasRevisions, true);
assert.equal(originalInspection.paragraphs[1].revisionAuthors[0], 'Tracked Author');
assert.equal(detectExportTailIndex(originalParts, { hostParagraphCount: 2 }), 3,
  'only the certified final empty body paragraph is identified as the export tail');
assert.equal(detectExportTailIndex(originalParts, { hostParagraphCount: 1 }), null,
  'a paragraph-count mismatch cannot certify the export tail');
assert.equal(detectExportTailIndex(originalParts, {}), null,
  'without the host count, no paragraph ID is treated as an export artifact');

// Word refreshes session rsids in document.xml and settings.xml. Those values
// are omitted only from the derived token input; real text, IDs, tracked-change
// identifiers and authors remain part of the source identity.
const volatileRsidParts = {
  ...originalParts,
  documentXml: originalParts.documentXml
    .replaceAll('00000001', '00ABCDEF')
    .replaceAll('00000002', '00FEDCBA')
    .replaceAll('00000003', '00123456'),
  additionalParts: { ...originalParts.additionalParts, '/word/settings.xml': settingsB }
};
const rawOriginalToken = await computeDocumentPartsRevisionToken(originalParts);
const rawVolatileToken = await computeDocumentPartsRevisionToken(volatileRsidParts);
assert.notEqual(rawOriginalToken.value, rawVolatileToken.value,
  'the generic public token is still exact over raw source parts');
const normalizedOriginalToken = await computeWordSourceRevisionToken(originalParts, { hostParagraphCount: 2 });
const normalizedVolatileToken = await computeWordSourceRevisionToken(volatileRsidParts, { hostParagraphCount: 2 });
assert.equal(normalizedOriginalToken.value, normalizedVolatileToken.value,
  'document rsid attributes and only the settings rsids list are normalized in the Word-derived token');
assert.deepEqual(originalParts, originalPartsCopy, 'computing the derived token never mutates original parts');
const comments = `<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="Reviewer"><w:p w:rsidR="00000001"><w:r><w:t>Retained comment.</w:t></w:r></w:p></w:comment></w:comments>`;
const commentsBefore = { ...originalParts, commentsXml: comments };
const commentsAfter = { ...originalParts, commentsXml: comments.replace('00000001', '99999999') };
assert.equal((await computeWordSourceRevisionToken(commentsBefore, { hostParagraphCount: 2 })).value,
  (await computeWordSourceRevisionToken(commentsAfter, { hostParagraphCount: 2 })).value, 'Word regenerates session attributes in comment paragraphs too');
assert.notEqual((await computeWordSourceRevisionToken({ ...commentsAfter, commentsXml: commentsAfter.commentsXml.replace('Reviewer', 'Other Reviewer') }, { hostParagraphCount: 2 })).value,
  (await computeWordSourceRevisionToken(commentsBefore, { hostParagraphCount: 2 })).value, 'comment authors remain meaningful source identity');
assert.match(originalParts.documentXml, /w14:paraId="1111AAAA"/,
  'real paragraph IDs remain in the original OOXML payload');
assert.match(originalParts.documentXml, /w:id="71" w:author="Tracked Author"/,
  'tracked insertion IDs and authors remain in the original OOXML payload');

const normalizedInspection = inspectDocumentParts(volatileRsidParts);
assert.deepEqual(
  normalizedInspection.paragraphs.slice(0, 2).map(({ paragraphId, fingerprint }) => ({ paragraphId, fingerprint })),
  originalInspection.paragraphs.slice(0, 2).map(({ paragraphId, fingerprint }) => ({ paragraphId, fingerprint })),
  'Word paragraph fingerprints use paragraph ID/index/table/view/text, not volatile rsid metadata'
);

// Only the one-based certified empty tail can have its w14:paraId excluded.
const changedExportTailId = {
  ...originalParts,
  documentXml: replaceOnce(originalParts.documentXml, '3333CCCC', '9999FFFF')
};
const tailToken = await computeWordSourceRevisionToken(originalParts, { hostParagraphCount: 2 });
assert.equal((await computeWordSourceRevisionToken(changedExportTailId, { hostParagraphCount: 2 })).value, tailToken.value,
  'the certified export-only tail ID can change without making canonical content stale');
assert.notEqual((await computeWordSourceRevisionToken(changedExportTailId)).value,
  (await computeWordSourceRevisionToken(originalParts)).value,
  'without a host paragraph count the trailing ID remains meaningful');
assert.notEqual((await computeWordSourceRevisionToken(changedExportTailId, { hostParagraphCount: 1 })).value,
  (await computeWordSourceRevisionToken(originalParts, { hostParagraphCount: 1 })).value,
  'a count mismatch does not normalize the trailing ID');

const changedRealParaId = {
  ...originalParts,
  documentXml: replaceOnce(originalParts.documentXml, '1111AAAA', 'AAAABBBB')
};
assert.notEqual((await computeWordSourceRevisionToken(changedRealParaId, { hostParagraphCount: 2 })).value,
  (await computeWordSourceRevisionToken(originalParts, { hostParagraphCount: 2 })).value,
  'a real target paragraph ID change remains detectable');
assert.notEqual(
  inspectDocumentParts(changedRealParaId).paragraphs[0].fingerprint,
  originalInspection.paragraphs[0].fingerprint,
  'the real paragraph fingerprint still includes its paragraph ID'
);

const changedInsertionId = {
  ...originalParts,
  documentXml: replaceOnce(originalParts.documentXml, 'w:id="71"', 'w:id="72"')
};
assert.notEqual((await computeWordSourceRevisionToken(changedInsertionId, { hostParagraphCount: 2 })).value,
  (await computeWordSourceRevisionToken(originalParts, { hostParagraphCount: 2 })).value,
  'a tracked insertion ID change remains detectable');
const changedInsertionAuthor = {
  ...originalParts,
  documentXml: replaceOnce(originalParts.documentXml, 'w:author="Tracked Author"', 'w:author="Other Author"')
};
assert.notEqual((await computeWordSourceRevisionToken(changedInsertionAuthor, { hostParagraphCount: 2 })).value,
  (await computeWordSourceRevisionToken(originalParts, { hostParagraphCount: 2 })).value,
  'a tracked insertion author change remains detectable');
const changedText = {
  ...originalParts,
  documentXml: replaceOnce(originalParts.documentXml, 'First real paragraph.', 'Changed real paragraph.')
};
assert.notEqual((await computeWordSourceRevisionToken(changedText, { hostParagraphCount: 2 })).value,
  (await computeWordSourceRevisionToken(originalParts, { hostParagraphCount: 2 })).value,
  'canonical text changes remain detectable');
const changedSettings = {
  ...originalParts,
  additionalParts: {
    ...originalParts.additionalParts,
    '/word/settings.xml': settingsA.replace('<w:doNotTrackMoves/>', '<w:doNotTrackMoves w:val="true"/>')
  }
};
assert.notEqual((await computeWordSourceRevisionToken(changedSettings, { hostParagraphCount: 2 })).value,
  (await computeWordSourceRevisionToken(originalParts, { hostParagraphCount: 2 })).value,
  'settings outside the w:rsids list remain detectable');

const nonEmptyTailParts = fixtureParts(bodyXml.replace('<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="3333CCCC" w:rsidR="00000003"/>',
  '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="3333CCCC"><w:r><w:t>Real final text.</w:t></w:r></w:p>'));
assert.equal(detectExportTailIndex(nonEmptyTailParts, { hostParagraphCount: 2 }), null,
  'a non-empty final paragraph is never discarded as an export artifact');

const tableTail = fixtureParts(`<w:p xmlns:w14="${W14}" w14:paraId="1111AAAA"><w:r><w:t>First real paragraph.</w:t></w:r></w:p>
  <w:p xmlns:w14="${W14}" w14:paraId="2222BBBB"><w:r><w:t>Second real paragraph.</w:t></w:r></w:p>
  <w:tbl><w:tr><w:tc><w:p xmlns:w14="${W14}" w14:paraId="3333CCCC"/></w:tc></w:tr></w:tbl>`);
assert.equal(detectExportTailIndex(tableTail, { hostParagraphCount: 2 }), null,
  'an empty table-cell paragraph is never mistaken for the body export tail');

// The reader queues the host paragraph count beside OOXML and selection in the
// same single sync. It remains optional for minimal Office.js mocks.
let sourceSyncCount = 0;
const paragraphCollection = { items: [], load: property => assert.equal(property, '$none') };
const reader = createWordDocumentSourceReader({
  loadSupport: async () => [{ inspectWordSource }],
  getDocumentId: () => 'source-identity-document',
  run: async callback => callback({
    document: {
      body: {
        getOoxml: () => ({ value: buildDocumentFragmentPackage(bodyXml, { appendTrailingParagraph: false }) }),
        paragraphs: paragraphCollection
      },
      getSelection: () => ({ text: 'First real paragraph.', load: property => assert.equal(property, 'text') })
    },
    async sync() { sourceSyncCount++; paragraphCollection.items = [{}, {}]; }
  })
});
const read = await reader();
assert.equal(read.hostParagraphCount, 2);
assert.equal(sourceSyncCount, 1);
assert.equal(detectExportTailIndex(read.parts, { hostParagraphCount: read.hostParagraphCount }), 3);

let optionalSyncCount = 0;
const optionalReader = createWordDocumentSourceReader({
  loadSupport: async () => [{ inspectWordSource }],
  getDocumentId: () => 'source-identity-document',
  run: async callback => callback({
    document: {
      body: { getOoxml: () => ({ value: buildDocumentFragmentPackage(bodyXml, { appendTrailingParagraph: false }) }) },
      getSelection: () => ({ text: '', load() {} })
    },
    async sync() { optionalSyncCount++; }
  })
});
assert.equal((await optionalReader()).hostParagraphCount, undefined);
assert.equal(optionalSyncCount, 1);

console.log('PASS: Word source identity ignores only verified host-volatiles and certified export-tail paragraph ID');

let currentParts = originalParts;
const session = createDocumentInspectionSession({
  readSource: async () => ({ documentId: 'word-source-test', parts: currentParts, hostParagraphCount: 2 }),
  computeRevisionToken: computeWordSourceRevisionToken,
  getExcludedParagraphIndex: detectExportTailIndex,
});
const inspected = await session.inspect({ includeDocumentEnd: true });
assert.equal(inspected.targets.length, 2, 'the certified export artifact never receives an editable paragraph handle');
currentParts = { ...volatileRsidParts, documentXml: volatileRsidParts.documentXml.replace('3333CCCC', '9999FFFF') };
const resolved = await session.resolveHandles([inspected.targets[0].handle, inspected.documentEnd.handle]);
assert.equal(resolved.status, 'ok', JSON.stringify(resolved.error));
assert.equal(resolved.targets[1].descriptor.paragraphId, '2222BBBB', 'append targets the last real paragraph rather than the export artifact');
assert.equal(resolved.targets[1].descriptor.exactText, 'Tracked accepted text.');
currentParts = changedRealParaId;
assert.equal((await session.resolveHandles([inspected.targets[0].handle])).status, 'error', 'a meaningful paragraph identity change still expires real handles');
