import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
  computeDocumentPartsRevisionToken,
  inspectDocumentParts
} from '@ansonlai/docx-redline-js';
import { buildDocumentCommentsPackage, buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { inspectWordSource } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { createWordDocumentSourceReader } from '../src/taskpane/modules/chat/word-document-source.js';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';
import {
  computeWordSourceRevisionToken,
  computeWordRestorationToken,
  computeWordRevertGuardToken,
  detectExportTailIndex
} from '../src/taskpane/modules/docx-redline-js-integration/word-source-identity.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
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

function addFlatOpcPart(flatOpc, name, contentType, xml) {
  const anchor = '<pkg:part pkg:name="/word/document.xml"';
  assert.ok(flatOpc.includes(anchor), 'Flat OPC fixture has its main document part');
  const part = `<pkg:part pkg:name="/word/${name}" pkg:contentType="${contentType}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
  return flatOpc.replace(anchor, `${part}\n  ${anchor}`);
}

function buildThreadedCommentFixture() {
  const documentXml = `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>
    <w:p w14:paraId="10000001"><w:commentRangeStart w:id="1"/><w:r><w:t>Root anchor text.</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>
    <w:p w14:paraId="10000002"><w:commentRangeStart w:id="2"/><w:r><w:t>Reply anchor text.</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>
    <w:p w14:paraId="10000003"><w:commentRangeStart w:id="3"/><w:r><w:t>Other root anchor.</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p>
  </w:body></w:document>`;
  const commentsXml = `<w:comments xmlns:w="${W}" xmlns:w14="${W14}">
    <w:comment w:id="1" w:author="Reviewer" w:date="2026-10-02T00:00:00Z"><w:p w14:paraId="AAAA0001"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:t>Root note.</w:t></w:r></w:p></w:comment>
    <w:comment w:id="2" w:author="Editor" w:date="2026-10-02T00:01:00Z"><w:p w14:paraId="CCCC0003"><w:r><w:t>Reply note.</w:t></w:r></w:p></w:comment>
    <w:comment w:id="3" w:author="Reviewer" w:date="2026-10-02T00:02:00Z"><w:p w14:paraId="BBBB0002"><w:r><w:t>Other root note.</w:t></w:r></w:p></w:comment>
  </w:comments>`;
  const commentsExtendedXml = `<w15:commentsEx xmlns:w15="${W15}">
    <w15:commentEx w15:paraId="AAAA0001" w15:done="0"/>
    <w15:commentEx w15:paraId="CCCC0003" w15:paraIdParent="AAAA0001" w15:done="0"/>
    <w15:commentEx w15:paraId="BBBB0002" w15:done="0"/>
  </w15:commentsEx>`;
  const commentsIdsXml = `<w16cid:commentsIds xmlns:w16cid="${W16CID}">
    <w16cid:commentId w16cid:paraId="AAAA0001" w16cid:durableId="11111111"/>
    <w16cid:commentId w16cid:paraId="CCCC0003" w16cid:durableId="22222222"/>
    <w16cid:commentId w16cid:paraId="BBBB0002" w16cid:durableId="33333333"/>
  </w16cid:commentsIds>`;
  let flatOpc = buildDocumentCommentsPackage(documentXml, commentsXml);
  flatOpc = addFlatOpcPart(flatOpc, 'commentsExtended.xml', 'application/xml', commentsExtendedXml);
  flatOpc = addFlatOpcPart(flatOpc, 'commentsIds.xml', 'application/xml', commentsIdsXml);
  return { flatOpc, parts: inspectWordSource(flatOpc).parts, commentsXml, commentsExtendedXml, commentsIdsXml };
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
const restoredOriginalToken = await computeWordRestorationToken(originalParts, { hostParagraphCount: 2 });
assert.equal((await computeWordRestorationToken(changedRealParaId, { hostParagraphCount: 2 })).value,
  restoredOriginalToken.value, 'restore verification tolerates only Word-regenerated body paragraph IDs');
const changedTextId = { ...originalParts, documentXml: originalParts.documentXml.replace('w14:paraId="1111AAAA"', 'w14:paraId="1111AAAA" w14:textId="77777777"') };
assert.equal((await computeWordRestorationToken(changedTextId, { hostParagraphCount: 2 })).value, restoredOriginalToken.value);
assert.notEqual((await computeWordSourceRevisionToken(changedTextId, { hostParagraphCount: 2 })).value, normalizedOriginalToken.value,
  'normal freshness keeps real paragraph text IDs');
for (const xml of [originalParts.documentXml.replace('First real paragraph.', 'Different text.'),
  originalParts.documentXml.replace('<w:t>First real paragraph.', '<w:rPr><w:b/></w:rPr><w:t>First real paragraph.')]) {
  assert.notEqual((await computeWordRestorationToken({ ...originalParts, documentXml: xml })).value, restoredOriginalToken.value,
    'restore verification retains literal content and inline formatting');
}
assert.deepEqual(originalParts, originalPartsCopy, 'restoration comparison never edits the saved source');
for (const changed of [changedInsertionId, changedInsertionAuthor]) {
  assert.notEqual((await computeWordRestorationToken(changed, { hostParagraphCount: 2 })).value,
    restoredOriginalToken.value, 'restoration must preserve tracked-change IDs and authors');
}
const restoredCommentsToken = await computeWordRestorationToken(commentsBefore);
for (const changedComments of [comments.replace('Reviewer', 'Other Reviewer'), comments.replace('Retained comment.', 'Different comment.'),
  comments.replace('w:id="1"', 'w:id="2"')]) {
  assert.notEqual((await computeWordRestorationToken({ ...commentsBefore, commentsXml: changedComments })).value,
    restoredCommentsToken.value, 'comment author, text and anchor identity cannot be normalized during restore');
}

// A body Replace can remap Word-generated paragraph IDs in comments and their
// w15/w16cid links. Restoration tolerates that remapping only when the
// canonical comment and reply relationships stay identical.
const threaded = buildThreadedCommentFixture();
const threadedPartsCopy = JSON.parse(JSON.stringify(threaded.parts));
const threadedFlatCopy = threaded.flatOpc;
const threadedRestoreToken = await computeWordRestorationToken(threaded.parts);
assert.deepEqual(threadedRestoreToken.commentIdentity, [
  { id: '1', author: 'Reviewer', date: '2026-10-02T00:00:00Z', text: 'Root note.', done: false,
    paragraphIndex: 1, targetRef: 'P1', anchoredText: 'Root anchor text.' },
  { id: '2', author: 'Editor', date: '2026-10-02T00:01:00Z', text: 'Reply note.', done: false,
    parentCommentId: '1', paragraphIndex: 2, targetRef: 'P2', anchoredText: 'Reply anchor text.' },
  { id: '3', author: 'Reviewer', date: '2026-10-02T00:02:00Z', text: 'Other root note.', done: false,
    paragraphIndex: 3, targetRef: 'P3', anchoredText: 'Other root anchor.' }
], 'restoration projects durable comment identity and anchor fields without volatile paragraph IDs');

let remappedFlatOpc = threaded.flatOpc;
for (const [before, after] of [
  ['10000001', '10000011'], ['10000002', '10000012'], ['10000003', '10000013'],
  ['AAAA0001', 'AAAA0011'], ['CCCC0003', 'CCCC0013'], ['BBBB0002', 'BBBB0012']
]) remappedFlatOpc = remappedFlatOpc.replaceAll(before, after);
const remappedThreadParts = inspectWordSource(remappedFlatOpc).parts;
const remappedRestoreToken = await computeWordRestorationToken(remappedThreadParts);
assert.equal(remappedRestoreToken.value, threadedRestoreToken.value,
  'Word-regenerated body/comment paragraph IDs and linked w15/w16cid paraIds are omitted during restore comparison');
assert.deepEqual(remappedRestoreToken.commentIdentity, threadedRestoreToken.commentIdentity,
  'the remapped reply still identifies the same root comment and unchanged anchor');
assert.notEqual((await computeWordSourceRevisionToken(remappedThreadParts)).value,
  (await computeWordSourceRevisionToken(threaded.parts)).value,
  'ordinary source freshness tokens remain strict about all real paragraph IDs');

const reparentedThreadParts = {
  ...threaded.parts,
  commentsExtendedXml: threaded.parts.commentsExtendedXml
    .replace('w15:paraIdParent="AAAA0001"', 'w15:paraIdParent="BBBB0002"')
};
const reparentedRestoreToken = await computeWordRestorationToken(reparentedThreadParts);
assert.equal(reparentedRestoreToken.value, threadedRestoreToken.value,
  'comment paragraph-link normalization leaves the derived parts token unchanged');
assert.notDeepEqual(reparentedRestoreToken.commentIdentity, threadedRestoreToken.commentIdentity,
  'canonical comment projection detects that the reply now belongs to a different root');
assert.equal(reparentedRestoreToken.commentIdentity.find(comment => comment.id === '2').parentCommentId, '3');

const reanchoredDocument = threaded.parts.documentXml
  .replaceAll('w:id="1"', 'w:id="TEMP"')
  .replaceAll('w:id="2"', 'w:id="1"')
  .replaceAll('w:id="TEMP"', 'w:id="2"');
const changedDurableIdsXml = threaded.parts.commentsIdsXml
  .replace('w16cid:durableId="11111111"', 'w16cid:durableId="99999999"');
const lostCommentParts = {
  ...threaded.parts,
  commentsXml: threaded.parts.commentsXml.replace(/\s*<w:comment w:id="2"[^>]*>[\s\S]*?<\/w:comment>/, '')
};
await assert.rejects(() => computeWordRestorationToken(lostCommentParts), /metadata cannot be resolved/,
  'lost comment metadata fails closed rather than being normalized away');

const restorationMutations = [
  ['re-anchored comment', { ...threaded.parts, documentXml: reanchoredDocument }],
  ['changed resolution state', {
    ...threaded.parts,
    commentsExtendedXml: threaded.parts.commentsExtendedXml.replace('w15:done="0"', 'w15:done="1"')
  }],
  ['changed durable comment ID', {
    ...threaded.parts,
    commentsIdsXml: changedDurableIdsXml,
    additionalParts: { ...threaded.parts.additionalParts, '/word/commentsIds.xml': changedDurableIdsXml }
  }],
  ['changed comment text', { ...threaded.parts, commentsXml: threaded.parts.commentsXml.replace('Root note.', 'Updated root note.') }],
  ['changed comment style', { ...threaded.parts, commentsXml: threaded.parts.commentsXml.replace('CommentText', 'CommentAnnotation') }]
];
for (const [label, changedParts] of restorationMutations) {
  const changedToken = await computeWordRestorationToken(changedParts);
  assert.notEqual(changedToken.value, threadedRestoreToken.value, `${label} remains in the full restoration parts token`);
  if (label !== 'changed comment style' && label !== 'changed durable comment ID') {
    assert.notDeepEqual(changedToken.commentIdentity, threadedRestoreToken.commentIdentity,
      `${label} changes the canonical comment identity projection`);
  } else {
    assert.deepEqual(changedToken.commentIdentity, threadedRestoreToken.commentIdentity,
      `${label} is detected by the full parts token even when canonical comment fields are unchanged`);
  }
}
assert.deepEqual(threaded.parts, threadedPartsCopy, 'restoration comparison never mutates saved source parts');
assert.equal(threaded.flatOpc, threadedFlatCopy, 'restoration comparison never rewrites saved Flat OPC checkpoint XML');

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

// Revert guard token: Word's own background changes after a write do not count; edits do.
{
  const guard = async (paragraphXml, settingsXml = settingsA) => (await computeWordRevertGuardToken(fixtureParts(paragraphXml, settingsXml))).value;
  const plain = await guard(`<w:p xmlns:w14="${W14}" w14:paraId="1111AAAA"><w:r><w:rPr><w:i/></w:rPr><w:t>Payment is due in thirty days.</w:t></w:r></w:p>`);
  const proofed = await guard(`<w:p xmlns:w14="${W14}" w14:paraId="9999FFFF" w14:textId="77777777" w:rsidR="00ABCDEF">`
    + `<w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>`
    + `<w:r w:rsidR="00000001"><w:rPr><w:i/></w:rPr><w:lastRenderedPageBreak/><w:t xml:space="preserve">Payment is due in </w:t></w:r>`
    + `<w:proofErr w:type="spellStart"/><w:r w:rsidR="00000002"><w:rPr><w:i/></w:rPr><w:t>thirty</w:t></w:r><w:proofErr w:type="spellEnd"/>`
    + `<w:r w:rsidRPr="00000003"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> days.</w:t></w:r></w:p>`,
    settingsA.replace('<w:doNotTrackMoves/>', '<w:doNotTrackMoves/><w:proofState w:spelling="clean" w:grammar="clean"/>'));
  assert.equal(proofed, plain, 'proofing marks, split runs, layout marks, _GoBack, paragraph ids and rsids are ignored');

  const typed = await guard(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Payment is due in sixty days.</w:t></w:r></w:p>`);
  const bolded = await guard(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Payment is due in </w:t></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>thirty</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> days.</w:t></w:r></w:p>`);
  const tracked = await guard(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Payment is due in thirty days.</w:t></w:r><w:ins w:id="5" w:author="User" w:date="2026-11-01T00:00:00Z"><w:r><w:rPr><w:i/></w:rPr><w:t> Net.</w:t></w:r></w:ins></w:p>`);
  const realBookmark = await guard(`<w:p><w:bookmarkStart w:id="0" w:name="Payment"/><w:bookmarkEnd w:id="0"/><w:r><w:rPr><w:i/></w:rPr><w:t>Payment is due in thirty days.</w:t></w:r></w:p>`);
  for (const [label, token] of [['typing', typed], ['formatting', bolded], ['a tracked insertion', tracked], ['a named bookmark', realBookmark]]) {
    assert.notEqual(token, plain, `${label} counts as a change`);
  }
}
console.log('PASS: the Revert guard token ignores Word background noise and still sees edits');
