import assert from 'node:assert/strict';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';

function escapeXml(text) {
  return String(text).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

function documentParts(texts) {
  const paragraphs = texts.map((text, index) => (
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0').toUpperCase()}"><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`
  )).join('');
  return {
    documentXml: `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`,
    commentsXml: null,
    commentsExtendedXml: null,
    commentsIdsXml: null,
    commentsExtensibleXml: null,
    numberingXml: null,
    stylesXml: '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    additionalParts: { 'word/settings.xml': '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>' }
  };
}

let parts = documentParts(['Alpha clause', 'Middle paragraph', 'Beta clause', 'Final paragraph']);
let documentId = 'addin:doc-1';
let revision = 0;
let readCount = 0;
const sourceReader = async () => {
  readCount++;
  return { documentId, revision, revisionView: 'accepted', parts };
};
const session = createDocumentInspectionSession({ readSource: sourceReader, maxResults: 4, maxTextChars: 128 });

const firstPage = await session.inspect({ search: 'clause', around: 1, limit: 2 });
assert.equal(firstPage.status, 'ok');
assert.equal(firstPage.revisionView, 'accepted');
assert.equal(firstPage.selection.totalMatches, 2);
assert.equal(firstPage.selection.truncated, false);
assert.equal(firstPage.selection.nextAfter, undefined, 'provider-facing inspection must not expose physical offsets');
assert.equal(firstPage.targets.length, 4, 'context paragraphs are returned around both matches');
for (const target of firstPage.targets) {
  assert.equal(typeof target.handle, 'string');
  assert.equal(typeof target.location, 'string');
  assert.equal(typeof target.text, 'string');
  assert.equal(Object.hasOwn(target, 'index'), false);
  assert.equal(Object.hasOwn(target, 'paragraphIndex'), false);
  assert.equal(Object.hasOwn(target, 'paragraphId'), false);
  assert.equal(Object.hasOwn(target, 'fingerprint'), false);
  assert.equal(Object.hasOwn(target, 'contextFor') && target.contextFor.some(Number.isInteger), false);
}

const alpha = firstPage.targets.find(target => target.text === 'Alpha clause');
const beta = firstPage.targets.find(target => target.text === 'Beta clause');
const beforeResolve = readCount;
const both = await session.resolveHandles([alpha.handle, beta.handle]);
assert.equal(both.status, 'ok');
assert.equal(readCount, beforeResolve + 1, 'all handles are verified against one fresh source read');
assert.equal(both.targets[0].descriptor.index, 1);
assert.equal(both.targets[1].descriptor.index, 3);
assert.equal(both.targets[0].descriptor.exactText, 'Alpha clause');
assert.equal(both.targets[0].fullTextInspected, true);
assert.equal(both.source.documentId, documentId);
assert.equal(both.source.revisionView, 'accepted');
assert.equal(both.source.revisionToken.scope, 'document-parts');

const resolvedSingle = await session.resolve(beta.handle);
assert.equal(resolvedSingle.status, 'ok');
assert.equal(resolvedSingle.target.kind, 'paragraph');
assert.equal(resolvedSingle.target.descriptor.paragraphId, '00000003');

const rangeInspection = await session.inspect({ range: { start: 1, end: 3 }, limit: 3 });
assert.equal(rangeInspection.targets.length, 3);
const resolvedRange = await session.resolveRange({
  startHandle: rangeInspection.targets[0].handle,
  endHandle: rangeInspection.targets[2].handle
});
assert.equal(resolvedRange.status, 'ok');
assert.equal(resolvedRange.range.paragraphCount, 3);
assert.equal(resolvedRange.range.start.index, 1);
assert.equal(resolvedRange.range.end.index, 3);

const incompleteRangeInspection = await session.inspect({ search: 'clause', limit: 2 });
const incompleteRange = await session.resolveRange({
  startHandle: incompleteRangeInspection.targets[0].handle,
  endHandle: incompleteRangeInspection.targets[1].handle
});
assert.equal(incompleteRange.status, 'error');
assert.equal(incompleteRange.error.code, 'RANGE_NOT_FULLY_INSPECTED', 'a search result cannot imply the unreturned paragraphs between matches');

const endInspection = await session.inspect({ includeDocumentEnd: true, limit: 2 });
assert.equal(typeof endInspection.documentEnd.handle, 'string');
assert.equal(Object.hasOwn(endInspection.documentEnd, 'index'), false);
const endTarget = await session.resolve(endInspection.documentEnd.handle);
assert.equal(endTarget.status, 'ok');
assert.equal(endTarget.target.kind, 'document-end');
assert.equal(endTarget.target.position, 'after');
assert.equal(endTarget.target.fullTextInspected, false, 'an end handle does not authorize rewriting unseen final text');
assert.equal(endTarget.target.descriptor.index, 4);

const cursorPage1 = await session.inspect({ limit: 2 });
assert.equal(cursorPage1.selection.truncated, true);
assert.equal(typeof cursorPage1.selection.nextCursor, 'string');
assert.equal(Object.hasOwn(cursorPage1.selection, 'nextAfter'), false);
const cursorPage2 = await session.inspect({ cursor: cursorPage1.selection.nextCursor });
assert.equal(cursorPage2.status, 'ok', JSON.stringify(cursorPage2));
assert.deepEqual(cursorPage2.targets.map(target => target.text), ['Beta clause', 'Final paragraph']);
const pagedRange = await session.resolveRange({
  startHandle: cursorPage1.targets[0].handle,
  endHandle: cursorPage2.targets[1].handle
});
assert.equal(pagedRange.status, 'error');
assert.equal(pagedRange.error.code, 'RANGE_NOT_FULLY_INSPECTED', 'pagination does not silently widen a range');

const suppliedSource = await sourceReader();
const readsBeforeSuppliedSource = readCount;
const resolvedWithSource = await session.resolveHandles([cursorPage2.targets[0].handle], { source: suppliedSource });
assert.equal(resolvedWithSource.status, 'ok');
assert.equal(readCount, readsBeforeSuppliedSource, 'trusted Word source capture can be reused by the compiler for the guarded batch');

const staleHandle = cursorPage2.targets[0].handle;
parts = documentParts(['Alpha clause', 'Middle paragraph changed', 'Beta clause', 'Final paragraph']);
const changedSource = await session.resolve(staleHandle);
assert.equal(changedSource.status, 'error');
assert.equal(changedSource.error.code, 'STALE_DOCUMENT_SOURCE');
assert.equal((await session.resolve(staleHandle)).error.code, 'STALE_OR_UNKNOWN_HANDLE');

const afterChange = await session.inspect({ range: { start: 1, end: 2 }, limit: 2 });
const oldDocumentHandle = afterChange.targets[0].handle;
documentId = 'addin:doc-2';
const switchedDocument = await session.resolve(oldDocumentHandle);
assert.equal(switchedDocument.status, 'error');
assert.equal(switchedDocument.error.code, 'DOCUMENT_IDENTITY_CHANGED');

documentId = 'addin:doc-1';
const afterSwitch = await session.inspect({ range: { start: 1, end: 2 }, limit: 2 });
const beforeRevisionChange = afterSwitch.targets[0].handle;
revision++;
const changedRevision = await session.resolve(beforeRevisionChange);
assert.equal(changedRevision.status, 'error');
assert.equal(changedRevision.error.code, 'STALE_DOCUMENT_SOURCE');

const afterRevision = await session.inspect({ range: { start: 1, end: 2 }, limit: 2 });
const undoSensitiveHandle = afterRevision.targets[0].handle;
session.invalidate('undo-or-refresh');
parts = documentParts(['Alpha clause', 'Middle paragraph changed', 'Beta clause', 'Final paragraph']);
const restoredSnapshot = await session.resolve(undoSensitiveHandle);
assert.equal(restoredSnapshot.status, 'error');
assert.equal(restoredSnapshot.error.code, 'STALE_OR_UNKNOWN_HANDLE', 'explicit host invalidation prevents reusing a handle even if content is restored');

const missingBaseline = createDocumentInspectionSession({ readSource: async () => ({ documentId }) });
const missing = await missingBaseline.inspect();
assert.equal(missing.status, 'error');
assert.equal(missing.error.code, 'MISSING_CANONICAL_SOURCE');

const wrongViewSession = createDocumentInspectionSession({
  readSource: async () => ({ documentId, parts, revisionView: 'rejected' }),
  revisionView: 'accepted'
});
const wrongView = await wrongViewSession.inspect();
assert.equal(wrongView.status, 'error');
assert.equal(wrongView.error.code, 'REVISION_VIEW_MISMATCH');

const markerSession = createDocumentInspectionSession({
  readSource: async () => ({ documentId: 'addin:markers', parts: documentParts(['[P99] Marker paragraph']) }),
  maxResults: 1
});
const markerInspection = await markerSession.inspect({ range: { start: 1, end: 1 }, limit: 1 });
assert.equal(markerInspection.targets[0].text, '[P99] Marker paragraph');
assert.doesNotMatch(markerInspection.targets[0].location, /\bP\d+\b/i, 'human location strips generated physical paragraph markers');
assert.doesNotMatch(JSON.stringify(markerInspection.targets[0].cues), /\bP\d+\b/i, 'cues do not disclose generated paragraph markers');

const longParts = documentParts(['A long paragraph with source text '.repeat(4), 'Another long paragraph with source text '.repeat(4)]);
const boundedSession = createDocumentInspectionSession({
  readSource: async () => ({ documentId: 'addin:bounded', parts: longParts }),
  maxResults: 2,
  maxTextChars: 1500,
  maxTotalTextChars: 24
});
const boundedInspection = await boundedSession.inspect({ range: { start: 1, end: 2 }, limit: 2 });
assert.equal(boundedInspection.status, 'ok');
assert.ok(boundedInspection.targets.some(target => target.textTruncated));
assert.ok(boundedInspection.targets.reduce((total, target) => total + target.text.length + target.location.length, 0) <= 24);
const boundedResolved = await boundedSession.resolve(boundedInspection.targets[0].handle);
assert.equal(boundedResolved.status, 'ok');
assert.equal(boundedResolved.target.fullTextInspected, false, 'truncated inspection cannot authorize a rewrite of the unseen paragraph text');
const boundedRange = await boundedSession.resolveRange({
  startHandle: boundedInspection.targets[0].handle,
  endHandle: boundedInspection.targets[1].handle
});
assert.equal(boundedRange.status, 'error');
assert.equal(boundedRange.error.code, 'RANGE_NOT_FULLY_INSPECTED');
const boundedEndInspection = await boundedSession.inspect({ includeDocumentEnd: true, limit: 1 });
const boundedEnd = await boundedSession.resolve(boundedEndInspection.documentEnd.handle);
assert.equal(boundedEnd.status, 'ok');
assert.equal(boundedEnd.target.fullTextInspected, false, 'an end handle keeps append context internally without authorizing a rewrite');

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const w15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const commentThreadParts = documentParts(['Alpha note', 'Reply note']);
commentThreadParts.documentXml = `<?xml version="1.0"?><w:document xmlns:w="${w}" xmlns:w14="${w14}"><w:body><w:p w14:paraId="00000001"><w:commentRangeStart w:id="0"/><w:r><w:t>Alpha note</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:p w14:paraId="00000002"><w:commentRangeStart w:id="1"/><w:r><w:t>Reply note</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>`;
commentThreadParts.commentsXml = `<w:comments xmlns:w="${w}" xmlns:w14="${w14}"><w:comment w:id="0" w:author="Reviewer" w:date="2026-10-02T00:00:00Z"><w:p w14:paraId="A0000001"><w:r><w:t>Root comment</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="Reply author" w:date="2026-10-02T00:00:00Z"><w:p w14:paraId="B0000002"><w:r><w:t>Reply comment</w:t></w:r></w:p></w:comment></w:comments>`;
commentThreadParts.commentsExtendedXml = `<w15:commentsEx xmlns:w15="${w15}"><w15:commentEx w15:paraId="A0000001" w15:done="1"/><w15:commentEx w15:paraId="B0000002" w15:paraIdParent="A0000001" w15:done="0"/></w15:commentsEx>`;
const commentsSession = createDocumentInspectionSession({
  readSource: async () => ({ documentId: 'addin:comments', parts: commentThreadParts })
});
const commentInspection = await commentsSession.inspect({ range: { start: 1, end: 2 }, limit: 2 });
assert.equal(commentInspection.status, 'ok');
assert.deepEqual(commentInspection.targets[0].cues.comments, [
  { author: 'Reviewer', text: 'Root comment', resolved: true, isReply: false }
]);
assert.deepEqual(commentInspection.targets[1].cues.comments, [
  { author: 'Reply author', text: 'Reply comment', resolved: false, isReply: true }
]);

const dependencyParts = documentParts(['Dependency target']);
const dependencySession = createDocumentInspectionSession({
  readSource: async () => ({ documentId: 'addin:dependencies', parts: dependencyParts }),
  maxResults: 1
});
const dependencyChanges = [
  ['commentsXml', null, `<w:comments xmlns:w="${w}"><w:comment w:id="5"><w:p><w:r><w:t>Changed</w:t></w:r></w:p></w:comment></w:comments>`],
  ['commentsExtendedXml', null, `<w15:commentsEx xmlns:w15="${w15}"/>`],
  ['numberingXml', null, `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"/></w:numbering>`],
  ['stylesXml', dependencyParts.stylesXml, `<w:styles xmlns:w="${w}"><w:style w:type="paragraph" w:styleId="Changed"/></w:styles>`]
];
for (const [partName, baseline, changed] of dependencyChanges) {
  dependencyParts[partName] = baseline;
  const dependencyInspection = await dependencySession.inspect({ range: { start: 1, end: 1 }, limit: 1 });
  assert.equal(dependencyInspection.status, 'ok', `${partName} baseline should inspect`);
  dependencyParts[partName] = changed;
  const staleDependency = await dependencySession.resolve(dependencyInspection.targets[0].handle);
  assert.equal(staleDependency.status, 'error', `${partName} change invalidates an existing handle`);
  assert.equal(staleDependency.error.code, 'STALE_DOCUMENT_SOURCE');
}
dependencyParts.additionalParts['word/commentsIds.xml'] = '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>';
const commentsIdsInspection = await dependencySession.inspect({ range: { start: 1, end: 1 }, limit: 1 });
dependencyParts.additionalParts['word/commentsIds.xml'] += '<!-- updated -->';
const staleCommentsIds = await dependencySession.resolve(commentsIdsInspection.targets[0].handle);
assert.equal(staleCommentsIds.status, 'error');
assert.equal(staleCommentsIds.error.code, 'STALE_DOCUMENT_SOURCE', 'the canonical token includes comments ID parts supplied through additionalParts');
dependencyParts.additionalParts['word/commentsExtensible.xml'] = '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"/>';
const commentsExtensibleInspection = await dependencySession.inspect({ range: { start: 1, end: 1 }, limit: 1 });
dependencyParts.additionalParts['word/commentsExtensible.xml'] += '<!-- updated -->';
const staleCommentsExtensible = await dependencySession.resolve(commentsExtensibleInspection.targets[0].handle);
assert.equal(staleCommentsExtensible.status, 'error');
assert.equal(staleCommentsExtensible.error.code, 'STALE_DOCUMENT_SOURCE', 'the canonical token includes comments extensible parts supplied through additionalParts');

console.log('PASS: canonical source-bound inspection handles, ranges, cursors, freshness and invalidation');
