// Reproduces the public standalone batch comment-delta contract on an
// existing threaded-comments DOCX. This is a consumer transport repro, not an
// engine defect: the batch returns a merge delta, while injectCommentsIntoPackage
// appends a part instead of merging an existing comments part.
//
//   node docs/library-issues/completed/2026-10-02-word-existing-comments-delta-reproducer.mjs
import '../../../tests/setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  ensureCommentsArtifactsInZip,
  ensureCommentsExtendedArtifactsInZip,
  injectCommentsIntoPackage,
  openDocx,
  validateDocxPackage,
} from '@ansonlai/docx-redline-js';
import { createParser } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';
import { MemoryZip, zipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { reconcileCommentSiblingParts } from '@ansonlai/docx-redline-js/services/comment-thread-parts.js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const root = resolve(fileURLToPath(new URL('../../..', import.meta.url)));
const fixturePath = resolve(root, 'tests/fixtures/wp6/threaded-comments.docx');
const sourceBytes = new Uint8Array(readFileSync(fixturePath));
const decoder = new TextDecoder();
const original = openDocx(sourceBytes);
const sourceInspection = original.inspect();
assert.equal(sourceInspection.status, 'ok', JSON.stringify(sourceInspection.error));
assert.equal(sourceInspection.comments.length, 3, 'the Word-authored fixture starts with three comments');
const targetParagraph = sourceInspection.paragraphs.find(
  paragraph => paragraph.exactText === 'This agreement is governed by local law.'
);
assert.ok(targetParagraph, 'the source paragraph is available');
const target = {
  index: targetParagraph.index,
  exactText: targetParagraph.exactText,
  paragraphId: targetParagraph.paragraphId,
  fingerprint: targetParagraph.fingerprint,
};
const getPartText = (doc, name) => {
  const bytes = doc.entries.get(name);
  return bytes ? decoder.decode(bytes) : null;
};
const documentXml = getPartText(original, 'word/document.xml');
const commentsXml = getPartText(original, 'word/comments.xml');
const commentsExtendedXml = getPartText(original, 'word/commentsExtended.xml');
assert.ok(documentXml && commentsXml && commentsExtendedXml, 'the source has complete comment package parts');

const batch = await applyOperationsToDocumentXml(documentXml, [
  {
    type: 'comment',
    target,
    textToComment: 'governed by local law',
    commentContent: 'Confirm this transition remains intentionally between the headings.',
  },
  {
    type: 'redline',
    target,
    modified: 'This agreement remains governed by local law.',
    structuredContent: false,
  },
], 'Existing Comments Reproducer', {
  commentsXml,
  commentsExtendedXml,
}, { atomic: true, strictTargets: true, generateRedlines: false });
assert.equal(batch.status, 'ok', JSON.stringify(batch.error || batch.results));
assert.equal(batch.commentsXmlMode, 'merge', 'ordinary comment output is explicitly a merge delta');
assert.equal(batch.commentsExtendedXml, null, 'an ordinary new root comment does not replace thread metadata');
assert.equal(batch.results.length, 2, 'the text rewrite and comment are both accounted');

const parser = createParser();
const NS_PACKAGE = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const NS_WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const countComments = xml => {
  const parsed = parser.parseFromString(xml, 'application/xml');
  return [...parsed.getElementsByTagNameNS(NS_WORD, 'comment')].length;
};
const bodyStarts = [...batch.documentXml.matchAll(/<w:commentRangeStart\b/g)].length;
const deltaCommentCount = countComments(batch.commentsXml);

// The published injector adds another comments part; it is not a merge API.
const flatPackage = `<pkg:package xmlns:pkg="${NS_PACKAGE}"><pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"><pkg:xmlData>${commentsXml}</pkg:xmlData></pkg:part></pkg:package>`;
const injected = injectCommentsIntoPackage(flatPackage, batch.commentsXml);
const injectedDoc = parser.parseFromString(injected, 'application/xml');
const commentPartCount = [...injectedDoc.getElementsByTagNameNS(NS_PACKAGE, 'part')]
  .filter(part => part.getAttribute('pkg:name') === '/word/comments.xml').length;

// Compose the published part helpers on a MemoryZip over the source DOCX. The
// complete merged parts can then be carried into the Office.js Flat OPC payload.
const zip = new MemoryZip(new Map([...original.entries].map(([name, bytes]) => [name, new Uint8Array(bytes)])));
zip.file('word/document.xml', batch.documentXml);
await ensureCommentsArtifactsInZip(zip, batch.commentsXml, {
  replaceExisting: batch.commentsXmlMode === 'replace',
});
if (batch.commentsExtendedXml) {
  await ensureCommentsExtendedArtifactsInZip(zip, batch.commentsExtendedXml);
}
const zipMergedComments = await zip.file('word/comments.xml').async('string');
const zipMergedIds = await zip.file('word/commentsIds.xml').async('string');
const zipMergedExtensible = await zip.file('word/commentsExtensible.xml').async('string');
const reconciled = reconcileCommentSiblingParts({
  commentsXml: zipMergedComments,
  commentsIdsXml: zipMergedIds,
  commentsExtensibleXml: zipMergedExtensible,
});
if (reconciled.commentsIdsXml) zip.file('word/commentsIds.xml', reconciled.commentsIdsXml);
if (reconciled.commentsExtensibleXml) zip.file('word/commentsExtensible.xml', reconciled.commentsExtensibleXml);
await validateDocxPackage(zip);
const mergedDoc = openDocx(zipDocx(zip.entries));
const mergedInspection = mergedDoc.inspect();
const zipMergedCommentCount = mergedInspection.comments.length;
const identity = ({ id, author, text, parentCommentId, done }) => ({
  id, author, text, parentCommentId: parentCommentId ?? null, done: done === true,
});
assert.deepEqual(
  mergedInspection.comments.slice(0, sourceInspection.comments.length).map(identity),
  sourceInspection.comments.map(identity),
  'the original root, reply, authors, texts, and resolved-thread state remain unchanged',
);
assert.equal(mergedInspection.comments.at(-1).text,
  'Confirm this transition remains intentionally between the headings.');
assert.equal(getPartText(mergedDoc, 'word/commentsExtended.xml'), commentsExtendedXml,
  'the existing extended-thread part is preserved for an ordinary new root comment');
assert.ok(mergedInspection.paragraphs.some(paragraph => paragraph.exactText === 'This agreement remains governed by local law.'),
  'the requested paragraph edit remains in the assembled DOCX');
const siblingIds = xml => {
  const parsed = parser.parseFromString(xml, 'application/xml');
  return [...parsed.getElementsByTagNameNS('*', 'commentId')].map(node => ({
    paraId: node.getAttribute('w16cid:paraId') || node.getAttribute('paraId'),
    durableId: node.getAttribute('w16cid:durableId') || node.getAttribute('durableId'),
  }));
};
const originalSiblingIds = siblingIds(getPartText(original, 'word/commentsIds.xml'));
const mergedSiblingIds = siblingIds(getPartText(mergedDoc, 'word/commentsIds.xml'));
assert.equal(mergedSiblingIds.length, 4, 'the new comment receives a sibling ID');
for (const existing of originalSiblingIds) {
  assert.ok(mergedSiblingIds.some(item => item.paraId === existing.paraId && item.durableId === existing.durableId),
    `existing durable ID ${existing.durableId} remains unchanged`);
}
const extensibleDoc = parser.parseFromString(getPartText(mergedDoc, 'word/commentsExtensible.xml'), 'application/xml');
assert.equal(extensibleDoc.getElementsByTagNameNS('*', 'commentExtensible').length, 4,
  'optional extensible records are reconciled for all four comments');

// Existing-thread text edits remain supported through the complete public DOCX API.
const textOnlyDoc = openDocx(sourceBytes);
const textOnly = await textOnlyDoc.applyOperations([{
  type: 'replace',
  target,
  replacements: [{ find: 'governed by local law', replace: 'applicable law' }],
}], { author: 'Existing Comments Reproducer', atomic: true, generateRedlines: false });
assert.equal(textOnly.status, 'ok', JSON.stringify(textOnly.error || textOnly.results));
assert.equal(textOnlyDoc.inspect().comments.length, 3, 'the existing comments survive an ordinary text edit');

const observations = {
  sourceComments: sourceInspection.comments.length,
  standaloneCommentMode: batch.commentsXmlMode,
  batchResults: batch.results.length,
  commentsInReturnedDelta: deltaCommentCount,
  commentAnchorsInReturnedDocument: bodyStarts,
  commentsPartsAfterPublishedInjector: commentPartCount,
  commentsAfterComposedPublicMerge: zipMergedCommentCount,
  preservedOriginalThreadState: mergedInspection.comments.slice(0, 3).every((comment, index) =>
    JSON.stringify(identity(comment)) === JSON.stringify(identity(sourceInspection.comments[index]))),
  preservedOriginalDurableIds: originalSiblingIds.every(existing =>
    mergedSiblingIds.some(item => item.paraId === existing.paraId && item.durableId === existing.durableId)),
  reconciledExtensibleRecords: extensibleDoc.getElementsByTagNameNS('*', 'commentExtensible').length,
  sameBatchParagraphEdit: mergedInspection.paragraphs.some(paragraph =>
    paragraph.exactText === 'This agreement remains governed by local law.'),
  textOnlyEditStatus: textOnly.status,
  textOnlyEditPreservedComments: textOnlyDoc.inspect().comments.length === sourceInspection.comments.length,
};
console.log(JSON.stringify(observations, null, 2));
assert.equal(deltaCommentCount, 1, 'the returned comment XML is a delta');
assert.equal(bodyStarts, 4, 'the returned document retains three existing anchors and adds one');
assert.equal(commentPartCount, 2, 'the published injector appends instead of merging the comments part');
assert.equal(zipMergedCommentCount, 4, 'the composed public helpers produce one complete four-comment part');
console.log('REPRODUCED AND FIXED IN CONSUMER: composed public helpers preserve threaded comments while merging the delta.');
