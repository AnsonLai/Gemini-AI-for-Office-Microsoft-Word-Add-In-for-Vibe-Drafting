// Reproduces the public standalone batch comment-delta contract on an
// existing threaded-comments DOCX. This is a consumer transport repro, not an
// engine defect: the batch returns a merge delta, while injectCommentsIntoPackage
// appends a part instead of merging an existing comments part.
//
//   node docs/library-issues/2026-10-02-word-existing-comments-delta-reproducer.mjs
import '../../tests/setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  ensureCommentsArtifactsInZip,
  injectCommentsIntoPackage,
  openDocx,
} from '@ansonlai/docx-redline-js';
import { createParser } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';
import { MemoryZip } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const root = resolve(fileURLToPath(new URL('../..', import.meta.url)));
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

// The public ZIP helper does merge the comments.xml part, but its package
// boundary is a .docx ZIP, not the Flat OPC string passed to body.insertOoxml.
const encoder = new TextEncoder();
const zip = new MemoryZip(new Map([
  ['word/comments.xml', encoder.encode(commentsXml)],
  ['[Content_Types].xml', encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')],
  ['word/_rels/document.xml.rels', encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')],
]));
await ensureCommentsArtifactsInZip(zip, batch.commentsXml);
const zipMergedComments = await zip.file('word/comments.xml').async('string');
const zipMergedCommentCount = countComments(zipMergedComments);

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
  commentsAfterPublicZipHelperMerge: zipMergedCommentCount,
  textOnlyEditStatus: textOnly.status,
  textOnlyEditPreservedComments: textOnlyDoc.inspect().comments.length === sourceInspection.comments.length,
};
console.log(JSON.stringify(observations, null, 2));
assert.equal(deltaCommentCount, 1, 'the returned comment XML is a delta');
assert.equal(bodyStarts, 4, 'the returned document retains three existing anchors and adds one');
assert.equal(commentPartCount, 2, 'the published injector appends instead of merging the comments part');
console.log('REPRODUCED: consumer transport must handle merge-mode comments deltas without replacing existing comments.');
