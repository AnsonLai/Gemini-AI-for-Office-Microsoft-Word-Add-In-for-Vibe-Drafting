import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
    acceptTrackedChangesInOoxml,
    openDocx,
    rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { createParser, createSerializer } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import * as consumer from '../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';
import * as wordAdapter from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';

const corePath = fileURLToPath(new URL('../src/taskpane/modules/docx-redline-js-integration/consumer-core.js', import.meta.url));
const coreSource = readFileSync(corePath, 'utf8');
assert.doesNotMatch(coreSource, /from\s+['"][^'"]*(?:word-ooxml|word-operation-runner|word-redline-runner|taskpane\.js|commands\/agentic-tools)[^'"]*['"]/, 'portable core must not import Word transport or UI modules');
assert.doesNotMatch(coreSource, /from\s+['"](?:node:)?fs(?:\/promises)?['"]/, 'portable core must not import filesystem APIs');
assert.equal(typeof globalThis.Word, 'undefined', 'the consumer core test starts without Office.js globals');
assert.deepEqual(Object.keys(consumer).sort(), ['inspectWordSource', 'prepareCanonicalBatch'],
    'the integration core exposes only canonical source inspection and package preparation');
assert.equal(wordAdapter.inspectWordSource, consumer.inspectWordSource);
assert.equal(wordAdapter.prepareCanonicalBatch, consumer.prepareCanonicalBatch);

const packageXml = buildDocumentFragmentPackage(
    '<w:p><w:r><w:t>Source sentence.</w:t></w:r></w:p>',
    { appendTrailingParagraph: false }
);

const packageNamespace = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const sourceCommentsXml = `<w:comments xmlns:w="${wordNamespace}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="7" w:author="Existing reviewer" w:date="2026-10-02T00:00:00Z"><w:p w14:paraId="00000007"><w:r><w:t>Existing feedback stays.</w:t></w:r></w:p></w:comment></w:comments>`;
const newCommentXml = `<w:comments xmlns:w="${wordNamespace}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="8" w:author="New reviewer" w:date="2026-10-02T00:00:00Z"><w:p w14:paraId="00000008"><w:r><w:t>New feedback.</w:t></w:r></w:p></w:comment></w:comments>`;

function withCommentsPart(xml, commentsXml) {
    const parser = createParser();
    const packageDoc = parser.parseFromString(xml, 'application/xml');
    const commentsDoc = parser.parseFromString(commentsXml, 'application/xml');
    const part = packageDoc.createElementNS(packageNamespace, 'pkg:part');
    part.setAttribute('pkg:name', '/word/comments.xml');
    part.setAttribute('pkg:contentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
    const data = packageDoc.createElementNS(packageNamespace, 'pkg:xmlData');
    data.appendChild(packageDoc.importNode(commentsDoc.documentElement, true));
    part.appendChild(data);
    packageDoc.documentElement.appendChild(part);
    return createSerializer().serializeToString(packageDoc);
}

function mockBatchResult(documentXml, commentsXml, commentsXmlMode = 'merge') {
    return {
        status: 'ok',
        hasChanges: true,
        documentXml,
        commentsXml,
        commentsXmlMode,
        results: [],
        receipts: []
    };
}

const packageName = node => node.getAttribute('pkg:name') || node.getAttribute('name');
function flatOpcFromDocxEntries(entries) {
    const packageParts = [...entries].filter(([name]) => name !== '[Content_Types].xml').map(([name, bytes]) => {
        const value = new TextDecoder().decode(bytes);
        const contentType = name.endsWith('.rels')
            ? 'application/vnd.openxmlformats-package.relationships+xml'
            : 'application/xml';
        if (/^\s*</.test(value)) {
            return `<pkg:part pkg:name="/${name}" pkg:contentType="${contentType}"><pkg:xmlData>${value.replace(/^\s*<\?xml[^?]*\?>\s*/, '')}</pkg:xmlData></pkg:part>`;
        }
        return `<pkg:part pkg:name="/${name}" pkg:contentType="application/octet-stream"><pkg:binaryData>${Buffer.from(bytes).toString('base64')}</pkg:binaryData></pkg:part>`;
    }).join('');
    return `<pkg:package xmlns:pkg="${packageNamespace}">${packageParts}</pkg:package>`;
}
function countXmlElements(xml, localName) {
    return createParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', localName).length;
}
function commentSiblingIds(xml) {
    const doc = createParser().parseFromString(xml, 'application/xml');
    return Array.from(doc.getElementsByTagNameNS('*', 'commentId')).map(node => ({
        paraId: node.getAttribute('w16cid:paraId') || node.getAttribute('paraId'),
        durableId: node.getAttribute('w16cid:durableId') || node.getAttribute('durableId')
    }));
}

const existingCommentPackage = withCommentsPart(packageXml, sourceCommentsXml);
const existingCommentInspection = consumer.inspectWordSource(existingCommentPackage);
assert.equal(existingCommentInspection.inspection.comments.length, 1);
const deltaMerge = await consumer.prepareCanonicalBatch(existingCommentPackage, [{
    type: 'comment',
    target: { exactText: 'Source sentence.' },
    textToComment: 'Source',
    commentContent: 'New feedback.'
}], {
    runner: async documentXml => mockBatchResult(documentXml, newCommentXml, 'merge')
});
assert.equal(deltaMerge.status, 'ready', JSON.stringify(deltaMerge.error));
assert.equal(deltaMerge.result.commentsXmlMode, 'merge');
assert.deepEqual(
    consumer.inspectWordSource(deltaMerge.insertionPayload).inspection.comments.map(({ id, author, text }) => ({ id, author, text })),
    [
        { id: '7', author: 'Existing reviewer', text: 'Existing feedback stays.' },
        { id: '8', author: 'New reviewer', text: 'New feedback.' }
    ],
    'merge-mode comments append the delta while retaining existing comment content'
);

const threadedTextEdit = await consumer.prepareCanonicalBatch(existingCommentPackage, [{
    type: 'replace',
    target: { exactText: 'Source sentence.' },
    replacements: [{ find: 'Source', replace: 'Updated' }]
}], { author: 'Portable Core Test', expectedRevision: true });
assert.equal(threadedTextEdit.status, 'ready', JSON.stringify(threadedTextEdit.result?.error));
const preservedComments = consumer.inspectWordSource(threadedTextEdit.insertionPayload).inspection.comments;
assert.deepEqual(preservedComments.map(({ id, author, text }) => ({ id, author, text })), [
    { id: '7', author: 'Existing reviewer', text: 'Existing feedback stays.' }
], 'ordinary text edits keep existing threaded-comment parts intact');

const replacementCommentsXml = sourceCommentsXml.replace(
    '</w:comments>',
    newCommentXml.match(/<w:comment\b[\s\S]*?<\/w:comment>/)[0] + '</w:comments>'
);
const fullReplacement = await consumer.prepareCanonicalBatch(existingCommentPackage, [{
    type: 'comment',
    target: { exactText: 'Source sentence.' },
    textToComment: 'Source',
    commentContent: 'New feedback.'
}], {
    runner: async documentXml => mockBatchResult(documentXml, replacementCommentsXml, 'replace')
});
assert.equal(fullReplacement.status, 'ready', JSON.stringify(fullReplacement.result?.error));
assert.equal(consumer.inspectWordSource(fullReplacement.insertionPayload).inspection.comments.length, 2,
    'a full replacement remains supported when it includes both existing and new comments');

const firstCommentOnly = await consumer.prepareCanonicalBatch(packageXml, [{
    type: 'comment',
    target: { exactText: 'Source sentence.' },
    textToComment: 'Source',
    commentContent: 'New feedback.'
}], {
    runner: async documentXml => mockBatchResult(documentXml, newCommentXml, 'merge')
});
assert.equal(firstCommentOnly.status, 'ready', JSON.stringify(firstCommentOnly.result?.error));
assert.equal(consumer.inspectWordSource(firstCommentOnly.insertionPayload).inspection.comments.length, 1,
    'a delta can be assembled safely when the source has no existing comments');
const firstCommentPackage = createParser().parseFromString(firstCommentOnly.insertionPayload, 'application/xml');
const firstCommentRelsPart = Array.from(firstCommentPackage.getElementsByTagNameNS(packageNamespace, 'part'))
    .find(part => packageName(part) === '/word/_rels/document.xml.rels');
const firstCommentRelsRoot = Array.from(firstCommentRelsPart.getElementsByTagNameNS(packageNamespace, 'xmlData')[0].childNodes)
    .find(node => node.nodeType === 1);
assert.ok(Array.from(firstCommentRelsRoot.getElementsByTagNameNS('*', 'Relationship'))
    .some(relationship => relationship.getAttribute('Type')?.endsWith('/comments')),
'new comments receive their package relationship');

const packageWithoutDocumentRels = createParser().parseFromString(packageXml, 'application/xml');
const documentRelsPart = Array.from(packageWithoutDocumentRels.getElementsByTagNameNS(packageNamespace, 'part'))
    .find(part => packageName(part) === '/word/_rels/document.xml.rels');
documentRelsPart.parentNode.removeChild(documentRelsPart);
const noRelsSource = createSerializer().serializeToString(packageWithoutDocumentRels.documentElement);
const noRelsComment = await consumer.prepareCanonicalBatch(noRelsSource, [{
    type: 'comment',
    target: { exactText: 'Source sentence.' },
    textToComment: 'Source',
    commentContent: 'First comment without source relationships.'
}], {
    runner: async documentXml => mockBatchResult(documentXml, newCommentXml, 'merge')
});
assert.equal(noRelsComment.status, 'ready', JSON.stringify(noRelsComment.error));
const noRelsOutput = createParser().parseFromString(noRelsComment.insertionPayload, 'application/xml');
assert.ok(Array.from(noRelsOutput.getElementsByTagNameNS(packageNamespace, 'part'))
    .some(part => packageName(part) === '/word/_rels/document.xml.rels'),
    'the package adapter creates document relationships when the source omitted the part');
assert.equal(consumer.inspectWordSource(noRelsComment.insertionPayload).inspection.comments.length, 1);

// Exercise the public batch result against a Word-authored source with a threaded reply,
// a resolved thread, commentsExtended metadata, and optional durable-ID sibling parts.
const threadedBytes = new Uint8Array(readFileSync(new URL('./fixtures/wp6/threaded-comments.docx', import.meta.url)));
const threadedDoc = openDocx(threadedBytes);
const threadedBefore = threadedDoc.inspect();
const threadedPackageXml = flatOpcFromDocxEntries(threadedDoc.entries);
const threadedSourceParts = consumer.inspectWordSource(threadedPackageXml);

// A commentsExtended merge delta cannot be applied to a package that already has thread
// metadata; the batch is refused before any insertion package exists.
const threadDelta = await consumer.prepareCanonicalBatch(threadedPackageXml, [{
    type: 'replace', target: { exactText: 'unused by the mocked runner' }, modified: 'unused'
}], {
    author: 'Portable Core Test',
    runner: async documentXml => ({
        ...mockBatchResult(documentXml, null),
        commentsExtendedXml: '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
        commentsExtendedXmlMode: 'merge'
    })
});
assert.equal(threadDelta.status, 'error');
assert.equal(threadDelta.error?.code, 'UNSUPPORTED_EXISTING_COMMENT_THREAD_METADATA');
assert.equal(threadDelta.insertionPayload, undefined);
const governed = threadedBefore.paragraphs.find(paragraph => paragraph.exactText === 'This agreement is governed by local law.');
assert.ok(governed, 'Word-authored threaded source contains the tracked paragraph target');
const threadedPrepared = await consumer.prepareCanonicalBatch(threadedPackageXml, source => {
        const target = source.paragraphs.find(paragraph => paragraph.paragraphId === governed.paragraphId);
        return [
            {
                type: 'comment',
                target: { index: target.index, exactText: target.exactText, paragraphId: target.paragraphId, fingerprint: target.fingerprint },
                textToComment: 'governed by local law',
                commentContent: 'Confirm this transition remains intentionally between the headings.'
            },
            {
                type: 'replace',
                target: { index: target.index, exactText: target.exactText, paragraphId: target.paragraphId, fingerprint: target.fingerprint },
                replacements: [{ find: 'local', replace: 'national' }]
            }
        ];
    }, { author: 'Portable Core Test', generateRedlines: false });
assert.equal(threadedPrepared.status, 'ready', JSON.stringify(threadedPrepared.error || threadedPrepared.result?.error));
assert.equal(threadedPrepared.result.commentsXmlMode, 'merge');
assert.equal(countXmlElements(threadedPrepared.result.commentsXml, 'comment'), 1,
    'the standalone runner still returns its one-comment delta');
const threadedAfter = consumer.inspectWordSource(threadedPrepared.insertionPayload);
assert.equal(threadedAfter.inspection.comments.length, 4, 'full package contains all three existing comments plus the addition');
const commentIdentity = ({ id, author, text, parentCommentId, done }) => ({
    id,
    author,
    text,
    parentCommentId: parentCommentId ?? null,
    done: done === true
});
assert.deepEqual(
    threadedAfter.inspection.comments.slice(0, threadedBefore.comments.length).map(commentIdentity),
    threadedBefore.comments.map(commentIdentity),
    'existing root, reply, author, text, parent, and resolved state are preserved'
);
assert.equal(threadedAfter.inspection.comments.at(-1).text,
    'Confirm this transition remains intentionally between the headings.');
assert.ok(threadedAfter.inspection.paragraphs.some(paragraph => paragraph.exactText === 'This agreement is governed by national law.'),
    'the same batch keeps its requested paragraph edit');
assert.equal(threadedAfter.parts.commentsExtendedXml, threadedSourceParts.parts.commentsExtendedXml,
    'a new root comment leaves existing thread metadata unchanged');
const beforeDurableIds = commentSiblingIds(threadedSourceParts.parts.commentsIdsXml);
const afterDurableIds = commentSiblingIds(threadedAfter.parts.commentsIdsXml);
assert.equal(afterDurableIds.length, 4, 'the merged comments part receives a sibling ID for the new comment');
for (const original of beforeDurableIds) {
    assert.ok(afterDurableIds.some(item => item.paraId === original.paraId && item.durableId === original.durableId),
        `existing durable ID ${original.durableId} is preserved`);
}
assert.equal(countXmlElements(threadedAfter.parts.commentsExtensibleXml, 'commentExtensible'), 4,
    'optional extensible metadata is reconciled against the complete merged comments part');

const threadedReply = await consumer.prepareCanonicalBatch(threadedPackageXml, [{
    type: 'comment_reply',
    parentCommentId: '0',
    commentContent: 'Add a reply while retaining this Word thread.'
}], { author: 'Portable Core Test' });
assert.equal(threadedReply.status, 'ready', JSON.stringify(threadedReply.error || threadedReply.result?.error));
assert.equal(threadedReply.result.commentsXmlMode, 'replace');
assert.equal(threadedReply.result.commentsExtendedXmlMode, 'replace');
assert.equal(typeof threadedReply.result.commentsExtendedXml, 'string',
    'the thread operation supplies complete replacement metadata');
const threadedReplyAfter = consumer.inspectWordSource(threadedReply.insertionPayload);
assert.equal(threadedReplyAfter.inspection.comments.length, 4);
assert.deepEqual(
    threadedReplyAfter.inspection.comments.slice(0, threadedBefore.comments.length).map(commentIdentity),
    threadedBefore.comments.map(commentIdentity),
    'replace-mode thread output retains existing comments and resolved state'
);
assert.equal(threadedReplyAfter.inspection.comments.at(-1).parentCommentId, '0');
assert.equal(threadedReplyAfter.inspection.comments.at(-1).text, 'Add a reply while retaining this Word thread.');

const inspected = consumer.inspectWordSource(packageXml);
assert.equal(inspected.inspection.status, 'ok');
assert.equal(inspected.inspection.paragraphs.length, 1);
assert.equal(inspected.inspection.paragraphs[0].exactText, 'Source sentence.');

const prepared = await consumer.prepareCanonicalBatch(packageXml, [{
    type: 'replace',
    target: { exactText: 'Source sentence.' },
    modified: 'Portable sentence.'
}], { author: 'Portable Core Test', expectedRevision: true });
assert.equal(prepared.status, 'ready', JSON.stringify(prepared.result?.error));
assert.equal(prepared.source.paragraphs[0].exactText, 'Source sentence.');
assert.equal(prepared.operations.length, 1);
assert.equal(typeof prepared.insertionPayload, 'string');
assert.equal(typeof prepared.result.documentXml, 'string');
const accepted = acceptTrackedChangesInOoxml(prepared.result.documentXml, { allAuthors: true });
const rejected = rejectTrackedChangesInOoxml(prepared.result.documentXml, { allAuthors: true });
assert.equal(accepted.status, undefined, JSON.stringify(accepted.error));
assert.equal(rejected.status, undefined, JSON.stringify(rejected.error));
const acceptedSource = consumer.inspectWordSource(accepted.oxml);
const rejectedSource = consumer.inspectWordSource(rejected.oxml);
assert.equal(acceptedSource.inspection.paragraphs[0].exactText, 'Portable sentence.');
assert.equal(rejectedSource.inspection.paragraphs[0].exactText, 'Source sentence.');

const refused = await consumer.prepareCanonicalBatch(packageXml, [{
    type: 'replace',
    target: { exactText: 'Missing source.' },
    modified: 'Should refuse.'
}], { author: 'Portable Core Test' });
assert.equal(refused.status, 'refused');
assert.equal(refused.result.status, 'error');
assert.equal(refused.insertionPayload, undefined, 'a refused library result must not prepare a host payload');

// Word always supplies a body Flat OPC package. A bare w:document can still be inspected
// (resolved Accept All/Reject All views), but only a package can be written back.
const bareDocument = `<w:document xmlns:w="${wordNamespace}"><w:body><w:p><w:r><w:t>Bare sentence.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
assert.equal(consumer.inspectWordSource(bareDocument).inspection.paragraphs[0].exactText, 'Bare sentence.');
assert.throws(() => consumer.inspectWordSource(`<w:p xmlns:w="${wordNamespace}"><w:r><w:t>Fragment.</w:t></w:r></w:p>`),
    /Flat OPC package or a w:document is required/, 'a paragraph fragment is not a Word source');
const bareWrite = await consumer.prepareCanonicalBatch(bareDocument, [{
    type: 'replace', target: { exactText: 'Bare sentence.' }, modified: 'Changed sentence.'
}], { author: 'Portable Core Test' });
assert.equal(bareWrite.status, 'error');
assert.equal(bareWrite.error?.code, 'WORD_FLAT_OPC_REQUIRED');
assert.equal(bareWrite.insertionPayload, undefined, 'no insertion package without a Flat OPC source');
assert.equal(typeof globalThis.Word, 'undefined', 'portable import and execution do not require or install Word globals');

console.log('PASS: narrow canonical DOCX consumer core tests');
