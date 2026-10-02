import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
    acceptTrackedChangesInOoxml,
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

const existingCommentPackage = withCommentsPart(packageXml, sourceCommentsXml);
const existingCommentInspection = consumer.inspectWordSource(existingCommentPackage);
assert.equal(existingCommentInspection.inspection.comments.length, 1);
const deltaRefusal = await consumer.prepareCanonicalBatch(existingCommentPackage, [{
    type: 'comment',
    target: { exactText: 'Source sentence.' },
    textToComment: 'Source',
    commentContent: 'New feedback.'
}], {
    runner: async documentXml => mockBatchResult(documentXml, newCommentXml, 'merge')
});
assert.equal(deltaRefusal.status, 'refused');
assert.equal(deltaRefusal.error?.code, 'UNSUPPORTED_EXISTING_COMMENTS');
assert.equal(deltaRefusal.result.commentsXmlMode, 'merge');
assert.equal(deltaRefusal.insertionPayload, undefined,
    'a comments delta cannot be emitted as a replacement package when source comments exist');

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
assert.equal(typeof globalThis.Word, 'undefined', 'portable import and execution do not require or install Word globals');

console.log('PASS: narrow canonical DOCX consumer core tests');
