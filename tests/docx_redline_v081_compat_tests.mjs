import assert from 'node:assert/strict';
import { openDocx, validateDocxPackage } from '@ansonlai/docx-redline-js';
import { MemoryZip, zipDocx, unzipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { assertRedlineResult, RedlineOperationError } from '../src/taskpane/modules/docx-redline-js-integration/redline-result.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const COMMENTS_EXTENDED_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';

function packageBytes({ footer = false, secondFooter = false, field = false } = {}) {
    const contentTypes = `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${footer ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}${secondFooter ? '<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}</Types>`;
    const documentRels = `<Relationships xmlns="${PR}">${footer ? `<Relationship Id="rId1" Type="${R}/footer" Target="footer1.xml"/>` : ''}${secondFooter ? `<Relationship Id="rId2" Type="${R}/footer" Target="footer2.xml"/>` : ''}</Relationships>`;
    const sectionBreak = secondFooter ? '<w:pPr><w:sectPr><w:footerReference w:type="default" r:id="rId1"/></w:sectPr></w:pPr>' : '';
    const documentXml = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"><w:body><w:p w14:paraId="1A2B3C4D"><w:r><w:t>Alpha beta gamma.</w:t></w:r></w:p><w:p w14:paraId="2A2B3C4D">${sectionBreak}<w:r><w:t>Second paragraph.</w:t></w:r></w:p><w:sectPr>${secondFooter ? '<w:footerReference w:type="default" r:id="rId2"/>' : footer ? '<w:footerReference w:type="default" r:id="rId1"/>' : ''}</w:sectPr></w:body></w:document>`;
    const entries = new Map([
        ['[Content_Types].xml', encoder.encode(contentTypes)],
        ['_rels/.rels', encoder.encode(`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
        ['word/document.xml', encoder.encode(documentXml)],
        ['word/_rels/document.xml.rels', encoder.encode(documentRels)]
    ]);
    if (footer) entries.set('word/footer1.xml', encoder.encode(field
        ? `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`
        : `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Draft footer</w:t></w:r></w:p></w:ftr>`));
    if (secondFooter) entries.set('word/footer2.xml', encoder.encode(`<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>First footer</w:t></w:r></w:p></w:ftr>`));
    return zipDocx(entries);
}

function part(bytes, name) {
    const data = unzipDocx(bytes).get(name);
    assert.ok(data, `Missing ${name}`);
    return decoder.decode(data);
}

function errorCode(result) {
    return result.results?.find(item => item.error)?.error?.code || result.error?.code;
}

async function testOpenInspectBatchAndResolution() {
    const doc = openDocx(packageBytes());
    const before = doc.inspect();
    assert.equal(before.status, 'ok');
    assert.equal(before.paragraphs[0].index, 1);
    assert.equal(before.paragraphs[0].ref, 'P1');
    assert.equal(before.paragraphs[0].exactText, 'Alpha beta gamma.');
    assert.ok(before.paragraphs[0].paragraphId);
    const result = await doc.applyOperations([
        { type: 'replace', target: { exactText: 'Second paragraph.' }, modified: 'Second updated paragraph.' },
        { type: 'replace', target: { exactText: 'Alpha beta gamma.' }, replacements: [{ find: 'beta', replace: 'BETA' }] }
    ], { author: 'Compatibility Test', atomic: true });
    assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
    assert.equal(result.written, true);
    assert.ok(result.toUint8Array() instanceof Uint8Array);
    assert.ok(result.toUint8Array().length > 0);
    assert.match(part(result.toUint8Array(), 'word/document.xml'), /BETA/);
    assert.equal(doc.inspect().paragraphs[0].text, 'Alpha BETA gamma.');
    const accepted = await doc.resolveRevisions('accept', { allAuthors: true });
    assert.equal(accepted.status, 'ok', JSON.stringify(accepted.error));
    assert.equal(doc.inspect().paragraphs[1].text, 'Second updated paragraph.');
}

async function testCommentedPackageAndThreadOperations() {
    const doc = openDocx(packageBytes());
    const added = await doc.applyOperations([{
        type: 'comment', target: { exactText: 'Alpha beta gamma.' },
        textToComment: 'beta', commentContent: 'Review this word.'
    }], { author: 'Reviewer', atomic: true });
    assert.equal(added.status, 'ok', JSON.stringify(added.error || added.results));
    assert.equal(added.written, true);
    const root = doc.inspect().comments[0];
    assert.ok(root?.id);
    const reply = await doc.applyOperations([{ type: 'comment_reply', parentCommentId: root.id, commentContent: 'Agreed.' }], { author: 'Editor', atomic: true });
    assert.equal(reply.status, 'ok', JSON.stringify(reply.error || reply.results));
    assert.equal(doc.inspect().comments.length, 2);
    const commentXml = part(doc.toUint8Array(), 'word/comments.xml');
    const extendedXml = part(doc.toUint8Array(), 'word/commentsExtended.xml');
    const commentsRels = part(doc.toUint8Array(), 'word/_rels/document.xml.rels');
    assert.ok(part(doc.toUint8Array(), '[Content_Types].xml').includes(COMMENTS_EXTENDED_TYPE));
    assert.equal((part(doc.toUint8Array(), 'word/document.xml').match(/<w:commentRangeStart\b/g) || []).length, 2);

    const edited = await doc.applyOperations([{ type: 'replace', target: { exactText: 'Second paragraph.' }, modified: 'Second edited paragraph.' }], { author: 'Editor', atomic: true });
    assert.equal(edited.status, 'ok', JSON.stringify(edited.error || edited.results));
    assert.equal(part(doc.toUint8Array(), 'word/comments.xml'), commentXml);
    assert.equal(part(doc.toUint8Array(), 'word/commentsExtended.xml'), extendedXml);
    assert.equal(part(doc.toUint8Array(), 'word/_rels/document.xml.rels'), commentsRels);
    const anchoredEdit = await doc.applyOperations([{
        type: 'replace', target: { exactText: 'Alpha beta gamma.' },
        replacements: [{ find: 'gamma', replace: 'delta' }]
    }], { author: 'Editor', atomic: true });
    assert.equal(anchoredEdit.status, 'ok', JSON.stringify(anchoredEdit.error || anchoredEdit.results));
    const anchorDocument = part(doc.toUint8Array(), 'word/document.xml');
    assert.equal((anchorDocument.match(/<w:commentRangeEnd\b/g) || []).length, 2);
    assert.equal((anchorDocument.match(/<w:commentReference\b/g) || []).length, 2);
    assert.equal(doc.inspect().paragraphs[0].text, 'Alpha beta delta.');

    const resolved = await doc.resolveComment(root.id, { resolved: true });
    assert.equal(resolved.status, 'ok', JSON.stringify(resolved.error));
    assert.ok(resolved.artifactsChanged.includes('word/commentsExtended.xml'));
    assert.equal(doc.inspect().comments.filter(comment => comment.done).length, 2);
    const alreadyResolved = await doc.resolveComment(root.id, { resolved: true });
    assert.equal(alreadyResolved.status, 'ok');
    assert.equal(alreadyResolved.hasChanges, false);
    assert.equal(alreadyResolved.written, false);
    const reopened = await doc.applyOperations([{ type: 'comment_resolve', commentId: root.id, resolved: false }], { atomic: true });
    assert.equal(reopened.status, 'ok', JSON.stringify(reopened.error || reopened.results));
    assert.equal(doc.inspect().comments.filter(comment => comment.done).length, 0);
    const beforeUnknown = doc.toUint8Array();
    const missing = await doc.resolveComment('99999');
    assert.equal(errorCode(missing), 'COMMENT_NOT_FOUND');
    assert.deepEqual(doc.toUint8Array(), beforeUnknown);
    const missingDelete = await doc.deleteComments({ ids: ['99999'] });
    assert.equal(errorCode(missingDelete), 'COMMENT_NOT_FOUND');
    assert.deepEqual(doc.toUint8Array(), beforeUnknown);

    const damagedEntries = unzipDocx(beforeUnknown);
    damagedEntries.set('[Content_Types].xml', encoder.encode(
        decoder.decode(damagedEntries.get('[Content_Types].xml'))
            .replace(COMMENTS_EXTENDED_TYPE, 'application/vnd.ms-word.commentsExtended+xml')
    ));
    await assert.rejects(() => validateDocxPackage(new MemoryZip(damagedEntries)), /commentsExtended|content type/i);
    const damaged = openDocx(zipDocx(damagedEntries));
    const repair = await damaged.applyOperations([
        { type: 'replace', target: { exactText: 'Second edited paragraph.' }, modified: 'Second repaired paragraph.' }
    ], { author: 'Editor', atomic: true });
    assert.equal(repair.status, 'ok', JSON.stringify(repair.error || repair.results));
    assert.ok(repair.artifactsChanged.includes('[Content_Types].xml'));
    assert.ok(part(damaged.toUint8Array(), '[Content_Types].xml').includes(COMMENTS_EXTENDED_TYPE));
    assert.ok(!part(damaged.toUint8Array(), '[Content_Types].xml').includes('application/vnd.ms-word.commentsExtended+xml'));
    const deleted = await doc.deleteComments({ ids: [root.id] });
    assert.equal(deleted.status, 'ok', JSON.stringify(deleted.error));
    assert.ok(deleted.artifactsChanged.includes('word/comments.xml'));
    assert.equal(doc.inspect().comments.length, 0);
}

async function testHeaderFooterAndFailClosed() {
    const doc = openDocx(packageBytes({ footer: true }));
    const footer = doc.inspect().headersFooters.find(item => item.kind === 'footer');
    assert.equal(footer?.path, 'word/footer1.xml');
    assert.equal(footer.paragraphs[0].text, 'Draft footer');
    const updated = await doc.applyOperations([
        { type: 'replace', target: { exactText: 'Alpha beta gamma.' }, modified: 'Alpha edited gamma.' },
        { type: 'replace', part: { kind: 'footer', type: 'default', section: 0 }, target: { exactText: 'Draft footer' }, modified: 'Final footer' }
    ], { author: 'Editor', atomic: true });
    assert.equal(updated.status, 'ok', JSON.stringify(updated.error || updated.results));
    assert.equal(updated.results[1].index, 2);
    assert.equal(updated.results[1].part, 'word/footer1.xml');
    assert.ok(updated.artifactsChanged.includes('word/footer1.xml'));
    const revisionIds = ['word/document.xml', 'word/footer1.xml']
        .flatMap(name => Array.from(part(doc.toUint8Array(), name).matchAll(/<w:(?:ins|del)\b[^>]*\bw:id="(\d+)"/g), match => match[1]));
    assert.ok(revisionIds.length >= 2);
    assert.equal(new Set(revisionIds).size, revisionIds.length, 'revision IDs must be unique across body and footer');
    assert.equal(doc.inspect().headersFooters[0].paragraphs[0].text, 'Final footer');
    const accepted = await doc.resolveRevisions('accept', { allAuthors: true });
    assert.equal(accepted.status, 'ok', JSON.stringify(accepted.error));
    assert.ok(accepted.artifactsChanged.includes('word/footer1.xml'));
    assert.doesNotMatch(part(doc.toUint8Array(), 'word/footer1.xml'), /<w:(?:ins|del)\b/);
    assert.equal(doc.inspect().headersFooters[0].paragraphs[0].text, 'Final footer');
    const before = doc.toUint8Array();
    const refused = await doc.applyOperations([
        { type: 'replace', target: { exactText: 'Second paragraph.' }, modified: 'Should roll back.' },
        { type: 'comment', part: { kind: 'footer' }, target: { exactText: 'Final footer' }, commentContent: 'Invalid' }
    ], { author: 'Editor', atomic: true });
    assert.equal(errorCode(refused), 'COMMENT_IN_HEADER_FOOTER');
    assert.equal(refused.written, false);
    assert.deepEqual(doc.toUint8Array(), before);
    const bodySearch = await doc.applyOperations([
        { type: 'replace', target: { exactText: 'Final footer' }, modified: 'Should not find footer' }
    ], { author: 'Editor', atomic: true });
    assert.ok(['TARGET_NOT_FOUND', 'ANCHOR_NOT_FOUND'].includes(errorCode(bodySearch)), JSON.stringify(bodySearch.error || bodySearch.results));
    assert.equal(bodySearch.written, false);
    assert.deepEqual(doc.toUint8Array(), before);
    const missingPart = await doc.applyOperations([
        { type: 'replace', part: 'word/footer99.xml', target: { exactText: 'Final footer' }, modified: 'No write' }
    ], { author: 'Editor', atomic: true });
    assert.equal(errorCode(missingPart), 'PART_NOT_FOUND');
    assert.equal(missingPart.written, false);
    assert.deepEqual(doc.toUint8Array(), before);

    const ambiguous = openDocx(packageBytes({ footer: true, secondFooter: true }));
    const ambiguousPart = await ambiguous.applyOperations([
        { type: 'replace', part: { kind: 'footer' }, target: { exactText: 'Draft footer' }, modified: 'No write' }
    ], { author: 'Editor', atomic: true });
    assert.equal(errorCode(ambiguousPart), 'PART_AMBIGUOUS');
    assert.equal(ambiguousPart.written, false);

    const fieldDoc = openDocx(packageBytes({ footer: true, field: true }));
    assert.equal(fieldDoc.inspect().headersFooters[0].hasFields, true);
    const fieldBefore = fieldDoc.toUint8Array();
    const fieldEdit = await fieldDoc.applyOperations([
        { type: 'replace', part: { kind: 'footer' }, target: { exactText: 'Page 1' }, modified: 'Page 2' }
    ], { author: 'Editor', atomic: true });
    assert.equal(errorCode(fieldEdit), 'FIELD_EDIT_REFUSED');
    assert.equal(fieldEdit.written, false);
    assert.deepEqual(fieldDoc.toUint8Array(), fieldBefore);
}

function testConsumerErrorPassThrough() {
    for (const code of ['COMMENT_NOT_FOUND', 'PARENT_ANCHOR_NOT_FOUND', 'COMMENT_IN_HEADER_FOOTER', 'PART_NOT_FOUND', 'PART_AMBIGUOUS', 'FIELD_EDIT_REFUSED', 'FUTURE_V081_ERROR']) {
        assert.throws(() => assertRedlineResult({
            status: 'error', hasChanges: false,
            error: { code, message: 'Compatibility error' }
        }, 'v0.8.1 compatibility test'), error => error instanceof RedlineOperationError && error.code === code);
    }
}

await testOpenInspectBatchAndResolution();
await testCommentedPackageAndThreadOperations();
await testHeaderFooterAndFailClosed();
testConsumerErrorPassThrough();
console.log('docx-redline-js v0.8.1 compatibility tests passed');
