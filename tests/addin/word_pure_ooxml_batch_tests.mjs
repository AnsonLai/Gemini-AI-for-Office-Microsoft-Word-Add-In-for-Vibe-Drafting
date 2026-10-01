import '../setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { executePureOoxmlBatch } from '../../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { insertOoxmlWithRangeFallback } from '../../src/taskpane/modules/docx-redline-js-integration/word-ooxml.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const original = buildDocumentFragmentPackage(
    `<w:p><w:r><w:t>Alpha beta gamma.</w:t></w:r></w:p>`
    + `<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>`,
    { appendTrailingParagraph: false }
);

function fixture(xml = original) {
    const writes = [];
    let reads = 0;
    let syncs = 0;
    const body = {
        getOoxml() { reads++; return { value: xml }; },
        insertOoxml(value, mode) { writes.push({ value, mode }); }
    };
    const context = { async sync() { syncs++; } };
    return { body, context, writes, get reads() { return reads; }, get syncs() { return syncs; } };
}

const successful = fixture();
let callbackCalls = 0;
const result = await executePureOoxmlBatch(successful.context, successful.body, source => {
    callbackCalls++;
    assert.equal(source.paragraphs.length, 2);
    assert.deepEqual(source.paragraphs.map(paragraph => paragraph.ref), ['P1', 'P2']);
    assert.equal(source.paragraphs[0].text, 'Alpha beta gamma.');
    return [
        { type: 'replace', target: { exactText: source.paragraphs[1].text }, modified: 'Second updated paragraph.' },
        { type: 'replace', target: { exactText: source.paragraphs[0].text }, replacements: [{ find: 'beta', replace: 'BETA' }] }
    ];
}, { author: 'Batch test' });
assert.equal(callbackCalls, 1);
assert.equal(successful.reads, 1);
assert.equal(successful.writes.length, 1);
assert.equal(successful.syncs, 2);
assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
assert.equal(result.written, true);
assert.equal(result.receipts.length, 2);
assert.match(successful.writes[0].value, /BETA/);
assert.match(successful.writes[0].value, /Second /);
assert.match(successful.writes[0].value, /updated /);
assert.match(successful.writes[0].value, /<pkg:package/);

const failed = fixture();
const refusal = await executePureOoxmlBatch(failed.context, failed.body, [
    { type: 'replace', target: { exactText: 'Alpha beta gamma.' }, modified: 'This would change.' },
    { type: 'replace', target: { exactText: 'Missing target.' }, modified: 'Cannot change.' }
], { author: 'Batch test' });
assert.equal(refusal.status, 'error');
assert.equal(refusal.rolledBack, true);
assert.equal(refusal.written, false);
assert.equal(failed.reads, 1);
assert.equal(failed.writes.length, 0);
assert.equal(failed.syncs, 1);
assert.equal(refusal.receipts.length, 2);

const stylePart = `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData><w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Custom"/></w:styles></pkg:xmlData></pkg:part>`;
const styled = fixture(original.replace('</pkg:package>', `${stylePart}</pkg:package>`));
const commentResult = await executePureOoxmlBatch(styled.context, styled.body, [
    { type: 'comment', target: { exactText: 'Alpha beta gamma.' }, textToComment: 'beta', commentContent: 'Please review.' }
], { author: 'Batch test' });
assert.equal(commentResult.status, 'ok', JSON.stringify(commentResult.error || commentResult.results));
assert.equal(commentResult.written, true);
assert.equal(styled.writes.length, 1);
assert.match(styled.writes[0].value, /word\/styles\.xml/);
assert.match(styled.writes[0].value, /styleId="Custom"/);
assert.match(styled.writes[0].value, /word\/comments\.xml/);
assert.match(styled.writes[0].value, /relationships\/comments/);
assert.match(styled.writes[0].value, /Please review/);

const rejectedWrite = fixture();
rejectedWrite.body.insertOoxml = () => { throw Object.assign(new Error('Word rejected package'), { code: 'FUTURE_WORD_ERROR' }); };
const writeError = await executePureOoxmlBatch(rejectedWrite.context, rejectedWrite.body, [
    { type: 'replace', target: { exactText: 'Alpha beta gamma.' }, modified: 'Updated.' }
], { author: 'Batch test' });
assert.equal(writeError.status, 'error');
assert.equal(writeError.error.code, 'WORD_OOXML_WRITE_FAILED');
assert.equal(writeError.written, false);
assert.equal(writeError.writeAttempted, true);
assert.equal(writeError.hostError.code, 'FUTURE_WORD_ERROR');
assert.equal(writeError.mutationOutcome, 'indeterminate');
assert.equal(writeError.receipts[0].committed, true, 'preserve engine commitment to prepared XML');
assert.deepEqual(writeError.receipts, writeError.engineResult.receipts);

const syncFailure = fixture();
let syncCalls = 0;
syncFailure.context.sync = async () => {
    if (++syncCalls === 2) throw new Error('Connection lost after insertion queued');
};
const uncertain = await executePureOoxmlBatch(syncFailure.context, syncFailure.body, [
    { type: 'replace', target: { exactText: 'Alpha beta gamma.' }, modified: 'Updated.' }
]);
assert.equal(syncFailure.writes.length, 1);
assert.equal(uncertain.written, false);
assert.equal(uncertain.writeAttempted, true);
assert.equal(uncertain.mutationOutcome, 'indeterminate');

const unknownError = { code: 'FUTURE_LIBRARY_ERROR', message: 'Future failure', details: { target: 42 } };
const engineRefusal = { status: 'error', rolledBack: true, error: unknownError, receipts: [{ finalDisposition: 'future_disposition' }] };
const unknownFixture = fixture();
const unknown = await executePureOoxmlBatch(unknownFixture.context, unknownFixture.body, [{ type: 'replace' }], {
    runner: async () => engineRefusal
});
assert.equal(unknown.error, unknownError);
assert.equal(unknown.receipts, engineRefusal.receipts);
assert.equal(unknown.writeAttempted, false);
assert.equal(unknown.mutationOutcome, 'rolled_back');
assert.equal(unknownFixture.writes.length, 0);

const noChange = fixture();
const noop = await executePureOoxmlBatch(noChange.context, noChange.body, [
    { type: 'replace', target: { exactText: 'Alpha beta gamma.' }, modified: 'Alpha beta gamma.' }
]);
assert.equal(noop.mutationOutcome, 'noop');
assert.equal(noop.writeAttempted, false);
assert.equal(noChange.writes.length, 0);

let insertions = 0;
let rangeReads = 0;
await assert.rejects(insertOoxmlWithRangeFallback({
    insertOoxml() { insertions++; },
    getRange() { rangeReads++; return { insertOoxml() { insertions++; } }; }
}, original, 'Replace', { async sync() { throw Object.assign(new Error('Unconfirmed write'), { code: 'GeneralException' }); } }));
assert.equal(insertions, 1, 'never replay after an unconfirmed host mutation');
assert.equal(rangeReads, 0);

let trackingSyncs = 0;
const restoreFailure = fixture();
restoreFailure.context.document = { changeTrackingMode: 'TrackAll', load() {} };
restoreFailure.context.sync = async () => {
    if (++trackingSyncs === 4) throw new Error('Cannot restore tracking');
};
globalThis.Word = { ChangeTrackingMode: { off: 'Off' } };
const afterWriteFailure = await executePureOoxmlBatch(restoreFailure.context, restoreFailure.body, [
    { type: 'replace', target: { exactText: 'Alpha beta gamma.' }, modified: 'Updated.' }
], { disableNativeTracking: true, baseTrackingMode: 'TrackAll' });
assert.equal(afterWriteFailure.written, true);
assert.equal(afterWriteFailure.writeAttempted, true);
assert.equal(afterWriteFailure.mutationOutcome, 'applied_with_host_error');
assert.equal(restoreFailure.writes.length, 1);
delete globalThis.Word;

console.log('PASS: pure OOXML batch tests');
