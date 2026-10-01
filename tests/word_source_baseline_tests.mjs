import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { captureWordSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { assertSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/redline-plan.js';

const packageXml = buildDocumentFragmentPackage(
    '<w:p w14:paraId="ABCDEF01"><w:r><w:t>Before </w:t></w:r>'
    + '<w:del w:id="1" w:author="Tester"><w:r><w:delText>old</w:delText></w:r></w:del>'
    + '<w:ins w:id="2" w:author="Tester"><w:r><w:t>new</w:t></w:r></w:ins>'
    + '<w:r><w:br/><w:t>after</w:t></w:r></w:p>',
    { appendTrailingParagraph: false }
);
const baseline = captureWordSourceBaseline(packageXml);
assert.deepEqual(captureWordSourceBaseline(packageXml), baseline, 'Repeated inspection must produce stable targeting identities');
assert.equal(baseline.length, 1);
assert.equal(baseline[0].index, 1);
assert.equal(baseline[0].exactText, 'Before new\nafter', 'Baseline must use accepted-view text and canonical manual breaks');
assert.equal(typeof baseline[0].fingerprint, 'string');
assert.ok(baseline[0].fingerprint.length > 0);
assert.throws(() => captureWordSourceBaseline('<pkg:package/>'), /document.xml/);

const change = [{ operation: 'edit_paragraph', paragraphIndex: 1, newContent: 'Replanned text' }];
const idChanged = captureWordSourceBaseline(packageXml.replace('ABCDEF01', 'ABCDEF02'));
assert.equal(idChanged[0].exactText, baseline[0].exactText);
assert.throws(() => assertSourceBaseline(change, idChanged, baseline), error => {
    assert.equal(error.code, 'STALE_DOCUMENT_CONTEXT');
    assert.deepEqual(error.diagnostic, {
        reason: 'FINGERPRINT_MISMATCH', paragraphIndex: 1,
        paragraphIdChanged: true, tableContextChanged: false
    });
    assert.doesNotMatch(JSON.stringify(error.diagnostic), /Before|ABCDEF/);
    return true;
});
assert.doesNotThrow(() => assertSourceBaseline(change, idChanged, idChanged), 'fresh inspection accepts its own unchanged identity');
const textChanged = captureWordSourceBaseline(packageXml.replace('Before ', 'Different '));
assert.throws(() => assertSourceBaseline(change, textChanged, baseline), error => error.diagnostic.reason === 'TEXT_MISMATCH');
assert.throws(() => assertSourceBaseline(change, baseline, []), error => error.diagnostic.reason === 'BASELINE_UNAVAILABLE');
console.log('PASS: canonical Word source baseline tests');
