import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { captureWordSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';

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
console.log('PASS: canonical Word source baseline tests');
