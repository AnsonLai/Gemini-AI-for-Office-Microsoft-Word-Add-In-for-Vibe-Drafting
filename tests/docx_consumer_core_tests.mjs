import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import {
    buildDocumentFragmentPackage
} from '@ansonlai/docx-redline-js/services/package-builder.js';
import * as portable from '../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';
import * as legacyWordRunner from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import * as legacyRedlineRunner from '../src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js';

const corePath = fileURLToPath(new URL('../src/taskpane/modules/docx-redline-js-integration/consumer-core.js', import.meta.url));
const coreSource = readFileSync(corePath, 'utf8');
assert.doesNotMatch(coreSource, /from\s+['"][^'"]*(?:word-ooxml|word-operation-runner|word-redline-runner|taskpane\.js|commands\/agentic-tools)[^'"]*['"]/, 'portable core must not import Word transport or UI modules');
assert.doesNotMatch(coreSource, /from\s+['"](?:node:)?fs(?:\/promises)?['"]/, 'portable core must not import filesystem APIs');
assert.equal(typeof globalThis.Word, 'undefined', 'the consumer core test starts without Office.js globals');

for (const name of [
    'prepareCanonicalBatch', 'captureSourceBaseline', 'applySharedOperationToParagraphOoxml',
    'applySharedOperationToScopeOoxml', 'planRedlineBatchOperations', 'assertSourceBaseline',
    'assertRedlineResult', 'prepareOperationInput', 'planAgenticListOperations', 'validateListRequest'
]) {
    assert.equal(typeof portable[name], 'function', `portable entrypoint exports ${name}`);
}

assert.equal(legacyWordRunner.captureWordSourceBaseline, portable.captureWordSourceBaseline,
    'the prior Word baseline export remains available as a compatibility alias');
assert.equal(legacyWordRunner.applySharedOperationToScopeOoxml, portable.applySharedOperationToScopeOoxml,
    'the prior Word runner still re-exports pure OOXML preparation');
assert.equal(legacyRedlineRunner.planRedlineBatchOperations, portable.planRedlineBatchOperations,
    'the prior redline runner still re-exports the pure operation mapper');

const packageXml = buildDocumentFragmentPackage(
    '<w:p><w:r><w:t>Source sentence.</w:t></w:r></w:p>',
    { appendTrailingParagraph: false }
);
const baseline = portable.captureSourceBaseline(packageXml);
assert.equal(baseline.length, 1);
assert.equal(baseline[0].exactText, 'Source sentence.');
const listRequest = { afterParagraphIndex: 1, text: 'Next item', indentLevel: 0 };
assert.equal(portable.validateListRequest(listRequest, 1).valid, true,
    'the portable entrypoint exposes pure list request validation');
const listOperation = portable.planAgenticListOperations({ paragraphs: [{
    index: 1,
    exactText: 'Existing item',
    list: { numId: '1', level: 0, format: 'bullet' }
}] }, listRequest)[0];
assert.equal(listOperation.target.exactText, 'Existing item');
assert.match(listOperation.modified, /Next item/);

const changes = [{
    operation: 'edit_paragraph', paragraphIndex: 1,
    replacements: [{ find: 'Source', replace: 'Portable' }]
}];
const prepared = await portable.prepareCanonicalBatch(packageXml,
    source => portable.planRedlineBatchOperations(changes, source.paragraphs),
    { author: 'Portable Core Test' });
assert.equal(prepared.status, 'ready', JSON.stringify(prepared.result?.error));
assert.equal(prepared.source.paragraphs[0].exactText, 'Source sentence.');
assert.equal(prepared.operations.length, 1);
assert.equal(typeof prepared.insertionPayload, 'string');
assert.equal(typeof prepared.result.documentXml, 'string');
const accepted = acceptTrackedChangesInOoxml(prepared.result.documentXml, { allAuthors: true });
const rejected = rejectTrackedChangesInOoxml(prepared.result.documentXml, { allAuthors: true });
assert.equal(accepted.status, undefined, JSON.stringify(accepted.error));
assert.equal(rejected.status, undefined, JSON.stringify(rejected.error));
assert.equal(portable.captureSourceBaseline(accepted.oxml)[0].exactText, 'Portable sentence.');
assert.equal(portable.captureSourceBaseline(rejected.oxml)[0].exactText, 'Source sentence.');

assert.throws(() => portable.assertSourceBaseline(
    [{ operation: 'edit_paragraph', paragraphIndex: 1 }],
    [{ ...baseline[0], exactText: 'Concurrent sentence.' }], baseline
), error => error.code === 'STALE_DOCUMENT_CONTEXT');

const refused = await portable.prepareCanonicalBatch(packageXml, [{
    type: 'redline', targetRef: 'P99', target: 'Missing source', modified: 'Should refuse'
}], { author: 'Portable Core Test' });
assert.equal(refused.status, 'refused');
assert.equal(refused.result.status, 'error');
assert.equal(refused.insertionPayload, undefined, 'a refused engine result must not prepare a host payload');
assert.equal(typeof globalThis.Word, 'undefined', 'portable import and execution do not require or install Word globals');

console.log('PASS: portable DOCX consumer core tests');
