import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { sanitizeChangeSet } from '../src/taskpane/modules/commands/change-validation.js';
import {
    applyRedlineChangesToWordContext,
    planRedlineBatchOperations
} from '../src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js';

const paragraphs = [
    { index: 1, ref: 'P1', text: 'First clause.', exactText: 'First clause.', paragraphId: 'AAAABBBB' },
    { index: 2, ref: 'P2', text: 'Second clause.', exactText: 'Second clause.', paragraphId: 'CCCCDDDD' },
    { index: 3, ref: 'P3', text: '', exactText: '', paragraphId: 'EEEEFFFF' }
];

function testPlanningUsesInitialParagraphDescriptors() {
    const operations = planRedlineBatchOperations([
        { operation: 'modify_text', paragraphIndex: 2, originalText: 'Second', replacementText: 'Updated' },
        { operation: 'replace_paragraph', paragraphIndex: 1, content: 'Rewritten first clause.' },
        { operation: 'edit_paragraph', paragraphIndex: 3, newContent: 'Filled blank.' }
    ], paragraphs);
    assert.deepEqual(operations.map(item => item.targetRef), ['P2', 'P1', 'P3']);
    assert.deepEqual(operations[0].replacements, [{ find: 'Second', replace: 'Updated' }]);
    assert.equal(operations[1].modified, 'Rewritten first clause.');
    assert.equal(operations[1].target.paragraphId, 'AAAABBBB');
    assert.equal(operations[2].target.exactText, '');
    assert.equal(operations[2].modified, 'Filled blank.');
    assert.ok(operations.every(item => item.structuredContent === true));
    const localized = planRedlineBatchOperations([
        { operation: 'edit_paragraph', paragraphIndex: 2,
          replacements: [{ find: 'Second', replace: 'Updated' }] }
    ], paragraphs);
    assert.deepEqual(localized[0].replacements, [{ find: 'Second', replace: 'Updated' }]);
    assert.equal(localized[0].modified, undefined);
}

function testRangeAndInsertionPlanning() {
    const range = planRedlineBatchOperations([
        { operation: 'replace_range', paragraphIndex: 1, endParagraphIndex: 2, content: 'Combined clause.' }
    ], paragraphs);
    assert.equal(range[0].targetEndRef, 'P2');
    assert.equal(range[0].modified, 'Combined clause.');

    const before = planRedlineBatchOperations([
        { operation: 'replace_range', paragraphIndex: 2, endParagraphIndex: 1, content: 'Inserted heading.' }
    ], paragraphs);
    assert.equal(before[0].modified, 'Inserted heading.\nSecond clause.');

    const append = planRedlineBatchOperations([
        { operation: 'replace_paragraph', paragraphIndex: 4, content: 'Appendix.' }
    ], paragraphs);
    assert.equal(append[0].targetRef, 'P3');
    assert.equal(append[0].modified, '\nAppendix.');
}

function testInvalidChangeFailsBeforeBatch() {
    assert.throws(() => planRedlineBatchOperations([
        { operation: 'replace_range', paragraphIndex: 1, endParagraphIndex: 20, content: 'Invalid' }
    ], paragraphs), error => error.code === 'TARGET_NOT_FOUND');
    assert.throws(() => planRedlineBatchOperations([
        { operation: 'modify_text', paragraphIndex: 1, replacementText: 'Invalid' }
    ], paragraphs), error => error.code === 'INVALID_OPERATION');
}

async function testCallerSubmitsOneBatchAndReportsRollback() {
    const body = {};
    const context = { document: { body } };
    const changes = [
        { operation: 'replace_paragraph', paragraphIndex: 1, content: 'First revised.' },
        { operation: 'modify_text', paragraphIndex: 2, originalText: 'Second', replacementText: 'Other' }
    ];
    let calls = 0;
    const batchRunner = async (actualContext, scope, makeOperations, options) => {
        calls += 1;
        assert.equal(actualContext, context);
        assert.equal(scope, body);
        assert.equal(options.author, 'Editor');
        const operations = await makeOperations({ paragraphs });
        assert.equal(operations.length, 2);
        assert.deepEqual(operations.map(item => item.targetRef), ['P1', 'P2']);
        return {
            status: 'error', written: false, rolledBack: true,
            error: { code: 'BATCH_OPERATION_FAILED', message: 'Atomic batch refused' },
            receipts: [
                { operationIndex: 1, committed: false, finalDisposition: 'rolled_back' },
                { operationIndex: 2, committed: false, finalDisposition: 'refused' }
            ],
            results: [{ index: 1, status: 'applied' }, { index: 2, status: 'error', error: { code: 'TARGET_NOT_FOUND', message: 'Missing target' } }]
        };
    };
    const result = await applyRedlineChangesToWordContext(context, changes, {
        author: 'Editor', batchRunner, onInfo: () => {}, onWarn: () => {}
    });
    assert.equal(calls, 1);
    assert.equal(result.changesApplied, 0);
    assert.equal(result.skipped.length, 2);
    assert.equal(result.skipped[1].code, 'TARGET_NOT_FOUND');
    assert.ok(result.skipped.every(item => item.rolledBack === true));
}

async function testEmptyParagraphIsHandledByEngine() {
    const xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr/></w:body></w:document>';
    const operations = planRedlineBatchOperations([
        { operation: 'edit_paragraph', paragraphIndex: 1, newContent: 'Filled blank.' }
    ], [{ index: 1, ref: 'P1', text: '', exactText: '' }]);
    const result = await applyOperationsToDocumentXml(xml, operations, 'Editor', null, {
        atomic: true, structuredContent: true, pairReplacements: true
    });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.match(result.documentXml, /Filled blank\./);
}

async function testEndToEndSingleWordWrite() {
    const original = buildDocumentFragmentPackage(
        '<w:p><w:r><w:t>First clause.</w:t></w:r></w:p>'
        + '<w:p><w:r><w:t>Second clause.</w:t></w:r></w:p>',
        { appendTrailingParagraph: false }
    );
    const writes = [];
    let reads = 0;
    const body = {
        getOoxml() { reads++; return { value: original }; },
        insertOoxml(xml) { writes.push(xml); }
    };
    const context = { document: { body }, async sync() {} };
    const proposed = [
        { operation: 'edit_paragraph', paragraphIndex: 2, anchorText: 'Second clause.',
          replacements: [{ find: 'Second', replace: 'Updated' }] },
        { operation: 'replace_paragraph', paragraphIndex: 1, content: 'Rewritten first clause.' }
    ];
    const validated = sanitizeChangeSet(proposed, 2, ['First clause.', 'Second clause.']);
    assert.equal(validated.rejected.length, 0);
    const outcome = await applyRedlineChangesToWordContext(context, validated.changes,
        { author: 'Editor', onInfo: () => {}, onWarn: () => {} });
    assert.equal(outcome.batchResult?.status, 'ok', JSON.stringify(outcome.batchResult?.error));
    assert.equal(outcome.changesApplied, 2);
    assert.equal(reads, 1);
    assert.equal(writes.length, 1);
    assert.match(writes[0], /Updated/);
    assert.match(writes[0], /Rewritten/);
}

async function testRangeAndAppendAgainstEngine() {
    const xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:body><w:p><w:r><w:t>First clause.</w:t></w:r></w:p>'
        + '<w:p><w:r><w:t>Second clause.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>';
    const source = paragraphs.slice(0, 2).map(({ index, ref, text, exactText }) => ({ index, ref, text, exactText }));
    const range = planRedlineBatchOperations([
        { operation: 'replace_range', paragraphIndex: 1, endParagraphIndex: 2, content: 'Combined clause.' }
    ], source);
    const rangeResult = await applyOperationsToDocumentXml(xml, range, 'Editor', null, {
        atomic: true, structuredContent: true, pairReplacements: true
    });
    assert.equal(rangeResult.status, 'ok', JSON.stringify(rangeResult.error));
    assert.match(rangeResult.documentXml, /<w:t>Combined<\/w:t>/);
    assert.match(rangeResult.documentXml, /clause\./);

    const append = planRedlineBatchOperations([
        { operation: 'replace_paragraph', paragraphIndex: 3, content: 'Appendix.' }
    ], source);
    const appendResult = await applyOperationsToDocumentXml(xml, append, 'Editor', null, {
        atomic: true, structuredContent: true, pairReplacements: true
    });
    assert.equal(appendResult.status, 'ok', JSON.stringify(appendResult.error));
    assert.match(appendResult.documentXml, /Appendix\./);
}

async function testRepeatedFindOccurrenceAgainstEngine() {
    const xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:body><w:p><w:r><w:t>term and term</w:t></w:r></w:p><w:sectPr/></w:body></w:document>';
    const changes = [{ operation: 'edit_paragraph', paragraphIndex: 1,
        replacements: [{ find: 'term', replace: 'phrase', occurrence: 2 }] }];
    const sanitized = sanitizeChangeSet(changes, 1, ['term and term']);
    assert.equal(sanitized.rejected.length, 0);
    const operations = planRedlineBatchOperations(sanitized.changes,
        [{ index: 1, exactText: 'term and term' }]);
    const result = await applyOperationsToDocumentXml(xml, operations, 'Editor', null, {
        atomic: true, structuredContent: true, pairReplacements: true
    });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.match(result.documentXml, /phrase/);
}

testPlanningUsesInitialParagraphDescriptors();
testRangeAndInsertionPlanning();
testInvalidChangeFailsBeforeBatch();
await testCallerSubmitsOneBatchAndReportsRollback();
await testEmptyParagraphIsHandledByEngine();
await testEndToEndSingleWordWrite();
await testRangeAndAppendAgainstEngine();
await testRepeatedFindOccurrenceAgainstEngine();
console.log('word redline batch plan tests passed');
