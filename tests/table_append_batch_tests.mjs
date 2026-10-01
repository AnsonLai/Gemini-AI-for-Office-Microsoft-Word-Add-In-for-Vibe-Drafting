import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
    acceptTrackedChangesInOoxml,
    configureLogger,
    inspectDocumentParts,
    rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { zipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { verifyAnchor } from '../src/taskpane/modules/commands/change-validation.js';
import { prepareCanonicalBatch } from '../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';
import {
    captureWordSourceBaseline
} from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import {
    applyRedlineChangesToWordContext
} from '../src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js';
import {
    planRedlineBatchOperations,
    planRedlineBatchOperationsWithMapping
} from '../src/taskpane/modules/docx-redline-js-integration/redline-plan.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encoder = new TextEncoder();
const decoder = new TextDecoder();
const author = 'Table Append Regression';
const table = '| Mountain | River | Forest |\n| --- | --- | --- |\n| Ocean | Valley | Canyon |\n| Meadow | Desert | Island |';
const original = [
    'Opening paragraph 1.', 'Opening paragraph 2.', 'Opening paragraph 3.',
    'Opening paragraph 4.', 'Opening paragraph 5.', 'Opening paragraph 6.',
    'Final paragraph stays unchanged.'
];
const revised = 'Closing paragraph stays unchanged.';

configureLogger({ log() {}, warn() {}, error() {} }, { level: 'silent' });

function xmlEscape(text) {
    return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function sourceDocumentXml(paragraphs = original) {
    const body = paragraphs.map(text => `<w:p><w:r><w:t>${xmlEscape(text)}</w:t></w:r></w:p>`).join('');
    return `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
}

function sourcePackage(paragraphs = original) {
    return buildDocumentFragmentPackage(
        paragraphs.map(text => `<w:p><w:r><w:t>${xmlEscape(text)}</w:t></w:r></w:p>`).join(''),
        { appendTrailingParagraph: false }
    );
}

function makeDocx(documentXml) {
    return zipDocx(new Map([
        ['[Content_Types].xml', `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`],
        ['_rels/.rels', `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
        ['word/_rels/document.xml.rels', `<Relationships xmlns="${REL}"/>`],
        ['word/document.xml', documentXml]
    ]));
}

function bodyStructure(documentXml) {
    const parser = new DOMParser();
    const document = parser.parseFromString(documentXml, 'application/xml');
    assert.equal(document.getElementsByTagName('parsererror').length, 0, 'document XML parses');
    const body = document.getElementsByTagNameNS(W, 'body')[0];
    assert.ok(body, 'document body exists');
    const bodyChildren = Array.from(body.childNodes).filter(node => node.nodeType === 1);
    const directText = node => Array.from(node.getElementsByTagNameNS(W, 't')).map(textNode => textNode.textContent).join('');
    const paragraphs = bodyChildren.filter(node => node.localName === 'p').map(directText);
    const tables = bodyChildren.filter(node => node.localName === 'tbl').map(node => {
        const rows = Array.from(node.getElementsByTagNameNS(W, 'tr'));
        const cells = rows.map(row => Array.from(row.getElementsByTagNameNS(W, 'tc')).map(directText));
        return { rows: cells.length, columns: cells[0]?.length || 0, cells };
    });
    return { paragraphs, tables };
}

function directParagraphs(documentXml) {
    return bodyStructure(documentXml).paragraphs;
}

function tablesIn(documentXml) {
    return bodyStructure(documentXml).tables;
}

function changePair(order = 'edit-first', { editAnchor = original[6], appendAnchor = original[6], replace = 'Closing' } = {}) {
    const edit = {
        operation: 'edit_paragraph', paragraphIndex: 7,
        anchorText: editAnchor,
        replacements: [{ find: 'Final', replace }]
    };
    const append = {
        operation: 'replace_paragraph', paragraphIndex: 8,
        anchorText: appendAnchor, content: table
    };
    return order === 'append-first' ? [append, edit] : [edit, append];
}

async function resolvePackageXml(documentXml, resolution) {
    const result = resolution === 'accept'
        ? acceptTrackedChangesInOoxml(documentXml, { allAuthors: true })
        : rejectTrackedChangesInOoxml(documentXml, { allAuthors: true });
    assert.notEqual(result.status, 'error', `${resolution} all: ${JSON.stringify(result.error)}`);
    return result.oxml;
}

function assertTable(tableShape) {
    assert.deepEqual(tableShape, {
        rows: 3,
        columns: 3,
        cells: [
            ['Mountain', 'River', 'Forest'],
            ['Ocean', 'Valley', 'Canyon'],
            ['Meadow', 'Desert', 'Island']
        ]
    });
}

async function prepareCoalesced(changes) {
    return prepareCanonicalBatch(sourcePackage(), source => (
        planRedlineBatchOperationsWithMapping(changes, source.paragraphs, {
            author, sourceDocumentXml: source.documentXml
        }).operations
    ), { author, existingRevisions: 'merge-same-author' });
}

function testCoalescesOnlyTheAdjacentEditAndAppend() {
    for (const order of ['edit-first', 'append-first']) {
        const changes = changePair(order, { editAnchor: 'Final paragraph', appendAnchor: original[6] });
        const result = planRedlineBatchOperationsWithMapping(changes, original.map((exactText, index) => ({ index: index + 1, exactText })), { author });
        assert.equal(result.operations.length, 1, `${order}: one engine operation`);
        assert.deepEqual(result.changeOperationIndexes, [1, 1], `${order}: both model changes map to receipt 1`);
        assert.equal(result.operations[0].modified, `${revised}\n${table}`);
        assert.equal(result.operations[0].targetRef, 'P7');
    }

    const duplicateAppends = planRedlineBatchOperationsWithMapping([
        { operation: 'replace_paragraph', paragraphIndex: 8, content: table },
        { operation: 'replace_paragraph', paragraphIndex: 8, content: table }
    ], original.map((exactText, index) => ({ index: index + 1, exactText })));
    assert.equal(duplicateAppends.operations.length, 2, 'duplicate appends remain separate and are not hidden');

    const rangeOverlap = planRedlineBatchOperationsWithMapping([
        { operation: 'replace_range', paragraphIndex: 7, endParagraphIndex: 7, content: revised },
        { operation: 'replace_paragraph', paragraphIndex: 8, content: table }
    ], original.map((exactText, index) => ({ index: index + 1, exactText })));
    assert.equal(rangeOverlap.operations.length, 2, 'arbitrary range overlap remains separate');
}

async function testPlainEditThenTablePreservesAcceptedAndRejectedViews() {
    for (const order of ['edit-first', 'append-first']) {
        const prepared = await prepareCoalesced(changePair(order, { editAnchor: 'Final paragraph', appendAnchor: original[6] }));
        assert.equal(prepared.status, 'ready', `${order}: ${JSON.stringify(prepared.result.error)}`);
        assert.equal(prepared.operations.length, 1);
        assert.equal(prepared.result.receipts.length, 1);
        assert.equal(prepared.result.receipts[0].committed, true);
        const acceptedXml = await resolvePackageXml(prepared.result.documentXml, 'accept');
        const rejectedXml = await resolvePackageXml(prepared.result.documentXml, 'reject');
        const accepted = bodyStructure(acceptedXml);
        const rejected = bodyStructure(rejectedXml);
        assert.deepEqual(accepted.paragraphs, [...original.slice(0, 6), revised]);
        assert.equal(accepted.tables.length, 1);
        assertTable(accepted.tables[0]);
        assert.deepEqual(rejected.paragraphs, original);
        assert.deepEqual(rejected.tables, []);
    }
}

function testLocalizedReplacementMustBeExactAndUnambiguous() {
    const paragraphs = original.map((exactText, index) => ({ index: index + 1, exactText }));
    const repeatedSource = [...original.slice(0, 6), 'Final Final paragraph.'];
    const ambiguous = [
        { operation: 'edit_paragraph', paragraphIndex: 7,
            replacements: [{ find: 'Final', replace: 'Closing' }] },
        { operation: 'replace_paragraph', paragraphIndex: 8, content: table }
    ];
    assert.throws(
        () => planRedlineBatchOperationsWithMapping(ambiguous, repeatedSource.map((exactText, index) => ({ index: index + 1, exactText }))),
        error => error.code === 'AMBIGUOUS_PATCH_SOURCE'
    );
    const missing = [
        { operation: 'edit_paragraph', paragraphIndex: 7,
            replacements: [{ find: 'Absent', replace: 'Closing' }] },
        { operation: 'replace_paragraph', paragraphIndex: 8, content: table }
    ];
    assert.throws(
        () => planRedlineBatchOperationsWithMapping(missing, paragraphs),
        error => error.code === 'PATCH_SOURCE_NOT_FOUND' || error.code === 'INVALID_OPERATION'
    );
    const selected = [
        { operation: 'edit_paragraph', paragraphIndex: 7,
            replacements: [{ find: 'Final', replace: 'Closing', occurrence: 2 }] },
        { operation: 'replace_paragraph', paragraphIndex: 8, content: table }
    ];
    const result = planRedlineBatchOperationsWithMapping(selected,
        repeatedSource.map((exactText, index) => ({ index: index + 1, exactText })));
    assert.equal(result.operations.length, 1);
    assert.equal(result.operations[0].modified, `Final Closing paragraph.\n${table}`);
}

function testFormattingRequestsFailClosed() {
    const changes = [
        { operation: 'edit_paragraph', paragraphIndex: 7,
            replacements: [{ find: original[6], replace: `++${original[6]}++` }] },
        { operation: 'replace_paragraph', paragraphIndex: 8, content: table }
    ];
    assert.throws(
        () => planRedlineBatchOperationsWithMapping(changes,
            original.map((exactText, index) => ({ index: index + 1, exactText })), { author }),
        error => error.code === 'UNSUPPORTED_TABLE_FORMATTING'
    );

    const fullReplacement = [{
        operation: 'replace_paragraph', paragraphIndex: 7,
        content: `++${original[6]}++\n${table}`
    }];
    assert.throws(
        () => planRedlineBatchOperationsWithMapping(fullReplacement,
            original.map((exactText, index) => ({ index: index + 1, exactText })), { author }),
        error => error.code === 'UNSUPPORTED_TABLE_FORMATTING'
    );

    const plainFullReplacement = planRedlineBatchOperations([{
        operation: 'replace_paragraph', paragraphIndex: 7,
        content: `${original[6]}\n${table}`
    }], original.map((exactText, index) => ({ index: index + 1, exactText })));
    assert.equal(plainFullReplacement.length, 1, 'plain paragraph-plus-table replacement remains available for Word characterization');
}

async function testAppendAnchorNeverAutocorrectsIntoItsSourceParagraph() {
    const texts = [...original];
    assert.deepEqual(verifyAnchor({ operation: 'replace_paragraph', paragraphIndex: 8, anchorText: original[6] }, texts), { ok: true });
    const mismatch = verifyAnchor({ operation: 'replace_paragraph', paragraphIndex: 8, anchorText: 'Different paragraph' }, texts);
    assert.equal(mismatch.ok, false);
    assert.equal(mismatch.reason, 'anchor_mismatch');
    const prefixed = planRedlineBatchOperations([
        { operation: 'edit_paragraph', paragraphIndex: 7, anchorText: 'Final paragraph',
            replacements: [{ find: 'Final', replace: 'Closing' }] },
        { operation: 'replace_paragraph', paragraphIndex: 8, anchorText: original[6], content: table }
    ], original.map((exactText, index) => ({ index: index + 1, exactText })));
    assert.equal(prefixed.length, 1, 'different valid prefixes still identify the same final paragraph');
}

async function testPriorSameAuthorUnderlineRefusesAppendBeforeAnyWrite() {
    const first = await prepareCanonicalBatch(sourcePackage(), source => planRedlineBatchOperations([
        { operation: 'edit_paragraph', paragraphIndex: 7, anchorText: original[6],
            replacements: [{ find: original[6], replace: `++${original[6]}++` }] }
    ], source.paragraphs), { author, existingRevisions: 'merge-same-author' });
    assert.equal(first.status, 'ready', JSON.stringify(first.result.error));
    const firstAccepted = await resolvePackageXml(first.result.documentXml, 'accept');
    const firstRejected = await resolvePackageXml(first.result.documentXml, 'reject');
    assert.match(firstAccepted, /<w:u\b/, 'the preceding underline is accepted');
    assert.doesNotMatch(firstRejected, /<w:u\b/, 'Reject All removes that underline revision');

    const baseline = captureWordSourceBaseline(first.insertionPayload);
    const writes = [];
    const body = {
        getOoxml: () => ({ value: first.insertionPayload }),
        insertOoxml: value => writes.push(value)
    };
    const outcome = await applyRedlineChangesToWordContext(
        { document: { body }, async sync() {} },
        [{ operation: 'replace_paragraph', paragraphIndex: 8, anchorText: original[6], content: table }],
        { author, sourceBaseline: baseline, onInfo() {}, onWarn() {} }
    );
    assert.equal(outcome.mutationOutcome, 'refused');
    assert.equal(outcome.error?.code, 'UNSUPPORTED_TABLE_FORMATTING');
    assert.equal(outcome.writeAttempted, false);
    assert.equal(outcome.written, false);
    assert.equal(writes.length, 0);
    assert.match(firstAccepted, /<w:u\b/, 'refusal leaves the previously accepted underline available');

    const replaceBody = {
        getOoxml: () => ({ value: first.insertionPayload }),
        insertOoxml: value => writes.push(value)
    };
    const replacementOutcome = await applyRedlineChangesToWordContext(
        { document: { body: replaceBody }, async sync() {} },
        [{ operation: 'replace_paragraph', paragraphIndex: 7, anchorText: original[6], content: `${original[6]}\n${table}` }],
        { author, sourceBaseline: baseline, onInfo() {}, onWarn() {} }
    );
    assert.equal(replacementOutcome.mutationOutcome, 'refused');
    assert.equal(replacementOutcome.error?.code, 'UNSUPPORTED_TABLE_FORMATTING');
    assert.equal(replacementOutcome.writeAttempted, false);
    assert.equal(writes.length, 0, 'same-author P7 full replacement is refused before insertion');

    const plainBody = { getOoxml: () => ({ value: sourcePackage() }), insertOoxml: value => writes.push(value) };
    const inlineOutcome = await applyRedlineChangesToWordContext(
        { document: { body: plainBody }, async sync() {} },
        [{ operation: 'replace_paragraph', paragraphIndex: 7, anchorText: original[6], content: `++${original[6]}++\n${table}` }],
        { author, onInfo() {}, onWarn() {} }
    );
    assert.equal(inlineOutcome.mutationOutcome, 'refused');
    assert.equal(inlineOutcome.error?.code, 'UNSUPPORTED_TABLE_FORMATTING');
    assert.equal(inlineOutcome.writeAttempted, false);
    assert.equal(writes.length, 0, 'new inline formatting plus table is refused before insertion');
}

async function testCoalescedReceiptMapsBackToBothInputChanges() {
    const changes = changePair('append-first');
    let capturedOperations;
    const result = await applyRedlineChangesToWordContext(
        { document: { body: {} } }, changes,
        {
            author,
            onInfo() {}, onWarn() {},
            batchRunner: async (_context, _scope, makeOperations) => {
                capturedOperations = makeOperations({ paragraphs: original.map((exactText, index) => ({ index: index + 1, exactText })) });
                return {
                    status: 'ok', hasChanges: true, written: true, writeAttempted: true,
                    receipts: [{ operationIndex: 1, committed: true, revisionItems: ['engine-owned'] }],
                    results: [{ index: 1, status: 'applied' }]
                };
            }
        }
    );
    assert.equal(capturedOperations.length, 1);
    assert.equal(result.changesApplied, 2);
    assert.deepEqual(result.skipped, []);
    assert.deepEqual(result.receipts, [{ operationIndex: 1, committed: true, revisionItems: ['engine-owned'] }]);

    const noFactory = await applyRedlineChangesToWordContext(
        { document: { body: {} } }, changes,
        {
            author, onInfo() {}, onWarn() {},
            batchRunner: async () => ({
                status: 'ok', hasChanges: true, written: true,
                receipts: [{ operationIndex: 1, committed: true }], results: []
            })
        }
    );
    assert.equal(noFactory.changesApplied, 1, 'custom runners that skip planning keep identity mapping');
    assert.equal(noFactory.skipped.length, 1);
}

async function buildHostFixtures(destination) {
    const changes = changePair('edit-first');
    const sourceXml = sourceDocumentXml();
    const prepared = await prepareCanonicalBatch(sourcePackage(), source => (
        planRedlineBatchOperationsWithMapping(changes, source.paragraphs, { author, sourceDocumentXml: source.documentXml }).operations
    ), { author, existingRevisions: 'merge-same-author' });
    assert.equal(prepared.status, 'ready', JSON.stringify(prepared.result.error));
    const acceptedXml = await resolvePackageXml(prepared.result.documentXml, 'accept');
    const rejectedXml = await resolvePackageXml(prepared.result.documentXml, 'reject');

    const underlineChange = [{ operation: 'edit_paragraph', paragraphIndex: 7, anchorText: original[6],
        replacements: [{ find: original[6], replace: `++${original[6]}++` }] }];
    const underlineOnly = await prepareCanonicalBatch(sourcePackage(), source => (
        planRedlineBatchOperations(underlineChange, source.paragraphs)
    ), { author, existingRevisions: 'merge-same-author' });
    assert.equal(underlineOnly.status, 'ready');
    // This raw consumer-core call deliberately bypasses the add-in's new guard
    // to preserve a separate installed-library known-failure fixture.
    const unguardedAppend = await prepareCanonicalBatch(underlineOnly.insertionPayload, source => (
        planRedlineBatchOperations([
            { operation: 'replace_paragraph', paragraphIndex: 8, anchorText: original[6], content: table }
        ], source.paragraphs)
    ), { author, existingRevisions: 'merge-same-author' });
    assert.equal(unguardedAppend.status, 'ready');
    const defectAccepted = await resolvePackageXml(unguardedAppend.result.documentXml, 'accept');
    const defectRejected = await resolvePackageXml(unguardedAppend.result.documentXml, 'reject');
    assert.doesNotMatch(defectAccepted, /<w:u\b/, 'engine reference confirms the known accepted-formatting loss');
    assert.match(defectRejected, /Final paragraph stays unchanged\./);

    mkdirSync(destination, { recursive: true });
    for (const [name, xml] of [
        ['table-append-source.docx', sourceXml],
        ['table-append-tracked.docx', prepared.result.documentXml],
        ['table-append-accepted.docx', acceptedXml],
        ['table-append-rejected.docx', rejectedXml],
        ['table-append-underline-known-failure-tracked.docx', unguardedAppend.result.documentXml],
        ['table-append-underline-known-failure-accepted.docx', defectAccepted],
        ['table-append-underline-known-failure-rejected.docx', defectRejected]
    ]) writeFileSync(resolve(destination, name), makeDocx(xml));

    const positiveNativeOperation = prepared.operations[0];
    const knownFailureNativeOperation = {
        type: 'redline',
        target: { index: 7, exactText: original[6] },
        targetRef: 'P7',
        structuredContent: true,
        modified: `++${original[6]}++\n${table}`
    };
    const emptyTables = [];
    const fullTable = [{ rows: 3, columns: 3, cells: [
        ['Mountain', 'River', 'Forest'], ['Ocean', 'Valley', 'Canyon'], ['Meadow', 'Desert', 'Island']
    ] }];
    const positiveExpected = {
        source: { paragraphs: original, tables: emptyTables },
        accepted: { paragraphs: [...original.slice(0, 6), revised, ''], tables: fullTable },
        rejected: { paragraphs: original, tables: emptyTables }
    };
    const failureExpected = {
        source: { paragraphs: original, tables: emptyTables },
        accepted: { paragraphs: [...original, ''], tables: fullTable },
        rejected: { paragraphs: original, tables: emptyTables }
    };
    const positiveInsertion = 'table-append-coalesced-insertion.xml';
    const knownFailureInsertion = 'table-append-underline-known-failure-insertion.xml';
    writeFileSync(resolve(destination, positiveInsertion), prepared.insertionPayload, 'utf8');
    writeFileSync(resolve(destination, knownFailureInsertion), unguardedAppend.insertionPayload, 'utf8');
    const manifest = {
        cases: [
            {
                name: 'table-append-coalesced-plain-edit',
                source: 'table-append-source.docx',
                tracked: 'table-append-tracked.docx',
                accepted: 'table-append-accepted.docx',
                rejected: 'table-append-rejected.docx',
                insertionXml: positiveInsertion,
                expectedBodyStructure: positiveExpected,
                expectedAcceptedText: [...original.slice(0, 6), revised, 'Mountain', 'River', 'Forest', 'Ocean', 'Valley', 'Canyon', 'Meadow', 'Desert', 'Island'].join('\n'),
                expectedRejectedText: original.join('\n'),
                expectedMinimumBodyRevisions: 1,
                expectedEngineReferencePackages: true,
                nativeOperations: [positiveNativeOperation]
            },
            {
                name: 'table-append-inline-format-known-library-defect',
                source: 'table-append-source.docx',
                tracked: 'table-append-underline-known-failure-tracked.docx',
                accepted: 'table-append-underline-known-failure-accepted.docx',
                rejected: 'table-append-underline-known-failure-rejected.docx',
                insertionXml: knownFailureInsertion,
                expectedBodyStructure: failureExpected,
                expectedAcceptedText: [...original, 'Mountain', 'River', 'Forest', 'Ocean', 'Valley', 'Canyon', 'Meadow', 'Desert', 'Island'].join('\n'),
                expectedRejectedText: original.join('\n'),
                expectedMinimumBodyRevisions: 1,
                expectedEngineReferencePackages: true,
                nativeOperations: [knownFailureNativeOperation],
                expectedFormatting: [{
                    sourceText: original[6], acceptedText: original[6],
                    bold: false, italic: false, underlineAccepted: true, underlineRejected: false
                }]
            }
        ]
    };
    writeFileSync(resolve(destination, 'manifest.json'), JSON.stringify(manifest, null, 2));
    return manifest;
}

testCoalescesOnlyTheAdjacentEditAndAppend();
await testPlainEditThenTablePreservesAcceptedAndRejectedViews();
testLocalizedReplacementMustBeExactAndUnambiguous();
testFormattingRequestsFailClosed();
await testAppendAnchorNeverAutocorrectsIntoItsSourceParagraph();
await testPriorSameAuthorUnderlineRefusesAppendBeforeAnyWrite();
await testCoalescedReceiptMapsBackToBothInputChanges();

const exportIndex = process.argv.indexOf('--export-host-dir');
if (exportIndex >= 0) {
    const destination = process.argv[exportIndex + 1];
    assert.ok(destination && !destination.startsWith('--'), '--export-host-dir requires a directory');
    const manifest = await buildHostFixtures(resolve(destination));
    console.log(`PASS: exported ${manifest.cases.length} table append Word fixtures to ${resolve(destination)}`);
}

console.log('PASS: table append coalescing, receipt mapping, fail-closed formatting guard, and accepted/rejected OOXML structure');
