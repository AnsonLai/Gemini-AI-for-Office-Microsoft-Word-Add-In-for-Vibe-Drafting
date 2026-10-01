import '../../tests/setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import {
    acceptTrackedChangesInOoxml,
    inspectDocumentParts,
    rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { prepareCanonicalBatch } from '../../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';
import { planRedlineBatchOperations } from '../../src/taskpane/modules/docx-redline-js-integration/redline-plan.js';

const author = 'Table Append Library Reproducer';
const table = '| Mountain | River | Forest |\n| --- | --- | --- |\n| Ocean | Valley | Canyon |\n| Meadow | Desert | Island |';
const sourceParagraphs = [
    'Opening paragraph 1.', 'Opening paragraph 2.', 'Opening paragraph 3.',
    'Opening paragraph 4.', 'Opening paragraph 5.', 'Opening paragraph 6.',
    'Final paragraph stays unchanged.'
];
const originalFinal = sourceParagraphs[6];
const sourcePackage = buildDocumentFragmentPackage(sourceParagraphs.map(text => (
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
)).join(''), { appendTrailingParagraph: false });

function revisionState(xml) {
    const inspection = inspectDocumentParts({ documentXml: xml });
    assert.equal(inspection.status, 'ok', JSON.stringify(inspection.errors));
    const paragraphs = inspection.paragraphs.filter(paragraph => !paragraph.inTable).map(paragraph => paragraph.exactText);
    const cells = inspection.paragraphs.filter(paragraph => paragraph.inTable).map(paragraph => paragraph.exactText);
    const tableCount = (xml.match(/<w:tbl\b/g) || []).length;
    return { paragraphs, cells, tableCount, underline: /<w:u\b/.test(xml) };
}

async function resolve(xml, action) {
    const result = action === 'accept'
        ? acceptTrackedChangesInOoxml(xml, { allAuthors: true })
        : rejectTrackedChangesInOoxml(xml, { allAuthors: true });
    assert.notEqual(result.status, 'error', `${action}: ${JSON.stringify(result.error)}`);
    return result.oxml;
}

// Turn 1: request underline on P7 and accept the generated same-author change.
const underline = await prepareCanonicalBatch(sourcePackage, source => planRedlineBatchOperations([
    {
        operation: 'edit_paragraph', paragraphIndex: 7, anchorText: originalFinal,
        replacements: [{ find: originalFinal, replace: `++${originalFinal}++` }]
    }
], source.paragraphs), { author, existingRevisions: 'merge-same-author' });
assert.equal(underline.status, 'ready', JSON.stringify(underline.result.error));
const underlineAccepted = revisionState(await resolve(underline.result.documentXml, 'accept'));
const underlineRejected = revisionState(await resolve(underline.result.documentXml, 'reject'));
assert.equal(underlineAccepted.underline, true, 'the initial underline operation works');
assert.equal(underlineRejected.underline, false);
assert.deepEqual(underlineRejected.paragraphs, sourceParagraphs);

// Turn 2: model P8 append intent, planned against the unchanged accepted-view P7.
// This deliberately omits sourceDocumentXml from the planner options to exercise
// the installed engine behind the add-in's safety refusal.
const append = await prepareCanonicalBatch(underline.insertionPayload, source => (
    planRedlineBatchOperations([
        { operation: 'replace_paragraph', paragraphIndex: 8, anchorText: originalFinal, content: table }
    ], source.paragraphs)
), { author, existingRevisions: 'merge-same-author' });
assert.equal(append.status, 'ready', JSON.stringify(append.result.error));
const appendAccepted = revisionState(await resolve(append.result.documentXml, 'accept'));
const appendRejected = revisionState(await resolve(append.result.documentXml, 'reject'));
assert.deepEqual(appendAccepted.paragraphs, sourceParagraphs);
assert.equal(appendAccepted.tableCount, 1);
assert.deepEqual(appendAccepted.cells, [
    'Mountain', 'River', 'Forest', 'Ocean', 'Valley', 'Canyon', 'Meadow', 'Desert', 'Island'
]);
assert.equal(appendAccepted.underline, false, 'known failure: same-author append discarded accepted underline');
assert.deepEqual(appendRejected.paragraphs, sourceParagraphs);
assert.equal(appendRejected.tableCount, 0);
assert.equal(appendRejected.underline, false);

// Independent second failure: asking for new Markdown underline and a table in
// one replacement string reports success and emits the table, but loses underline.
const mixed = await prepareCanonicalBatch(sourcePackage, source => ([{
    type: 'redline',
    target: { index: 7, exactText: originalFinal },
    targetRef: 'P7',
    structuredContent: true,
    modified: `++${originalFinal}++\n${table}`
}]), { author, existingRevisions: 'merge-same-author' });
assert.equal(mixed.status, 'ready', JSON.stringify(mixed.result.error));
const mixedAccepted = revisionState(await resolve(mixed.result.documentXml, 'accept'));
assert.equal(mixedAccepted.underline, false, 'known failure: requested new underline is absent');
assert.equal(mixedAccepted.tableCount, 1, 'the structural table is retained while its inline formatting request is lost');
assert.deepEqual(mixedAccepted.cells, [
    'Mountain', 'River', 'Forest', 'Ocean', 'Valley', 'Canyon', 'Meadow', 'Desert', 'Island'
]);

console.log(JSON.stringify({
    implementation: '@ansonlai/docx-redline-js 0.8.3 installed package',
    sameAuthorTrackedAppend: {
        firstUnderlineAccepted: underlineAccepted.underline,
        appendedTableAccepted: appendAccepted.tableCount === 1,
        acceptedUnderlineAfterAppend: appendAccepted.underline,
        rejectAllRestoresSevenParagraphs: JSON.stringify(appendRejected.paragraphs) === JSON.stringify(sourceParagraphs),
        rejectAllRemovesTable: appendRejected.tableCount === 0
    },
    combinedNewUnderlineAndTable: {
        engineStatus: mixed.result.status,
        acceptedUnderline: mixedAccepted.underline,
        acceptedTableCount: mixedAccepted.tableCount,
        acceptedTableCells: mixedAccepted.cells
    },
    outcome: 'KNOWN_LIBRARY_DEFECT_REPRODUCED'
}, null, 2));
