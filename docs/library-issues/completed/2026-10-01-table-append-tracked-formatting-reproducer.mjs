import '../../../tests/setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import {
    acceptTrackedChangesInOoxml,
    inspectDocumentParts,
    rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const author = 'Table Append Library Reproducer';
const table = '| Mountain | River | Forest |\n| --- | --- | --- |\n| Ocean | Valley | Canyon |\n| Meadow | Desert | Island |';
const sourceParagraphs = [
    'Opening paragraph 1.', 'Opening paragraph 2.', 'Opening paragraph 3.',
    'Opening paragraph 4.', 'Opening paragraph 5.', 'Opening paragraph 6.',
    'Final paragraph stays unchanged.'
];
const originalFinal = sourceParagraphs[6];
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const sourceDocumentXml = `<w:document xmlns:w="${W}"><w:body>${sourceParagraphs.map(text => (
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
)).join('')}<w:sectPr/></w:body></w:document>`;

function revisionState(xml) {
    const inspection = inspectDocumentParts({ documentXml: xml });
    assert.equal(inspection.status, 'ok', JSON.stringify(inspection.errors));
    const paragraphs = inspection.paragraphs.filter(paragraph => !paragraph.inTable).map(paragraph => paragraph.exactText);
    const cells = inspection.paragraphs.filter(paragraph => paragraph.inTable).map(paragraph => paragraph.exactText);
    const tableCount = (xml.match(/<w:tbl\b/g) || []).length;
    return { paragraphs, cells, tableCount, underline: /<w:u\b/.test(xml) };
}

function resolve(xml, action) {
    const result = action === 'accept'
        ? acceptTrackedChangesInOoxml(xml, { allAuthors: true })
        : rejectTrackedChangesInOoxml(xml, { allAuthors: true });
    assert.notEqual(result.status, 'error', `${action}: ${JSON.stringify(result.error)}`);
    return result.oxml;
}

async function replaceFinalParagraph(documentXml, modified, existingRevisions) {
    const inspection = inspectDocumentParts({ documentXml });
    assert.equal(inspection.status, 'ok');
    const target = inspection.paragraphs.find(paragraph => paragraph.exactText === originalFinal);
    assert.ok(target, 'the final paragraph is resolved from the current canonical source');
    return applyOperationsToDocumentXml(documentXml, [{
        type: 'redline', target, modified, structuredContent: true
    }], author, null, {
        atomic: true,
        structuredContent: true,
        pairReplacements: true,
        generateRedlines: true,
        ...(existingRevisions ? { existingRevisions } : {})
    });
}

// Turn 1: add a tracked underline to the final paragraph.
const underline = await replaceFinalParagraph(sourceDocumentXml, `++${originalFinal}++`);
assert.equal(underline.status, 'ok', JSON.stringify(underline.error));
const underlineAccepted = revisionState(resolve(underline.documentXml, 'accept'));
const underlineRejected = revisionState(resolve(underline.documentXml, 'reject'));
assert.equal(underlineAccepted.underline, true);
assert.equal(underlineRejected.underline, false);
assert.deepEqual(underlineRejected.paragraphs, sourceParagraphs);

// Turn 2: append a Markdown table after the same paragraph with the author's
// pending underline. This is the public-library regression guard for 0.8.4.
const append = await replaceFinalParagraph(underline.documentXml, `${originalFinal}\n${table}`, 'merge-same-author');
assert.equal(append.status, 'ok', JSON.stringify(append.error));
const appendAccepted = revisionState(resolve(append.documentXml, 'accept'));
const appendRejected = revisionState(resolve(append.documentXml, 'reject'));
assert.deepEqual(appendAccepted.paragraphs, [...sourceParagraphs, '']);
assert.equal(appendAccepted.tableCount, 1);
assert.deepEqual(appendAccepted.cells, [
    'Mountain', 'River', 'Forest', 'Ocean', 'Valley', 'Canyon', 'Meadow', 'Desert', 'Island'
]);
assert.equal(appendAccepted.underline, true, '0.8.4 keeps accepted underline after the same-author table append');
assert.deepEqual(appendRejected.paragraphs, sourceParagraphs);
assert.equal(appendRejected.tableCount, 0);
assert.equal(appendRejected.underline, false);

// Independent same-batch case: requested underline and table survive together.
const mixed = await replaceFinalParagraph(sourceDocumentXml, `++${originalFinal}++\n${table}`);
assert.equal(mixed.status, 'ok', JSON.stringify(mixed.error));
const mixedAccepted = revisionState(resolve(mixed.documentXml, 'accept'));
assert.equal(mixedAccepted.underline, true, '0.8.4 retains the requested underline');
assert.equal(mixedAccepted.tableCount, 1);
assert.deepEqual(mixedAccepted.cells, [
    'Mountain', 'River', 'Forest', 'Ocean', 'Valley', 'Canyon', 'Meadow', 'Desert', 'Island'
]);

console.log(JSON.stringify({
    implementation: '@ansonlai/docx-redline-js public standalone-runner API',
    sameAuthorTrackedAppend: {
        firstUnderlineAccepted: underlineAccepted.underline,
        appendedTableAccepted: appendAccepted.tableCount === 1,
        acceptedUnderlineAfterAppend: appendAccepted.underline,
        rejectAllRestoresSevenParagraphs: JSON.stringify(appendRejected.paragraphs) === JSON.stringify(sourceParagraphs),
        rejectAllRemovesTable: appendRejected.tableCount === 0
    },
    combinedNewUnderlineAndTable: {
        engineStatus: mixed.status,
        acceptedUnderline: mixedAccepted.underline,
        acceptedTableCount: mixedAccepted.tableCount,
        acceptedTableCells: mixedAccepted.cells
    },
    outcome: '0.8.4_REGRESSION_GUARD_PASSED'
}, null, 2));
