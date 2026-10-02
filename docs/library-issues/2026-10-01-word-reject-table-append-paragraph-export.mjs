// Manual-only fixture exporter for the open Word Reject All table-append issue.
// It uses only public library APIs; it does not call a model or provider.
import '../../tests/setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
    acceptTrackedChangesInOoxml,
    inspectDocumentParts,
    rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';
import { zipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const AUTHOR = 'Word Reject All Table Append Reproducer';
const paragraphs = [
    'Opening paragraph 1.', 'Opening paragraph 2.', 'Opening paragraph 3.',
    'Opening paragraph 4.', 'Opening paragraph 5.', 'Opening paragraph 6.',
    'Final paragraph stays unchanged.'
];
const table = '| Mountain | River | Forest |\n| --- | --- | --- |\n| Ocean | Valley | Canyon |\n| Meadow | Desert | Island |';
const sourceDocumentXml = `<w:document xmlns:w="${W}"><w:body>${paragraphs.map(text => (
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
)).join('')}<w:sectPr/></w:body></w:document>`;

function makeDocx(documentXml) {
    return zipDocx(new Map([
        ['[Content_Types].xml', `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`],
        ['_rels/.rels', `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
        ['word/_rels/document.xml.rels', `<Relationships xmlns="${REL}"/>`],
        ['word/document.xml', documentXml]
    ]));
}

function state(documentXml) {
    const inspected = inspectDocumentParts({ documentXml });
    assert.equal(inspected.status, 'ok', JSON.stringify(inspected.errors));
    return {
        paragraphs: inspected.paragraphs.filter(item => !item.inTable).map(item => item.exactText),
        cells: inspected.paragraphs.filter(item => item.inTable).map(item => item.exactText),
        tables: (documentXml.match(/<w:tbl\b/g) || []).length
    };
}

const sourceInspection = inspectDocumentParts({ documentXml: sourceDocumentXml });
assert.equal(sourceInspection.status, 'ok');
const target = sourceInspection.paragraphs.at(-1);
const operation = {
    type: 'redline',
    target,
    modified: `${target.exactText}\n${table}`,
    structuredContent: true
};
const tracked = await applyOperationsToDocumentXml(sourceDocumentXml, [operation], AUTHOR, null, {
    atomic: true,
    structuredContent: true,
    pairReplacements: true,
    generateRedlines: true
});
assert.equal(tracked.status, 'ok', JSON.stringify(tracked.error || tracked.results));
const accepted = acceptTrackedChangesInOoxml(tracked.documentXml, { allAuthors: true });
const rejected = rejectTrackedChangesInOoxml(tracked.documentXml, { allAuthors: true });
assert.notEqual(accepted.status, 'error', JSON.stringify(accepted.error));
assert.notEqual(rejected.status, 'error', JSON.stringify(rejected.error));
assert.equal(state(accepted.oxml).tables, 1);
assert.equal(state(rejected.oxml).tables, 0);
assert.deepEqual(state(rejected.oxml).paragraphs, paragraphs, 'the public engine resolver restores the source exactly');

const args = process.argv.slice(2);
const outputOption = args.indexOf('--export-dir');
assert.ok(outputOption >= 0 && args[outputOption + 1] && !args[outputOption + 1].startsWith('--'),
    'usage: node docs/library-issues/2026-10-01-word-reject-table-append-paragraph-export.mjs --export-dir .cache/word-reject-table-append');
const outputDir = resolve(args[outputOption + 1]);
mkdirSync(outputDir, { recursive: true });
for (const [name, xml] of [
    ['source.docx', sourceDocumentXml],
    ['tracked.docx', tracked.documentXml],
    ['accepted.docx', accepted.oxml],
    ['rejected.docx', rejected.oxml]
]) writeFileSync(resolve(outputDir, name), makeDocx(xml));
writeFileSync(resolve(outputDir, 'insertion.xml'), tracked.documentXml, 'utf8');

const emptyTables = [];
const expectedTable = [{ rows: 3, columns: 3, cells: [
    ['Mountain', 'River', 'Forest'],
    ['Ocean', 'Valley', 'Canyon'],
    ['Meadow', 'Desert', 'Island']
] }];
const manifest = {
    cases: [{
        name: 'public-library-table-append-reject-all',
        source: 'source.docx',
        tracked: 'tracked.docx',
        accepted: 'accepted.docx',
        rejected: 'rejected.docx',
        insertionXml: 'insertion.xml',
        expectedBodyStructure: {
            source: { paragraphs, tables: emptyTables },
            accepted: { paragraphs: [...paragraphs, ''], tables: expectedTable },
            rejected: { paragraphs, tables: emptyTables }
        },
        expectedAcceptedText: [...paragraphs, 'Mountain', 'River', 'Forest', 'Ocean', 'Valley', 'Canyon', 'Meadow', 'Desert', 'Island'].join('\n'),
        expectedRejectedText: paragraphs.join('\n'),
        expectedMinimumBodyRevisions: 1,
        expectedEngineReferencePackages: true,
        nativeOperations: [operation]
    }]
};
writeFileSync(resolve(outputDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
console.log(`Exported one public-library Word Reject All fixture to ${outputDir}`);
console.log(`Public engine Reject All: ${state(rejected.oxml).paragraphs.length} paragraphs, ${state(rejected.oxml).tables} tables.`);
