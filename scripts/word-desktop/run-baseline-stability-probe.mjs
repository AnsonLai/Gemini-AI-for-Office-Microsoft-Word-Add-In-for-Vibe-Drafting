import '../../tests/setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { configureLogger } from '@ansonlai/docx-redline-js';
import { zipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { captureWordSourceBaseline } from '../../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '../..');
const undoMode = process.argv.includes('--undo');
const outputDir = join(repoRoot, '.cache', undoMode ? 'word-undo-probe' : 'word-baseline-probe');
const sourcePath = join(outputDir, 'baseline-source.docx');
const powerShellPath = join(here, undoMode ? 'capture-undo-word-xml.ps1' : 'capture-baseline-word-xml.ps1');
const paragraphs = [
    'Paragraph P1 baseline.',
    'Paragraph P2 baseline.',
    'Paragraph P3 baseline.',
    'Target paragraph P4 baseline.',
    'Paragraph P5 baseline.'
];

configureLogger({ log() {}, warn() {}, error() {} }, { level: 'silent' });
mkdirSync(outputDir, { recursive: true });
const documentXml = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"><w:body>${paragraphs.map((text, index) => (
    `<w:p w14:paraId="${(0xA0000000 + index + 1).toString(16).toUpperCase()}"><w:r><w:t>${text}</w:t></w:r></w:p>`
)).join('')}<w:sectPr/></w:body></w:document>`;
const sourceBytes = zipDocx(new Map([
    ['[Content_Types].xml', `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`],
    ['_rels/.rels', `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="${REL}"/>`],
    ['word/document.xml', documentXml]
]));
writeFileSync(sourcePath, sourceBytes);

if (!process.argv.includes('--analyze-existing')) {
    execFileSync('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', powerShellPath,
        '-SourcePath', sourcePath, '-OutputDir', outputDir
    ], { stdio: 'inherit', timeout: 120000 });
}

const report = JSON.parse(readFileSync(join(outputDir, 'word-capture-report.json'), 'utf8').replace(/^\uFEFF/, ''));
const snapshots = report.captures.map(capture => {
    const xml = readFileSync(join(outputDir, capture.file), 'utf8');
    return {
        name: capture.name,
        sha256: createHash('sha256').update(xml).digest('hex'),
        baseline: captureWordSourceBaseline(xml).map(({ index, exactText, fingerprint, paragraphId, inTable }) => ({
            index, exactText, fingerprint, paragraphId, inTable
        }))
    };
});
const baselineByName = Object.fromEntries(snapshots.map(snapshot => [snapshot.name, snapshot.baseline]));
const initial = baselineByName['01-source-read-a'];
const compareNames = [
    '02-source-read-b', '03-tracking-on-no-edit', '04-tracking-off-no-edit',
    '06-after-reject-all', '07-after-reject-tracking-off',
    '08-after-save-restart-read-a', '09-after-save-restart-read-b',
    '11-after-undo', '12-after-undo-tracking-off',
    '14-after-untracked-undo', '15-after-untracked-undo-read-b'
].filter(name => Object.hasOwn(baselineByName, name));
const unexecutedSteps = undoMode ? ['save/reopen (Word restart)'] : [
    'save/reopen after Reject All',
    'tracked edit followed by Undo',
    'untracked edit followed by Undo'
];
const comparisons = compareNames.map(name => ({
    name,
    baselineMatchesInitial: JSON.stringify(baselineByName[name]) === JSON.stringify(initial),
    p4: baselineByName[name]?.[3] ?? null,
    fingerprintChangedIndexes: initial.flatMap((paragraph, index) => (
        paragraph.fingerprint !== baselineByName[name]?.[index]?.fingerprint ? [index + 1] : []
    )),
    paragraphIdChangedIndexes: initial.flatMap((paragraph, index) => (
        paragraph.paragraphId !== baselineByName[name]?.[index]?.paragraphId ? [index + 1] : []
    ))
}));
assert.ok(initial?.[3], 'Word source capture contains target paragraph P4');

const output = {
    environment: {
        wordVersion: report.wordVersion ?? null,
        wordBuild: report.wordBuild ?? null,
        transport: 'Word COM Content.WordOpenXML; this is not proof of Office.js body.getOoxml behavior',
        captureLimit: report.captureLimit,
        parsedParagraphIdsAvailable: initial.some(paragraph => paragraph.paragraphId !== null)
    },
    fixture: { file: basename(sourcePath), paragraphs },
    snapshots,
    comparisons,
    unexecutedSteps,
    undoResults: report.undoResults ?? null,
    editedP4: baselineByName['05-tracked-edit']?.[3] ?? null,
    conclusion: comparisons.every(item => item.baselineMatchesInitial)
        ? (undoMode ? 'After Undo, Word COM paragraph text/fingerprints matched the initial baseline. Restart was not run; Office.js was not tested.' : 'Word COM paragraph text/fingerprints stayed stable across repeated reads, tracking toggles, and Reject All. Paragraph IDs were absent from parsed source. Save/reopen and undo were not run, so this does not explain the reported restart behavior; Office.js was not tested.')
        : 'Word COM paragraph text/fingerprints differed from the initial baseline in at least one comparison; see comparisons.'
};
writeFileSync(join(outputDir, 'baseline-comparison.json'), JSON.stringify(output, null, 2));
console.log(JSON.stringify(output, null, 2));
