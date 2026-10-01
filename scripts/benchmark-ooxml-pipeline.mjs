import '../tests/setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { cpus, hostname, platform, release, arch, totalmem } from 'node:os';
import { dirname, resolve } from 'node:path';
import { execFileSync } from 'node:child_process';
import { performance } from 'node:perf_hooks';
import { fileURLToPath } from 'node:url';
import { acceptTrackedChangesInOoxml, configureLogger, openDocx } from '@ansonlai/docx-redline-js';
import { zipDocx, unzipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';
import { createDocumentSession } from '../browser-demo/document-session.js';
import { planRedlineBatchOperations } from '../src/taskpane/modules/docx-redline-js-integration/redline-plan.js';
import { captureSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';

configureLogger({ info() {}, warn() {}, error() {} });

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');
const warmups = positiveInteger('DOCX_BENCH_WARMUPS', 2, true);
const iterations = positiveInteger('DOCX_BENCH_ITERATIONS', 10);
const operationCount = positiveInteger('DOCX_BENCH_OPERATIONS', 10);
const encoder = new TextEncoder();
const decoder = new TextDecoder();
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const AUTHOR = 'OOXML Performance Benchmark';

function positiveInteger(name, fallback, allowZero = false) {
    const value = process.env[name] ?? String(fallback);
    const pattern = allowZero ? /^(?:0|[1-9]\d*)$/ : /^[1-9]\d*$/;
    if (!pattern.test(value)) throw new Error(`${name} must be ${allowZero ? 'a non-negative' : 'a positive'} integer`);
    return Number(value);
}

function rounded(value) { return Number(value.toFixed(3)); }

function statistics(samples) {
    const sorted = [...samples].sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    return {
        sampleCount: samples.length,
        medianMs: rounded(sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2),
        p95Ms: rounded(sorted[Math.ceil(sorted.length * 0.95) - 1]),
        minMs: rounded(sorted[0]),
        maxMs: rounded(sorted.at(-1)),
        samplesMs: samples.map(rounded)
    };
}

async function measure({ prepare = async () => undefined, action, verify }) {
    for (let i = 0; i < warmups; i += 1) {
        const context = await prepare();
        verify(await action(context));
    }
    const samples = [];
    for (let i = 0; i < iterations; i += 1) {
        const context = await prepare();
        const start = performance.now();
        const value = await action(context);
        samples.push(performance.now() - start);
        verify(value);
    }
    return statistics(samples);
}

function textFor(number) {
    return `Clause ${number}: The receiving party shall retain confidential records for thirty days and return all copies on request.`;
}

function paragraphXml(number, text, extra = '') {
    return `<w:p w14:paraId="${number.toString(16).padStart(8, '0').toUpperCase()}" ${extra}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function packageFromDocumentXml(documentXml, extraParts = new Map()) {
    const entries = new Map([
        ['[Content_Types].xml', encoder.encode(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
        ['_rels/.rels', encoder.encode(`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
        ['word/document.xml', encoder.encode(documentXml)]
    ]);
    for (const [name, contents] of extraParts) entries.set(name, typeof contents === 'string' ? encoder.encode(contents) : contents);
    return zipDocx(entries);
}

function plainFixture(paragraphCount) {
    const body = Array.from({ length: paragraphCount }, (_, i) => paragraphXml(i + 1, textFor(i + 1))).join('');
    const documentXml = `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
    return { bytes: packageFromDocumentXml(documentXml), provenance: 'synthetic independent legal-style paragraphs' };
}

function fullParagraphFixture() {
    const original = Array.from({ length: 24 }, (_, i) => `This extended provision sentence ${i + 1} describes the preservation and transfer of review materials under the agreement.`).join(' ');
    const documentXml = `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphXml(1, original)}<w:sectPr/></w:body></w:document>`;
    return { bytes: packageFromDocumentXml(documentXml), provenance: 'synthetic single long paragraph', original };
}

function tableFixture() {
    const documentXml = `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphXml(1, 'Table workload introduction.')}<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraphXml(2, 'Existing table cell term.')}</w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraphXml(3, 'Untouched table cell.')}</w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraphXml(4, 'Second row cell.')}</w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraphXml(5, 'Second row control.')}</w:tc></w:tr></w:tbl>${paragraphXml(6, 'Table workload conclusion.')}<w:sectPr/></w:body></w:document>`;
    return { bytes: packageFromDocumentXml(documentXml), provenance: 'synthetic existing 2x2 Word table; localized cell-text edit only' };
}

function extractPackage(bytes) {
    const entries = unzipDocx(bytes);
    const xml = entries.get('word/document.xml');
    assert.ok(xml, 'fixture has word/document.xml');
    return {
        entries,
        documentXml: typeof xml === 'string' ? xml : decoder.decode(xml),
        context: Object.fromEntries([
            ['commentsXml', 'word/comments.xml'],
            ['commentsExtendedXml', 'word/commentsExtended.xml'],
            ['commentsIdsXml', 'word/commentsIds.xml'],
            ['commentsExtensibleXml', 'word/commentsExtensible.xml'],
            ['numberingXml', 'word/numbering.xml'],
            ['stylesXml', 'word/styles.xml']
        ].flatMap(([key, part]) => {
            const value = entries.get(part);
            return value == null ? [] : [[key, typeof value === 'string' ? value : decoder.decode(value)]];
        }))
    };
}

function findParagraph(inspection, predicate, description) {
    const paragraph = inspection.paragraphs.find(predicate);
    assert.ok(paragraph, `fixture contains ${description}`);
    assert.ok(paragraph.exactText, `${description} is non-empty`);
    return paragraph;
}

function redlineRequest(operation, paragraph, extra = {}) {
    return {
        operation,
        paragraphIndex: paragraph.index,
        ...extra
    };
}

function makeRedlineMap(aiChanges) {
    return paragraphs => planRedlineBatchOperations(aiChanges, paragraphs);
}

function makeCommentMap(paragraph, textToComment) {
    return paragraphs => {
        const target = paragraphs[paragraph.index - 1];
        assert.equal(target?.index, paragraph.index, 'comment target remains at its inspected index');
        return [{
            type: 'comment',
            target: {
                index: target.index,
                exactText: target.exactText,
                ...(target.paragraphId ? { paragraphId: target.paragraphId } : {}),
                ...(target.fingerprint ? { fingerprint: target.fingerprint } : {}),
                inTable: target.inTable
            },
            textToComment,
            commentContent: 'Performance benchmark comment.'
        }];
    };
}

const paragraphTotal = positiveInteger('DOCX_BENCH_PARAGRAPHS', 100);
const fixtures = [
    { name: 'noop', ...plainFixture(Math.max(10, Math.min(paragraphTotal, 1000))) },
    { name: 'localized_multi_edit', ...plainFixture(Math.max(operationCount, Math.min(paragraphTotal, 1000))) },
    { name: 'full_paragraph_rewrite', ...fullParagraphFixture() },
    {
        name: 'nested_list_localized_edit',
        bytes: new Uint8Array(await readFile(resolve(root, 'tests/fixtures/agentic-lists/nested-lists-source.docx'))),
        provenance: 'Word-authored Microsoft Word Desktop nested-list fixture; edit existing list paragraph text only'
    },
    { name: 'table_cell_localized_edit', ...tableFixture() },
    {
        name: 'commented_document_comment',
        bytes: new Uint8Array(await readFile(resolve(root, 'tests/fixtures/wp6/threaded-comments.docx'))),
        provenance: 'Word-authored Microsoft Word Desktop threaded-comments fixture; add one anchored comment'
    }
];

const scenarios = [];
for (const fixture of fixtures) {
    const pkg = extractPackage(fixture.bytes);
    const sourceDoc = openDocx(fixture.bytes);
    const inspection = sourceDoc.inspect();
    assert.equal(inspection.status, 'ok', `${fixture.name}: source document inspects`);
    const paragraphs = inspection.paragraphs;
    let map;
    let description;
    let expectedChanges = true;
    let expectedText;
    if (fixture.name === 'noop') {
        map = () => [];
        description = 'Empty operation batch; no package mutation expected.';
        expectedChanges = false;
    } else if (fixture.name === 'localized_multi_edit') {
        const count = Math.min(operationCount, paragraphs.length);
        const changes = Array.from({ length: count }, (_, i) => {
            const paragraph = paragraphs[Math.floor(i * paragraphs.length / count)];
            return redlineRequest('modify_text', paragraph, {
                originalText: 'thirty',
                replacementText: 'sixty'
            });
        });
        map = makeRedlineMap(changes);
        description = `${count} distinct localized replacements across ${paragraphs.length} paragraphs.`;
        expectedText = 'sixty';
    } else if (fixture.name === 'full_paragraph_rewrite') {
        const target = findParagraph(inspection, p => p.index === 1, 'full-paragraph target');
        const replacement = Array.from({ length: 24 }, (_, i) => `The reviewing party will store returned originals in secured archive ${i + 1} until final written authorization permits their disposal.`).join(' ');
        map = makeRedlineMap([redlineRequest('edit_paragraph', target, { content: replacement })]);
        description = `Replace one ${target.exactText.length}-character paragraph in full.`;
        expectedText = 'secured archive';
    } else if (fixture.name === 'nested_list_localized_edit') {
        const target = findParagraph(inspection, p => p.exactText.includes('Bullet Insertion Anchor'), 'nested-list anchor paragraph');
        assert.ok(target.list, 'selected paragraph belongs to an existing list');
        map = makeRedlineMap([redlineRequest('modify_text', target, {
            originalText: 'Insertion Anchor', replacementText: 'Insertion Anchor Revised'
        })]);
        description = 'Localized text change within an existing nested list paragraph; no list structure is created or reordered.';
        expectedText = 'Insertion Anchor Revised';
    } else if (fixture.name === 'table_cell_localized_edit') {
        const target = findParagraph(inspection, p => p.exactText.includes('Existing table cell term.'), 'existing table-cell paragraph');
        assert.equal(target.inTable, true, 'selected paragraph is in an existing table cell');
        map = makeRedlineMap([redlineRequest('modify_text', target, {
            originalText: 'table cell term', replacementText: 'table cell phrase'
        })]);
        description = 'Localized text change within an existing table cell; no table structure is created or reordered.';
        expectedText = 'table cell phrase';
    } else {
        const target = findParagraph(inspection, p => p.exactText.trim().length >= 8, 'commented-document anchor paragraph');
        const textToComment = target.exactText.trim().slice(0, Math.min(12, target.exactText.trim().length));
        map = makeCommentMap(target, textToComment);
        description = 'Add one comment to an existing Word-authored threaded-comments package.';
        expectedText = 'Performance benchmark comment.';
    }
    const operations = map(paragraphs);
    assert.ok(Array.isArray(operations), `${fixture.name}: operation mapper returns an array`);
    const commentWorkload = fixture.name === 'commented_document_comment';
    const expectedOperationCount = operations.length;
    const initialCommentCount = inspection.comments.length;
    const sourceMetadata = {
        fixture: fixture.provenance,
        sourceSha256: createHash('sha256').update(fixture.bytes).digest('hex'),
        sourceDocxBytes: fixture.bytes.length,
        sourceDocumentXmlBytes: encoder.encode(pkg.documentXml).length,
        paragraphCount: paragraphs.length,
        tableParagraphCount: paragraphs.filter(p => p.inTable).length,
        listParagraphCount: paragraphs.filter(p => p.list).length,
        existingCommentCount: initialCommentCount,
        workloadDescription: description,
        operationCount: expectedOperationCount,
        expectedChanges,
        expectedText
    };
    scenarios.push({ fixture, pkg, inspection, paragraphs, map, operations, commentWorkload, sourceMetadata });
}

function verifyMapped(scenario, operations) {
    assert.equal(operations.length, scenario.sourceMetadata.operationCount);
    assert.ok(operations.every(operation => typeof operation.type === 'string'));
}

function verifyCore(scenario, result) {
    assert.equal(result.status, 'ok', `${scenario.fixture.name}: core returned ${result.status}: ${JSON.stringify({ error: result.error, results: result.results })}`);
    assert.equal(result.hasChanges, scenario.sourceMetadata.expectedChanges, `${scenario.fixture.name}: expected hasChanges value`);
    if (scenario.sourceMetadata.expectedChanges) {
        if (scenario.commentWorkload) {
            assert.ok(result.commentsXml?.includes(scenario.sourceMetadata.expectedText), 'comment text emitted by core');
        } else {
            const accepted = acceptTrackedChangesInOoxml(result.documentXml, { author: AUTHOR });
            assert.ok(typeof accepted.oxml === 'string', `${scenario.fixture.name}: core output can be accepted`);
            const acceptedInspection = openDocx(packageFromDocumentXml(accepted.oxml)).inspect();
            assert.ok(acceptedInspection.paragraphs.some(p => p.exactText.includes(scenario.sourceMetadata.expectedText)),
                `${scenario.fixture.name}: replacement text emitted by core (${acceptedInspection.paragraphs.map(p => p.exactText).join(' | ')})`);
        }
    }
}

function verifyPackage(scenario, result, outputBytes) {
    assert.equal(result.status, 'ok', `${scenario.fixture.name}: package API returned ${result.status}: ${JSON.stringify(result.error)}`);
    assert.equal(result.written, scenario.sourceMetadata.expectedChanges, `${scenario.fixture.name}: expected write outcome`);
    assert.ok(outputBytes instanceof Uint8Array && outputBytes.length > 0, 'package serialization returns bytes');
    const reopened = openDocx(outputBytes);
    const reopenedInspection = reopened.inspect();
    assert.equal(reopenedInspection.status, 'ok', `${scenario.fixture.name}: output reopens`);
    if (scenario.sourceMetadata.expectedChanges) {
        if (scenario.commentWorkload) {
            assert.equal(reopenedInspection.comments.length, scenario.sourceMetadata.existingCommentCount + 1, 'new comment survives serialization and reopen');
        } else {
            assert.ok(reopenedInspection.paragraphs.some(p => p.exactText.includes(scenario.sourceMetadata.expectedText)), `${scenario.fixture.name}: replacement survives serialization and reopen`);
        }
    }
}

const workloadResults = [];
for (const scenario of scenarios) {
    const { fixture, pkg, inspection, operations, map } = scenario;
    const context = pkg.context;
    const coreResult = await applyOperationsToDocumentXml(pkg.documentXml, operations, AUTHOR, context, {
        atomic: true, structuredContent: true, pairReplacements: true, generateRedlines: true
    });
    verifyCore(scenario, coreResult);
    const opened = openDocx(fixture.bytes);
    const packageResult = await opened.applyOperations(operations, { author: AUTHOR, atomic: true, generateRedlines: true });
    verifyPackage(scenario, packageResult, opened.toUint8Array());

    const measurements = {
        consumerOperationMapping: await measure({
            prepare: async () => inspection.paragraphs,
            action: paragraphs => map(paragraphs),
            verify: operationsForCheck => verifyMapped(scenario, operationsForCheck)
        }),
        portableSourceBaselineCapture: await measure({
            action: () => captureSourceBaseline(pkg.documentXml),
            verify: baseline => assert.equal(baseline.length, inspection.paragraphs.length, 'portable baseline paragraph count')
        }),
        packageOpen: await measure({
            action: () => openDocx(fixture.bytes),
            verify: doc => assert.equal(doc.inspect().status, 'ok', 'package open output inspects')
        }),
        packageInspection: await measure({
            prepare: async () => openDocx(fixture.bytes),
            action: doc => doc.inspect(),
            verify: result => assert.equal(result.status, 'ok', 'package inspection status')
        }),
        browserSessionOpen: await measure({
            action: () => createDocumentSession(fixture.bytes),
            verify: session => assert.ok(session && typeof session.getPromptParagraphs === 'function')
        }),
        browserPromptProjection: await measure({
            prepare: async () => createDocumentSession(fixture.bytes),
            action: session => session.getPromptParagraphs(),
            verify: paragraphs => assert.ok(Array.isArray(paragraphs), 'browser prompt projection returns paragraphs')
        }),
        coreApplication: await measure({
            action: () => applyOperationsToDocumentXml(pkg.documentXml, operations, AUTHOR, context, {
                atomic: true, structuredContent: true, pairReplacements: true, generateRedlines: true
            }),
            verify: result => verifyCore(scenario, result)
        }),
        packageApply: await measure({
            prepare: async () => openDocx(fixture.bytes),
            action: doc => doc.applyOperations(operations, { author: AUTHOR, atomic: true, generateRedlines: true }),
            verify: result => assert.equal(result.status, 'ok', 'package apply status')
        }),
        packageSerialize: await measure({
            prepare: async () => {
                const doc = openDocx(fixture.bytes);
                const result = await doc.applyOperations(operations, { author: AUTHOR, atomic: true, generateRedlines: true });
                assert.equal(result.status, 'ok', 'prepare package for save-only timing');
                return doc;
            },
            action: doc => doc.toUint8Array(),
            verify: bytes => assert.ok(bytes instanceof Uint8Array && bytes.length > 0, 'package save output bytes')
        }),
        packageOpenApplySave: await measure({
            action: async () => {
                const doc = openDocx(fixture.bytes);
                const result = await doc.applyOperations(operations, { author: AUTHOR, atomic: true, generateRedlines: true });
                const bytes = doc.toUint8Array();
                return { result, bytes };
            },
            verify: ({ result, bytes }) => verifyPackage(scenario, result, bytes)
        })
    };

    workloadResults.push({
        name: fixture.name,
        source: scenario.sourceMetadata,
        timing: measurements
    });
}

const scriptBytes = await readFile(fileURLToPath(import.meta.url));
const lock = JSON.parse(await readFile(resolve(root, 'node_modules/@ansonlai/docx-redline-js/package.json'), 'utf8'));
let gitCommit = null;
try { gitCommit = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim(); } catch { /* Optional source metadata. */ }
let dirtyTree = null;
try { dirtyTree = execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim().length > 0; } catch { /* Optional source metadata. */ }

const report = {
    schemaVersion: 1,
    measuredAt: new Date().toISOString(),
    benchmark: {
        name: 'OOXML consumer pipeline observational benchmark',
        harness: 'scripts/benchmark-ooxml-pipeline.mjs',
        harnessSha256: createHash('sha256').update(scriptBytes).digest('hex'),
        gitCommit,
        workingTreeDirtyAtRun: dirtyTree,
        applicationPackageVersion: JSON.parse(await readFile(resolve(root, 'package.json'), 'utf8')).version,
        docxRedlinePackageVersion: lock.version,
        docxRedlinePackagePin: JSON.parse(await readFile(resolve(root, 'package.json'), 'utf8')).dependencies['@ansonlai/docx-redline-js'],
        nodeVersion: process.version,
        v8Version: process.versions.v8,
        npmUserAgent: process.env.npm_config_user_agent || null
    },
    machine: {
        hostname: hostname(),
        platform: platform(),
        osRelease: release(),
        architecture: arch(),
        cpuModel: cpus()[0]?.model || null,
        logicalCpuCount: cpus().length,
        totalMemoryBytes: totalmem()
    },
    sampling: {
        warmupsPerMetric: warmups,
        measuredIterationsPerMetric: iterations,
        nearestRankP95: true,
        rawSamplesAreMilliseconds: true
    },
    methodology: {
        mapping: 'Consumer redline operation mapping uses the production planRedlineBatchOperations helper; the comment workload maps the anchored comment intent directly to the public document operation shape.',
        portableBaselineCapture: 'Times captureSourceBaseline(documentXml) after package-open/source fixture preparation.',
        coreApplication: 'Times applyOperationsToDocumentXml(documentXml, operations, ...) using pre-mapped operations; excludes package open/save and host transport.',
        packageOpen: 'Times openDocx(sourceBytes) only.',
        packageInspection: 'Times doc.inspect() on an already-open document.',
        browserSessionOpen: 'Times createDocumentSession(sourceBytes).',
        browserPromptProjection: 'Times session.getPromptParagraphs() on an already-open browser document session; includes its inspect, parse, and paragraph projection work.',
        packageApply: 'Times DocxDocument.applyOperations on an already-open package; excludes source open and final serialization.',
        packageSerialize: 'Times toUint8Array() on an already-edited package; source open and mutation are prepared outside the timer.',
        packageOpenApplySave: 'Times openDocx + applyOperations + toUint8Array end to end.',
        exclusions: ['Microsoft Word/Office.js read/write and context.sync latency', 'LLM/provider/network/tool latency', 'disk I/O', 'browser rendering and user interaction'],
        structuralFidelity: 'Nested-list and table cases are localized text edits inside existing structures. The benchmark does not exercise or certify list/table structural mutation.',
        gates: 'Observational only; no universal latency threshold or pass/fail budget.'
    },
    results: workloadResults
};

const outputArg = process.argv.find(arg => arg.startsWith('--output='));
if (outputArg) {
    const outputPath = resolve(root, outputArg.slice('--output='.length));
    await mkdir(dirname(outputPath), { recursive: true });
    await writeFile(outputPath, `${JSON.stringify(report, null, 2)}\n`);
}
console.log(JSON.stringify(report, null, 2));
