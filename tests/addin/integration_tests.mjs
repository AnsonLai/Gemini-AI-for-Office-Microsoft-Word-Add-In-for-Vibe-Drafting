import '../setup-xml-provider.mjs';
import { parseXml, serializeXml } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';

/**
 * Integration Tests
 *
 * Combines NDA Redlining Tests and Redline Toggle Tests.
 * Run with: node tests/addin/integration_tests.mjs
 *
 * NDA cases edit the complete sample word/document.xml the way production does
 * (consumer-core.js / document-tools.js): strong paragraph targets from
 * inspectDocumentParts (paragraphId + fingerprint) and applyOperationsToDocumentXml with
 * atomic + strictTargets. applyRedlineToOxml is limited to paragraph, range or table scope
 * (library README, Output Shape Matrix) and parses modifiedText as Markdown, so replaying a
 * whole document through it corrupts unrelated paragraphs.
 *
 * Every NDA case asserts paragraph by paragraph on both the Accept All and Reject All views
 * (not OOXML substrings), so a corrupted result cannot pass:
 *   - Accept All text equals the source with exactly the requested edit applied.
 *   - Reject All text equals the source exactly.
 *   - Untouched paragraphs are byte-identical to the source in both views.
 *   - Character formatting (bold/underline) lands on exactly the intended text.
 */

import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import {
    applyRedlineToOxml,
    inspectDocumentParts,
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml,
    mergeNumberingXmlBySchemaOrder
} from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

// --- Config ---
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, '../sample_doc/word/document.xml');
const NUMBERING_PATH = path.join(__dirname, '../sample_doc/word/numbering.xml');

// --- Helper Functions ---
async function readSampleDoc() {
    return await fs.readFile(DOC_PATH, 'utf-8');
}

async function readSampleNumbering() {
    return await fs.readFile(NUMBERING_PATH, 'utf-8');
}

function assertSuccessfulResult(result, context) {
    if (result?.status === 'error' || result?.error) {
        const code = result?.error?.code || 'OPERATION_ERROR';
        throw new Error(`[${code}] ${context}: ${result?.error?.message || 'operation failed'}`);
    }
    return result;
}

// --- Whole-document harness ---
function childElements(node, name) {
    const out = [];
    for (let i = 0; i < node.childNodes.length; i++) {
        const child = node.childNodes[i];
        if (child.nodeType === 1 && child.nodeName === name) out.push(child);
    }
    return out;
}

function isToggleOn(element) {
    if (!element) return null;
    const val = (element.getAttribute('w:val') || '').toLowerCase();
    return !(val === '0' || val === 'false' || val === 'off');
}

// Effective bold/underline of a run. An explicit <w:b/> wins over the Strong character style
// the sample uses for its title and "British Columbia".
function runFlags(run) {
    const rPr = childElements(run, 'w:rPr')[0];
    if (!rPr) return { b: false, u: false };
    const explicitBold = isToggleOn(childElements(rPr, 'w:b')[0]);
    const style = childElements(rPr, 'w:rStyle')[0]?.getAttribute('w:val');
    const underline = childElements(rPr, 'w:u')[0];
    return {
        b: explicitBold === null ? style === 'Strong' : explicitBold,
        u: !!underline && (underline.getAttribute('w:val') || 'single') !== 'none'
    };
}

// Per document-order paragraph: one entry per visible character with its bold/underline flags.
function characterFormats(documentXml) {
    const doc = parseXml(documentXml);
    const paragraphs = doc.getElementsByTagName('w:p');
    const result = [];
    for (let i = 0; i < paragraphs.length; i++) {
        const chars = [];
        const runs = paragraphs[i].getElementsByTagName('w:r');
        for (let r = 0; r < runs.length; r++) {
            const flags = runFlags(runs[r]);
            const texts = runs[r].getElementsByTagName('w:t');
            for (let t = 0; t < texts.length; t++) {
                for (const ch of texts[t].textContent) chars.push({ ch, b: flags.b, u: flags.u });
            }
        }
        result.push(chars);
    }
    return result;
}

const formatKey = (c) => `${c.ch}|${c.b ? 1 : 0}|${c.u ? 1 : 0}`;
const plainText = (chars) => chars.map(c => c.ch).join('');

function paragraphXmlById(documentXml) {
    const doc = parseXml(documentXml);
    const map = new Map();
    const paragraphs = doc.getElementsByTagName('w:p');
    for (let i = 0; i < paragraphs.length; i++) {
        const id = paragraphs[i].getAttribute('w14:paraId');
        if (id) map.set(id, serializeXml(paragraphs[i]));
    }
    return map;
}

function countTables(documentXml) {
    return parseXml(documentXml).getElementsByTagName('w:tbl').length;
}

function strongTarget(paragraph) {
    return {
        exactText: paragraph.exactText, index: paragraph.index, paragraphId: paragraph.paragraphId,
        fingerprint: paragraph.fingerprint, revisionView: paragraph.revisionView, inTable: paragraph.inTable
    };
}

function findParagraphs(source, predicate, name, expectedCount) {
    const matches = source.filter(predicate);
    if (matches.length !== expectedCount) {
        throw new Error(`${name}: expected ${expectedCount} matching source paragraph(s), found ${matches.length}`);
    }
    return matches;
}

function describeDifferences(label, actual, expected, limit = 6) {
    const diffs = [];
    const length = Math.max(actual.length, expected.length);
    for (let i = 0; i < length; i++) {
        if (actual[i] !== expected[i]) {
            diffs.push(`#${i + 1} expected ${JSON.stringify(expected[i])?.slice(0, 90)} but got ${JSON.stringify(actual[i])?.slice(0, 90)}`);
        }
    }
    if (!diffs.length) return null;
    return `${label} differs from expected in ${diffs.length} paragraph(s) (${actual.length} actual vs ${expected.length} expected): ${diffs.slice(0, limit).join('; ')}`;
}

/**
 * Apply one batch to the complete sample document through the supported API and return both
 * review views. `build` receives the inspected source paragraphs and returns operations.
 */
async function applyToSampleDocument(name, build) {
    const documentXml = await readSampleDoc();
    const numberingXml = await readSampleNumbering();
    const inspection = inspectDocumentParts({ documentXml, numberingXml });
    if (inspection.status !== 'ok') throw new Error(`${name}: sample document could not be inspected`);
    const source = inspection.paragraphs;
    const operations = build(source);

    const result = assertSuccessfulResult(await applyOperationsToDocumentXml(
        documentXml, operations, 'TestUser', { numberingXml },
        { atomic: true, strictTargets: true, generateRedlines: true, pairReplacements: true }
    ), name);
    if (result.status !== 'ok' || result.rolledBack) throw new Error(`${name}: batch did not complete (status ${result.status}, rolledBack ${result.rolledBack})`);
    if (!result.hasChanges) throw new Error(`${name}: no changes detected by engine`);

    let outputNumbering = numberingXml;
    for (const part of result.numberingXmlParts || []) outputNumbering = mergeNumberingXmlBySchemaOrder(outputNumbering, part);
    const acceptedXml = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    const rejectedXml = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
    const view = (xml) => inspectDocumentParts({ documentXml: xml, numberingXml: outputNumbering }).paragraphs;
    return {
        name, source, result, sourceXml: documentXml, acceptedXml, rejectedXml,
        accepted: view(acceptedXml), rejected: view(rejectedXml)
    };
}

/**
 * Accept All paragraph plan: every source paragraph unchanged except those in `replacements`
 * (0-based source index -> array of { text, from }, where `from` is the 0-based source index the
 * paragraph derives from, or null for a new paragraph).
 */
function buildExpected(source, replacements = new Map()) {
    return source.flatMap((p, from) => replacements.get(from) ?? [{ text: p.exactText, from }]);
}

/**
 * Check both views against an explicit plan. `touched` lists the 0-based source paragraphs the
 * batch targets; every other paragraph must be byte-identical in both views. `formatEdits`
 * ({ from, find, prop: 'b'|'u', value }) states the only character formatting expected to change.
 */
function verifyViews(run, { expected, touched = [], formatEdits = [] }) {
    const { name, source, accepted, rejected } = run;
    const problems = [];
    const touchedSet = new Set(touched);

    const acceptDiff = describeDifferences('Accept All', accepted.map(p => p.exactText), expected.map(e => e.text));
    if (acceptDiff) problems.push(acceptDiff);
    const rejectDiff = describeDifferences('Reject All', rejected.map(p => p.exactText), source.map(p => p.exactText));
    if (rejectDiff) problems.push(rejectDiff);

    const sourceFormats = characterFormats(run.sourceXml);
    const acceptedFormats = characterFormats(run.acceptedXml);
    const rejectedFormats = characterFormats(run.rejectedXml);
    if (sourceFormats.length !== source.length || acceptedFormats.length !== accepted.length || rejectedFormats.length !== rejected.length) {
        problems.push('paragraph count of the formatting scan does not match the inspected paragraph count');
    }

    const expectedFormats = sourceFormats.map(chars => chars.map(c => ({ ...c })));
    for (const edit of formatEdits) {
        const chars = expectedFormats[edit.from];
        const start = plainText(chars).indexOf(edit.find);
        if (start < 0) { problems.push(`format target "${edit.find}" not found in source paragraph ${edit.from + 1}`); continue; }
        for (let i = start; i < start + edit.find.length; i++) chars[i][edit.prop] = edit.value;
    }

    const sourceXml = paragraphXmlById(run.sourceXml);
    const acceptedXml = paragraphXmlById(run.acceptedXml);
    const rejectedXml = paragraphXmlById(run.rejectedXml);
    const formatProblems = [];
    const byteProblems = [];
    expected.forEach((entry, i) => {
        if (entry.from === null) return;
        const origin = source[entry.from];
        if (entry.text === origin.exactText && acceptedFormats[i]) {
            const wantKeys = expectedFormats[entry.from].map(formatKey).join(',');
            if (acceptedFormats[i].map(formatKey).join(',') !== wantKeys) {
                const changed = acceptedFormats[i].filter((c, k) => formatKey(c) !== formatKey(sourceFormats[entry.from][k] || {}));
                formatProblems.push(`Accept All paragraph #${i + 1} "${origin.exactText.slice(0, 40)}": formatting differs from expected (formatting changed on ${JSON.stringify(plainText(changed))})`);
            }
        }
        if (!touchedSet.has(entry.from) && origin.paragraphId && acceptedXml.get(origin.paragraphId) !== sourceXml.get(origin.paragraphId)) {
            byteProblems.push(`Accept All paragraph #${i + 1} (${origin.paragraphId})`);
        }
    });
    source.forEach((origin, i) => {
        if (rejectedFormats[i] && rejectedFormats[i].map(formatKey).join(',') !== sourceFormats[i].map(formatKey).join(',')) {
            formatProblems.push(`Reject All paragraph #${i + 1} "${origin.exactText.slice(0, 40)}": formatting not restored`);
        }
        if (!touchedSet.has(i) && origin.paragraphId && rejectedXml.get(origin.paragraphId) !== sourceXml.get(origin.paragraphId)) {
            byteProblems.push(`Reject All paragraph #${i + 1} (${origin.paragraphId})`);
        }
    });
    if (formatProblems.length) problems.push(formatProblems.slice(0, 6).join('; '));
    if (byteProblems.length) problems.push(`untouched paragraphs not byte-identical to the source: ${byteProblems.slice(0, 6).join(', ')}`);
    if (formatEdits.length && !/<w:rPrChange\b/.test(run.result.documentXml)) {
        problems.push('formatting was not recorded as a tracked w:rPrChange');
    }
    if (problems.length) throw new Error(`${name}: ${problems.join(' | ')}`);
}

function assertList(name, label, actual, wanted) {
    const got = actual.list;
    for (const [key, value] of Object.entries(wanted)) {
        if (got?.[key] !== value) {
            throw new Error(`${name}: ${label} list ${key} expected ${JSON.stringify(value)} but got ${JSON.stringify(got?.[key])} (list ${JSON.stringify(got)})`);
        }
    }
}

async function runCase(name, body) {
    console.log(`\n=== Test: ${name} ===`);
    try {
        await body();
        console.log('✅ PASS');
        return null;
    } catch (e) {
        console.error(`💥 FAIL: ${e.message}`);
        return { name, error: e };
    }
}

// --- NDA Redline Tests ---
async function runNdaTests() {
    console.log('Starting NDA Redlining Tests...');
    const failures = [];
    const check = async (name, body) => { const failure = await runCase(name, body); if (failure) failures.push(failure); };

    await check('Add Instructions Paragraph', async () => {
        const name = 'Add Instructions Paragraph';
        const text = 'Instructions: Please review the following NDA carefully.';
        const run = await applyToSampleDocument(name, (source) => [{
            type: 'redline', target: strongTarget(source[0]),
            modified: `${text}\n${source[0].exactText}`, structuredContent: true
        }]);
        verifyViews(run, {
            expected: buildExpected(run.source, new Map([[0, [{ text, from: null }, { text: run.source[0].exactText, from: 0 }]]])),
            touched: [0]
        });
        if (run.accepted[0].list || run.accepted[0].inTable) throw new Error(`${name}: inserted paragraph must be a plain body paragraph`);
        if (!/<w:ins[\s>]/.test(run.result.documentXml)) throw new Error(`${name}: expected tracked insertion markup`);
    });

    await check('Underline Title', async () => {
        const name = 'Underline Title';
        const find = 'NON-DISCLOSURE AGREEMENT';
        const run = await applyToSampleDocument(name, (source) => {
            const [title] = findParagraphs(source, p => p.exactText.includes(find), name, 1);
            return [{ type: 'format', target: strongTarget(title), textToFormat: find, properties: { underline: true } }];
        });
        const from = run.source.findIndex(p => p.exactText.includes(find));
        verifyViews(run, {
            expected: buildExpected(run.source), touched: [from],
            formatEdits: [{ from, find, prop: 'u', value: true }]
        });
    });

    await check('Recitals to Ordered List', async () => {
        const name = 'Recitals to Ordered List';
        const run = await applyToSampleDocument(name, (source) => {
            const [recitals] = findParagraphs(source, p => p.exactText.startsWith('A. The Disclosing'), name, 1);
            const items = recitals.exactText.split('\n');
            if (items.length !== 3) throw new Error(`${name}: expected three typed recitals, found ${items.length}`);
            return [{
                type: 'list-change', target: strongTarget(recitals), structuredContent: true,
                modified: items.map((item, i) => `${i + 1}. ${item.replace(/^[A-C]\.\s+/, '')}`).join('\n')
            }];
        });
        const from = run.source.findIndex(p => p.exactText.startsWith('A. The Disclosing'));
        const items = run.source[from].exactText.split('\n').map(item => item.replace(/^[A-C]\.\s+/, ''));
        verifyViews(run, {
            expected: buildExpected(run.source, new Map([[from, items.map(text => ({ text, from: null }))]])),
            touched: [from]
        });
        const sourceNumIds = new Set(run.source.filter(p => p.list).map(p => p.list.numId));
        const numId = run.accepted[from].list?.numId;
        if (!numId || numId === '0' || sourceNumIds.has(numId)) throw new Error(`${name}: expected a new positive numId, got ${JSON.stringify(numId)}`);
        items.forEach((_, k) => assertList(name, `recital ${k + 1}`, run.accepted[from + k], { numId, level: 0, label: `${k + 1}.`, format: 'decimal' }));
        if (run.rejected[from].list) throw new Error(`${name}: Reject All must restore the original typed A./B./C. paragraph, not a list`);
    });

    await check('Parties to Table', async () => {
        const name = 'Parties to Table';
        // Replaces the party blocks ("Disclosing Party: ...", "And", "Receiving Party: ...") with a table.
        const cells = ['Disclosing Party', 'Receiving Party', '[Name of Disclosing Party]', '[Name of Receiving Party]',
            '[Address of Disclosing Party]', '[Address of Receiving Party]'];
        const run = await applyToSampleDocument(name, (source) => {
            const start = source.find(p => p.exactText.startsWith('Disclosing Party:'));
            const end = source.find(p => p.exactText.startsWith('Receiving Party:'));
            if (!start || !end) throw new Error(`${name}: party blocks not found in the sample`);
            return [{
                type: 'redline', target: strongTarget(start), targetEnd: strongTarget(end), structuredContent: true,
                modified: '| Disclosing Party | Receiving Party |\n| --- | --- |\n| [Name of Disclosing Party] | [Name of Receiving Party] |\n| [Address of Disclosing Party] | [Address of Receiving Party] |'
            }];
        });
        const start = run.source.findIndex(p => p.exactText.startsWith('Disclosing Party:'));
        const end = run.source.findIndex(p => p.exactText.startsWith('Receiving Party:'));
        const replacements = new Map();
        for (let i = start; i <= end; i++) replacements.set(i, i === start ? cells.map(text => ({ text, from: null })) : []);
        verifyViews(run, {
            expected: buildExpected(run.source, replacements),
            touched: Array.from({ length: end - start + 1 }, (_, k) => start + k)
        });
        cells.forEach((_, k) => {
            const cell = run.accepted[start + k];
            const want = { tableIndex: 1, rowIndex: 1 + Math.floor(k / 2), cellIndex: 1 + (k % 2) };
            if (!cell.inTable || JSON.stringify(cell.table) !== JSON.stringify(want)) {
                throw new Error(`${name}: cell ${JSON.stringify(cell.exactText)} expected at ${JSON.stringify(want)} but got ${JSON.stringify(cell.table)}`);
            }
        });
        if (countTables(run.acceptedXml) !== 2 || countTables(run.rejectedXml) !== 1 || countTables(run.sourceXml) !== 1) {
            throw new Error(`${name}: table counts wrong (source ${countTables(run.sourceXml)}, accepted ${countTables(run.acceptedXml)}, rejected ${countTables(run.rejectedXml)})`);
        }
    });

    await check('Bold Signature Fields', async () => {
        const name = 'Bold Signature Fields';
        const labels = (p) => (p.exactText.startsWith('By:') ? 'By:' : 'Title:');
        const run = await applyToSampleDocument(name, (source) => findParagraphs(
            source, p => p.inTable && /^(By:|Title:)/.test(p.exactText), name, 4
        ).map(p => ({ type: 'format', target: strongTarget(p), textToFormat: labels(p), properties: { bold: true } })));
        const targets = run.source.map((p, from) => ({ p, from })).filter(({ p }) => p.inTable && /^(By:|Title:)/.test(p.exactText));
        verifyViews(run, {
            expected: buildExpected(run.source), touched: targets.map(t => t.from),
            formatEdits: targets.map(({ p, from }) => ({ from, find: labels(p), prop: 'b', value: true }))
        });
    });

    await check('Add Date Row', async () => {
        const name = 'Add Date Row';
        // Known library defect in @ansonlai/docx-redline-js 0.9.0 (docs/library-issues/2026-10-12-
        // table-reconciliation-underscore-collapse.md): the Markdown pre-processor collapses literal
        // underscore signature lines in Accept All, including cells the edit did not change.
        // The strict assertions stay; the case is reported (not failed) only when the sole mismatch
        // is that signature, and passes normally with no marker once the library is fixed.
        const blank = '___________________';
        const run = await applyToSampleDocument(name, (source) => {
            const [header] = findParagraphs(source, p => p.inTable && p.exactText === 'DISCLOSING PARTY:', name, 1);
            return [{
                type: 'table-reconciliation', target: strongTarget(header), structuredContent: true,
                modified: [
                    '| DISCLOSING PARTY: | RECEIVING PARTY: |', '| --- | --- |',
                    '| _________________________ | _________________________ |',
                    '| By: [Name] | By: [Name] |', '| Title: | Title: |', `| Date: ${blank} | Date: ${blank} |`
                ].join('\n')
            }];
        });
        const tableParagraphs = run.source.map((p, from) => ({ p, from })).filter(({ p }) => p.inTable);
        const last = tableParagraphs[tableParagraphs.length - 1].from;
        const expected = buildExpected(run.source, new Map([[last, [
            { text: run.source[last].exactText, from: last },
            { text: `Date: ${blank}`, from: null }, { text: `Date: ${blank}`, from: null }
        ]]]));
        verifyViews(run, {
            expected,
            touched: tableParagraphs.map(t => t.from)
        });
        [1, 2].forEach((cellIndex) => {
            const cell = run.accepted[last + cellIndex];
            if (!cell.inTable || cell.table.rowIndex !== 5 || cell.table.cellIndex !== cellIndex) {
                throw new Error(`${name}: new Date cell ${cellIndex} landed at ${JSON.stringify(cell.table)}`);
            }
        });
    });

    await check('Add Confidential Info List Item', async () => {
        const name = 'Add Confidential Info List Item';
        // "Technical data, ..." belongs to the sample's numbered (decimal) Confidential Information
        // list, so the new entry is a sibling item of that list, written with decimal markers the way
        // production does for insert_list_item. A bullet marker here would convert the list's kind.
        const added = 'Photographs, videos, and other recordings of prototypes and physical hardware.';
        const run = await applyToSampleDocument(name, (source) => {
            const [item] = findParagraphs(source, p => p.exactText.startsWith('Technical data, specifications'), name, 1);
            return [{
                type: 'list-change', target: strongTarget(item), structuredContent: true,
                modified: `1. ${item.exactText}\n2. ${added}`
            }];
        });
        const from = run.source.findIndex(p => p.exactText.startsWith('Technical data, specifications'));
        verifyViews(run, {
            expected: buildExpected(run.source, new Map([[from, [
                { text: run.source[from].exactText, from }, { text: added, from: null }
            ]]])),
            touched: [from]
        });
        const sibling = run.source[from].list;
        const labelAfter = String(Number.parseInt(sibling.label, 10) + 1) + '.';
        assertList(name, 'existing item', run.accepted[from], { numId: sibling.numId, level: sibling.level, label: sibling.label, format: sibling.format });
        assertList(name, 'new item', run.accepted[from + 1], { numId: sibling.numId, level: sibling.level, label: labelAfter, format: sibling.format });
        const labels = run.accepted.filter(p => p.list?.numId === sibling.numId).map(p => p.list.label);
        const wantLabels = labels.map((_, i) => `${i + 1}.`);
        if (labels.join(',') !== wantLabels.join(',')) throw new Error(`${name}: list ${sibling.numId} labels expected ${wantLabels} but got ${labels}`);
    });

    await check('Change Archival Copies', async () => {
        const name = 'Change Archival Copies';
        const find = 'one (1) copy';
        const replace = 'two (2) copies';
        const run = await applyToSampleDocument(name, (source) => {
            const [target] = findParagraphs(source, p => p.exactText.includes(find), name, 1);
            return [{ type: 'redline', target: strongTarget(target), structuredContent: false, replacements: [{ find, replace }] }];
        });
        const from = run.source.findIndex(p => p.exactText.includes(find));
        verifyViews(run, {
            expected: buildExpected(run.source, new Map([[from, [{ text: run.source[from].exactText.replace(find, replace), from }]]])),
            touched: [from]
        });
        if (!/<w:ins[\s>]/.test(run.result.documentXml) || !/<w:del[\s>]/.test(run.result.documentXml)) {
            throw new Error(`${name}: expected tracked insertion and deletion markup`);
        }
    });

    await check('Add Sub-bullet Archival', async () => {
        const name = 'Add Sub-bullet Archival';
        const added = 'Must be legally required.';
        const run = await applyToSampleDocument(name, (source) => {
            const [item] = findParagraphs(source, p => p.exactText.endsWith('compliance and record-keeping.'), name, 1);
            return [{
                type: 'list-change', target: strongTarget(item), structuredContent: true,
                modified: `1. ${item.exactText}\n    1. ${added}`
            }];
        });
        const from = run.source.findIndex(p => p.exactText.endsWith('compliance and record-keeping.'));
        verifyViews(run, {
            expected: buildExpected(run.source, new Map([[from, [
                { text: run.source[from].exactText, from }, { text: added, from: null }
            ]]])),
            touched: [from]
        });
        const parent = run.source[from].list;
        assertList(name, 'parent item', run.accepted[from], { numId: parent.numId, level: parent.level });
        assertList(name, 'sub-bullet', run.accepted[from + 1], { numId: parent.numId, level: parent.level + 1, label: `${parent.label}1.`, format: parent.format });
        assertList(name, 'following item', run.accepted[from + 2], { numId: run.source[from + 1].list.numId, level: run.source[from + 1].list.level, label: run.source[from + 1].list.label });
    });

    await check('Unbold British Columbia', async () => {
        const name = 'Unbold British Columbia';
        const find = 'British Columbia';
        const run = await applyToSampleDocument(name, (source) => {
            const [target] = findParagraphs(source, p => p.exactText.includes(find), name, 1);
            return [{ type: 'format', target: strongTarget(target), textToFormat: find, properties: { bold: false } }];
        });
        const from = run.source.findIndex(p => p.exactText.includes(find));
        const before = characterFormats(run.sourceXml)[from];
        const start = plainText(before).indexOf(find);
        if (!before.slice(start, start + find.length).every(c => c.b)) throw new Error(`${name}: sample precondition failed, first "${find}" is not bold`);
        verifyViews(run, {
            expected: buildExpected(run.source), touched: [from],
            formatEdits: [{ from, find, prop: 'b', value: false }]
        });
    });

    await check('Rewrite Disclosure Provision', async () => {
        const name = 'Rewrite Disclosure Provision';
        const newSection = 'The Receiving Party agrees to fight any orders to disclose Confidential Information. They will involve the Disclosing Party immediately. Only after all legal avenues are exhausted may they disclose the information.';
        const run = await applyToSampleDocument(name, (source) => {
            const [target] = findParagraphs(source, p => p.exactText.startsWith('If the Receiving Party is required by law, regulation, or court order'), name, 1);
            return [{ type: 'redline', target: strongTarget(target), modified: newSection, structuredContent: false }];
        });
        const from = run.source.findIndex(p => p.exactText.startsWith('If the Receiving Party is required by law, regulation, or court order'));
        verifyViews(run, {
            expected: buildExpected(run.source, new Map([[from, [{ text: newSection, from }]]])),
            touched: [from]
        });
        if (!/<w:del[\s>]/.test(run.result.documentXml)) throw new Error(`${name}: expected tracked deletion markup`);
    });

    return failures;
}

// --- Redline Toggle Tests ---
async function runRedlineToggleTests() {
    console.log('\nStarting Redline Toggle Regression Tests...');

    const originalOoxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:t>Hello world</w:t>
            </w:r>
        </w:p>
    `;
    const originalText = "Hello world";
    const modifiedText = "Hello Gemini";

    console.log('\n--- Test 1: Redlines ENABLED ---');
    const resultEnabled = await applyRedlineToOxml(originalOoxml, originalText, modifiedText, {
        author: 'TestUser',
        generateRedlines: true
    });
    assertSuccessfulResult(resultEnabled, 'Redlines enabled');
    if (resultEnabled.oxml.includes('<w:ins') && resultEnabled.oxml.includes('<w:del')) {
        console.log('✅ PASS: Track changes generated in redline mode');
    } else {
        throw new Error('Track changes were not generated in redline mode');
    }

    console.log('\n--- Test 2: Redlines DISABLED ---');
    const resultDisabled = await applyRedlineToOxml(originalOoxml, originalText, modifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });
    assertSuccessfulResult(resultDisabled, 'Redlines disabled');
    const hasInsDisabled = resultDisabled.oxml.includes('<w:ins');
    const hasDelDisabled = resultDisabled.oxml.includes('<w:del');
    const hasNewText = resultDisabled.oxml.includes('Gemini');

    if (!hasInsDisabled && !hasDelDisabled && hasNewText) {
        console.log('✅ PASS: No track changes generated when disabled');
    } else {
        console.log(`❌ FAIL: Redline toggle NOT honored. Ins: ${hasInsDisabled}, Del: ${hasDelDisabled}, NewText: ${hasNewText}`);
        console.log('Partial Output:', resultDisabled.oxml.substring(0, 200));
        throw new Error('Redline toggle was not honored');
    }


    console.log('\n--- Test 3: List Expansion with Redlines DISABLED ---');
    const listModifiedText = "Original text\n* Item 1\n* Item 2";
    const resultListDisabled = await applyRedlineToOxml(originalOoxml, originalText, listModifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });
    assertSuccessfulResult(resultListDisabled, 'List expansion with redlines disabled');
    if (!resultListDisabled.oxml.includes('<w:ins') && resultListDisabled.oxml.includes('Item 1')) {
        console.log('✅ PASS: List expansion honors redline toggle');
    } else {
        throw new Error('List expansion redline toggle issue');
    }

    console.log('\n--- Test 4: Table Reconciliation with Redlines DISABLED ---');
    const tableOriginalOoxml = `
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:tblGrid><w:gridCol/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
    `;
    const tableOriginalText = "Header 1";
    const tableModifiedText = "| Header Updated |\n| --- |";
    const resultTableDisabled = await applyRedlineToOxml(tableOriginalOoxml, tableOriginalText, tableModifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });
    assertSuccessfulResult(resultTableDisabled, 'Table reconciliation with redlines disabled');

    // Strict check to avoid matching <w:insideH> in table properties
    const hasInsTable = /<w:ins\b/.test(resultTableDisabled.oxml);
    const hasDelTable = /<w:del\b/.test(resultTableDisabled.oxml);
    const hasUpdatedText = resultTableDisabled.oxml.includes('Header Updated');

    if (!hasInsTable && !hasDelTable && hasUpdatedText) {

        console.log('✅ PASS: Table reconciliation honors redline toggle');
    } else {
        console.log(`❌ FAIL: Table reconciliation redline toggle issue. Ins: ${hasInsTable}, Del: ${hasDelTable}, UpdatedText: ${hasUpdatedText}`);
        console.log('Output Preview:', resultTableDisabled.oxml.substring(0, 500));
        console.log('Includes Header Updated:', resultTableDisabled.oxml.includes('Header Updated'));
        throw new Error('Table reconciliation redline toggle issue');
    }

}

// --- Main Runner ---
(async () => {
    const ndaFailures = await runNdaTests();
    await runRedlineToggleTests();
    if (ndaFailures.length) {
        console.error(`\n${ndaFailures.length} NDA case(s) FAILED: ${ndaFailures.map(f => f.name).join(', ')}`);
        for (const failure of ndaFailures) console.error(`  - ${failure.name}: ${failure.error.message}`);
        process.exit(1);
    }
    console.log('\nALL INTEGRATION TESTS COMPLETE.');
})().catch(error => {
    console.error('INTEGRATION TEST FAILURE:', error?.stack || error);
    process.exit(1);
});
