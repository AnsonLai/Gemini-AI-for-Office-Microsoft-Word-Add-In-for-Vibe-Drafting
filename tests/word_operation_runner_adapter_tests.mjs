import './setup-xml-provider.mjs';

import assert from 'assert';
import {
    applyWordOperation
} from '../src/taskpane/modules/reconciliation/integration/word-operation-runner.js';
import {
    extractReplacementNodesFromOoxml,
    getParagraphText
} from '../src/taskpane/modules/reconciliation/standalone.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

if (!global.Word) {
    global.Word = {};
}
if (!global.Word.ChangeTrackingMode) {
    global.Word.ChangeTrackingMode = {
        off: 'Off',
        trackAll: 'TrackAll'
    };
}
if (!global.Word.InsertLocation) {
    global.Word.InsertLocation = {
        replace: 'Replace',
        Replace: 'Replace'
    };
}

function createMockContext(initialTrackingMode = global.Word.ChangeTrackingMode.trackAll) {
    const syncCalls = [];
    const context = {
        document: {
            changeTrackingMode: initialTrackingMode,
            load(property) {
                syncCalls.push({ type: 'load', property });
            }
        },
        async sync() {
            syncCalls.push({ type: 'sync' });
        }
    };
    return { context, syncCalls };
}

function createMockParagraph(paragraphOoxml, options = {}) {
    const insertCalls = [];
    const paragraph = {
        getOoxml() {
            return { value: paragraphOoxml };
        },
        insertOoxml(payload, mode) {
            if (options.insertError) {
                const error = new Error(options.insertError.message || 'insert failed');
                if (options.insertError.code) error.code = options.insertError.code;
                throw error;
            }
            insertCalls.push({ payload, mode, via: 'paragraph' });
        },
        getRange(rangeKind = 'Whole') {
            return {
                insertOoxml(payload, mode) {
                    insertCalls.push({ payload, mode, via: `range:${rangeKind}` });
                }
            };
        }
    };
    return { paragraph, insertCalls };
}

function createMockRange(scopeOoxml) {
    const insertCalls = [];
    const range = {
        getOoxml() {
            return { value: scopeOoxml };
        },
        insertOoxml(payload, mode) {
            insertCalls.push({ payload, mode });
        }
    };
    return { range, insertCalls };
}

function buildParagraphXml(text, options = {}) {
    if (options.listNumId) {
        const ilvl = Number.isInteger(options.listIlvl) ? options.listIlvl : 0;
        return `<w:p xmlns:w="${NS_W}"><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${options.listNumId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    }
    return `<w:p xmlns:w="${NS_W}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function buildScopeDocumentXml(paragraphXmlList) {
    const bodyXml = paragraphXmlList
        .map(xml => xml.replace(` xmlns:w="${NS_W}"`, ''))
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${bodyXml}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function extractTextsFromPayload(payloadOoxml) {
    const xml = String(payloadOoxml || '');
    if (!xml) return [];

    if (xml.includes('<pkg:package')) {
        const extracted = extractReplacementNodesFromOoxml(xml);
        return (extracted.replacementNodes || [])
            .filter(node => node && node.namespaceURI === NS_W && node.localName === 'p')
            .map(node => getParagraphText(node).trim())
            .filter(Boolean);
    }

    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const paragraph = doc.getElementsByTagNameNS(NS_W, 'p')[0];
    if (!paragraph) return [];
    return [getParagraphText(paragraph).trim()].filter(Boolean);
}

async function testSingleParagraphRedlineApply() {
    const original = 'Alpha target text.';
    const modified = 'Alpha target text updated.';
    const { paragraph, insertCalls } = createMockParagraph(buildParagraphXml(original));
    const { context } = createMockContext();

    const applied = await applyWordOperation(
        context,
        {
            type: 'redline',
            targetRef: 'P1',
            target: original,
            modified
        },
        { paragraph },
        {
            author: 'AdapterTest',
            generateRedlines: false,
            disableNativeTracking: true,
            baseTrackingMode: global.Word.ChangeTrackingMode.trackAll
        }
    );

    assert.strictEqual(applied, true, 'single paragraph redline should apply');
    assert.strictEqual(insertCalls.length, 1, 'single paragraph redline should perform one OOXML insertion');
    const texts = extractTextsFromPayload(insertCalls[0].payload);
    assert.strictEqual(texts.includes(modified), true, 'single paragraph redline output should contain modified text');
}

async function testAdapterCallsRunnerAndNoOpsWithoutChanges() {
    const original = 'No changes';
    const { paragraph, insertCalls } = createMockParagraph(buildParagraphXml(original));
    const { context } = createMockContext();
    const runnerCalls = [];

    const applied = await applyWordOperation(
        context,
        {
            type: 'redline',
            targetRef: 'P1',
            target: original,
            modified: original
        },
        { paragraph },
        {
            author: 'AdapterTest',
            generateRedlines: false,
            runner: async (...args) => {
                runnerCalls.push(args);
                return { hasChanges: false, warnings: ['no changes'] };
            }
        }
    );

    assert.strictEqual(applied, false, 'adapter should return false when runner reports no changes');
    assert.strictEqual(runnerCalls.length, 1, 'adapter should call shared standalone runner');
    assert.strictEqual(insertCalls.length, 0, 'adapter should not insert OOXML when no changes are produced');
}

async function testRangeListInsertionRedlineApply() {
    const existingItems = [
        'Business plans, strategies, financial information, pricing, and marketing data.',
        'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.',
        "Information concerning the Disclosing Party's employees, contractors, customers, and suppliers.",
        'Any notes, analyses, compilations, studies, or other materials prepared by the Receiving Party that contain, reflect, or are derived from the foregoing.'
    ];
    const insertedItem = 'Photographs, videos, and other recordings of prototypes and physical hardware.';
    const modified = [
        `1. ${existingItems[0]}`,
        `2. ${insertedItem}`,
        `3. ${existingItems[1]}`,
        `4. ${existingItems[2]}`,
        `5. ${existingItems[3]}`
    ].join('\n');

    const scopeXml = buildScopeDocumentXml([
        buildParagraphXml(existingItems[0], { listNumId: '77', listIlvl: 0 }),
        buildParagraphXml(existingItems[1], { listNumId: '77', listIlvl: 0 }),
        buildParagraphXml(existingItems[2], { listNumId: '77', listIlvl: 0 }),
        buildParagraphXml(existingItems[3], { listNumId: '77', listIlvl: 0 })
    ]);
    const { range, insertCalls } = createMockRange(scopeXml);
    const { context } = createMockContext();

    const applied = await applyWordOperation(
        context,
        {
            type: 'redline',
            targetRef: 'P1',
            targetEndRef: 'P4',
            target: existingItems[0],
            modified
        },
        { range },
        {
            author: 'AdapterTest',
            generateRedlines: true,
            disableNativeTracking: true,
            baseTrackingMode: global.Word.ChangeTrackingMode.trackAll
        }
    );

    assert.strictEqual(applied, true, 'range list insertion redline should apply');
    assert.strictEqual(insertCalls.length, 1, 'range list insertion should insert one package');
    const texts = extractTextsFromPayload(insertCalls[0].payload);
    assert.strictEqual(texts.length, 5, 'range list insertion should output five list paragraphs');
    assert.strictEqual(texts[1], insertedItem, 'range list insertion should place new item at index 2');
}

async function testSingleParagraphConcatenationInsertionShape() {
    const original = 'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.';
    const inserted = 'Photographs, videos, and other recordings of prototypes and physical hardware.';
    const modified = `${inserted}${original}`;
    const { paragraph, insertCalls } = createMockParagraph(buildParagraphXml(original, { listNumId: '77', listIlvl: 0 }));
    const { context } = createMockContext();

    const applied = await applyWordOperation(
        context,
        {
            type: 'redline',
            targetRef: 'P1',
            target: original,
            modified
        },
        { paragraph },
        {
            author: 'AdapterTest',
            generateRedlines: true,
            disableNativeTracking: true,
            baseTrackingMode: global.Word.ChangeTrackingMode.trackAll
        }
    );

    assert.strictEqual(applied, true, 'single-paragraph concatenation insertion shape should apply');
    assert.strictEqual(insertCalls.length, 1, 'single-paragraph insertion shape should insert one package');
    const texts = extractTextsFromPayload(insertCalls[0].payload);
    assert.strictEqual(texts.length >= 2, true, 'single-paragraph insertion shape should preserve multi-paragraph redline output');
    assert.strictEqual(texts.includes(inserted), true, 'single-paragraph insertion shape should include inserted item text');
    assert.strictEqual(texts.includes(original), true, 'single-paragraph insertion shape should retain original item text');
}

async function testSingleParagraphPlainInsertionBeforeShape() {
    const original = 'NON-DISCLOSURE AGREEMENT';
    const insertedMarkdown = '**INSTRUCTIONS:** Please review this Non-Disclosure Agreement carefully.';
    const insertedPlain = 'INSTRUCTIONS: Please review this Non-Disclosure Agreement carefully.';
    const modified = `${insertedMarkdown}\n${original}`;
    const { paragraph, insertCalls } = createMockParagraph(buildParagraphXml(original));
    const { context } = createMockContext();

    const applied = await applyWordOperation(
        context,
        {
            type: 'redline',
            targetRef: 'P1',
            target: original,
            modified
        },
        { paragraph },
        {
            author: 'AdapterTest',
            generateRedlines: true,
            disableNativeTracking: true,
            baseTrackingMode: global.Word.ChangeTrackingMode.trackAll
        }
    );

    assert.strictEqual(applied, true, 'single-paragraph plain insertion-before shape should apply');
    assert.strictEqual(insertCalls.length, 1, 'single-paragraph plain insertion-before shape should insert one package');
    const texts = extractTextsFromPayload(insertCalls[0].payload);
    assert.strictEqual(
        texts.length >= 2,
        true,
        'single-paragraph plain insertion-before shape should preserve multi-paragraph output'
    );
    assert.strictEqual(texts[0], insertedPlain, 'plain insertion-before shape should place inserted paragraph first');
    assert.strictEqual(texts.includes(original), true, 'plain insertion-before shape should retain original paragraph text');
    const payload = String(insertCalls[0].payload || '');
    assert.strictEqual(payload.includes('<w:b'), true, 'plain insertion-before shape should render markdown bold formatting');
}

async function testRangeTableReplaceApply() {
    const scopeXml = buildScopeDocumentXml([
        buildParagraphXml('Disclosing Party: [Name of Disclosing Party]'),
        buildParagraphXml('And'),
        buildParagraphXml('Receiving Party: [Name of Receiving Party]')
    ]);
    const { range, insertCalls } = createMockRange(scopeXml);
    const { context } = createMockContext();

    const applied = await applyWordOperation(
        context,
        {
            type: 'redline',
            targetRef: 'P1',
            targetEndRef: 'P3',
            target: 'Disclosing Party: [Name of Disclosing Party]',
            modified: [
                '| Disclosing Party: | [Name of Disclosing Party] [Address of Disclosing Party] (the "Disclosing Party") |',
                '| Receiving Party: | [Name of Receiving Party] [Address of Receiving Party] (the "Receiving Party") |'
            ].join('\n')
        },
        { range },
        {
            author: 'AdapterTest',
            generateRedlines: true
        }
    );

    assert.strictEqual(applied, true, 'range table replacement should apply');
    assert.strictEqual(insertCalls.length, 1, 'range table replacement should insert one package');
    const extracted = extractReplacementNodesFromOoxml(insertCalls[0].payload);
    const nodeNames = (extracted.replacementNodes || []).map(node => node.localName);
    assert.strictEqual(nodeNames.includes('tbl'), true, 'range table replacement should output a table node');
}

async function testSingleParagraphHighlightApply() {
    const original = 'Highlight target text.';
    const { paragraph, insertCalls } = createMockParagraph(buildParagraphXml(original));
    const { context } = createMockContext();

    const applied = await applyWordOperation(
        context,
        {
            type: 'highlight',
            targetRef: 'P1',
            target: original,
            textToHighlight: 'target',
            color: 'yellow'
        },
        { paragraph },
        {
            author: 'AdapterTest',
            generateRedlines: false
        }
    );

    assert.strictEqual(applied, true, 'single paragraph highlight should apply');
    assert.strictEqual(insertCalls.length, 1, 'single paragraph highlight should perform one OOXML insertion');
    const payload = String(insertCalls[0].payload || '');
    assert.strictEqual(payload.includes('w:highlight'), true, 'highlight output should include w:highlight formatting');
}

async function testSingleParagraphCommentApply() {
    const original = 'Comment target text.';
    const { paragraph, insertCalls } = createMockParagraph(buildParagraphXml(original));
    const { context } = createMockContext();

    const applied = await applyWordOperation(
        context,
        {
            type: 'comment',
            targetRef: 'P1',
            target: original,
            textToComment: 'target',
            commentContent: 'Review this phrasing.'
        },
        { paragraph },
        {
            author: 'AdapterTest',
            generateRedlines: true
        }
    );

    assert.strictEqual(applied, true, 'single paragraph comment should apply');
    assert.strictEqual(insertCalls.length, 1, 'single paragraph comment should perform one OOXML insertion');
    const payload = String(insertCalls[0].payload || '');
    assert.strictEqual(payload.includes('/word/comments.xml'), true, 'comment output package should include comments part');
    assert.strictEqual(payload.includes('Review this phrasing.'), true, 'comment output should include comment text');
}

async function testInsertionErrorsPropagateWithoutLegacyFallback() {
    const original = 'Alpha target text.';
    const modified = 'Alpha target text updated.';
    const { paragraph } = createMockParagraph(buildParagraphXml(original), {
        insertError: {
            code: 'InvalidArgument',
            message: 'Synthetic insertion failure'
        }
    });
    const { context } = createMockContext();

    await assert.rejects(
        async () => applyWordOperation(
            context,
            {
                type: 'redline',
                targetRef: 'P1',
                target: original,
                modified
            },
            { paragraph },
            {
                author: 'AdapterTest',
                generateRedlines: false
            }
        ),
        /Synthetic insertion failure/,
        'adapter should surface insertion errors without legacy runtime fallback branches'
    );
}

async function run() {
    await testSingleParagraphRedlineApply();
    await testAdapterCallsRunnerAndNoOpsWithoutChanges();
    await testRangeListInsertionRedlineApply();
    await testSingleParagraphConcatenationInsertionShape();
    await testSingleParagraphPlainInsertionBeforeShape();
    await testRangeTableReplaceApply();
    await testSingleParagraphHighlightApply();
    await testSingleParagraphCommentApply();
    await testInsertionErrorsPropagateWithoutLegacyFallback();
    console.log('PASS: word operation runner adapter tests');
}

run().catch(error => {
    console.error('FAIL:', error?.message || error);
    process.exit(1);
});
