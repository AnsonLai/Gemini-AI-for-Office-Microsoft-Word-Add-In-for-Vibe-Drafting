import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parse as parseJavaScript } from 'acorn';
import { readFile } from 'node:fs/promises';
import { createDocumentSession, seedParagraphMarkers } from '../browser-demo/document-session.js';
import {
    openDocx,
    parseParagraphReference,
    splitLeadingParagraphMarker,
    stripLeadingParagraphMarker
} from '@ansonlai/docx-redline-js';

const FIXTURE = fileURLToPath(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url));
const inputBytes = new Uint8Array(readFileSync(FIXTURE));

function paragraphTarget(inspection, exactText) {
    const paragraph = inspection.paragraphs.find(item => item.exactText === exactText);
    assert.ok(paragraph, `fixture target not found: ${exactText}`);
    return {
        index: paragraph.index,
        exactText: paragraph.exactText,
        ...(paragraph.paragraphId ? { paragraphId: paragraph.paragraphId } : {}),
        ...(paragraph.fingerprint ? { fingerprint: paragraph.fingerprint } : {}),
        inTable: paragraph.inTable
    };
}

async function testOpenInspectAndAtomicMixedBatchRoundTrip() {
    const session = createDocumentSession(inputBytes);
    const before = session.inspect();
    assert.equal(before.status, 'ok');
    assert.ok(before.paragraphs.length > 10);
    assert.equal(before.paragraphs[2].list?.format, 'bullet');

    const promptParagraphs = session.getPromptParagraphs();
    assert.equal(promptParagraphs[0].index, 1);
    assert.equal(promptParagraphs[0].text, 'Untouched bold sentinel');
    assert.match(promptParagraphs[0].formattedText, /\*\*Untouched bold sentinel\*\*/,
        'the prompt projection retains run formatting without editing package XML');
    assert.equal(promptParagraphs[2].formattedText, '- Bullet Root A');
    assert.equal(promptParagraphs[3].formattedText, '  - Bullet Insertion Anchor',
        'the prompt projection keeps nested bullet depth');
    assert.equal(promptParagraphs[7].formattedText, '1. Number Root A');
    assert.equal(promptParagraphs[8].formattedText, '  1.1. Number Nested Anchor',
        'ordered labels and nested levels come from document inspection');
    assert.equal(promptParagraphs[9].formattedText, '2. Number Root B');
    assert.equal(promptParagraphs[11].formattedText, '3. Number Continued Item',
        'numbering continuation comes from inspection rather than isolated paragraph parsing');

    const result = await session.applyOperations([
        {
            type: 'redline',
            target: paragraphTarget(before, 'Plain paragraph before bullet list.'),
            replacements: [{ find: 'before', replace: 'ahead of' }]
        },
        {
            type: 'comment',
            target: paragraphTarget(before, 'Bullet Root A'),
            textToComment: 'Root',
            commentContent: 'Please review this list item.'
        },
        {
            type: 'highlight',
            target: paragraphTarget(before, 'Number Nested Anchor'),
            textToHighlight: 'Nested',
            color: 'yellow'
        }
    ], { author: 'Browser Session Test', generateRedlines: true, sanitizeInput: true });

    assert.equal(result.status, 'ok', JSON.stringify(result.error || result.results));
    assert.equal(result.written, true);
    assert.equal(result.results.length, 3);
    assert.ok(result.results.every(item => item.status !== 'error'));
    assert.ok(result.receipts.some(receipt => receipt.commentIds.length > 0));

    const outputBytes = session.toUint8Array();
    assert.ok(outputBytes instanceof Uint8Array);
    const reopened = openDocx(outputBytes);
    const after = reopened.inspect();
    assert.equal(after.status, 'ok');
    assert.equal(after.paragraphs[1].exactText, 'Plain paragraph ahead of bullet list.');
    assert.equal(after.paragraphs[1].hasRevisions, true);
    assert.equal(after.comments.length, 1);
    assert.equal(after.comments[0].anchoredText, 'Root');
    assert.ok(after.paragraphs[8].hasRevisions, 'the highlight remains represented as a tracked formatting revision');

    for (const partName of [
        '[Content_Types].xml', 'word/document.xml', 'word/numbering.xml',
        'word/styles.xml', 'word/_rels/document.xml.rels', 'word/comments.xml'
    ]) {
        assert.ok(reopened.entries.has(partName), `saved package lost ${partName}`);
    }
}

function extractProductionFunctions(source, names) {
    const ast = parseJavaScript(source, { ecmaVersion: 'latest', sourceType: 'module' });
    const wanted = new Set(names);
    const nodes = (ast.body || []).filter(node => node.type === 'FunctionDeclaration' && wanted.has(node.id?.name));
    assert.equal(nodes.length, wanted.size, 'all requested production function declarations were found');
    return nodes.map(node => source.slice(node.start, node.end)).join('\n');
}

async function testProductionChatParserAndReceiptSummary() {
    const demoSource = await readFile(new URL('../browser-demo/demo.js', import.meta.url), 'utf8');
    const source = extractProductionFunctions(demoSource, [
        'parseGeminiChatResponse', 'applyChatOperations', 'buildOpSummaryHtml', 'escapeHtml'
    ]);
    const { parseGeminiChatResponse, applyChatOperations, buildOpSummaryHtml } = new Function(
        'ALLOWED_HIGHLIGHT_COLORS', 'splitLeadingParagraphMarker', 'parseParagraphReference',
        'stripLeadingParagraphMarker', 'shouldGenerateRedlines', 'log',
        `${source}\nreturn { parseGeminiChatResponse, applyChatOperations, buildOpSummaryHtml };`
    )(
        ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red'],
        splitLeadingParagraphMarker,
        parseParagraphReference,
        stripLeadingParagraphMarker,
        () => true,
        () => {}
    );

    const parsed = parseGeminiChatResponse('Localized edit.\n---OPERATIONS---\n' + JSON.stringify([
        {
            type: 'redline', targetRef: 'P3', target: '[P3] The term renews annually.',
            replacements: [{ find: 'annually', replace: 'each year', occurrence: 1 }]
        },
        {
            type: 'redline', targetRef: 'P4', target: 'Other target', modified: 'Whole paragraph',
            replacements: [{ find: 'Other', replace: 'Different' }]
        }
    ]));
    assert.equal(parsed.operations.length, 1, 'the parser accepts localized replacements and rejects mixed forms');
    assert.deepEqual(parsed.operations[0].replacements, [
        { find: 'annually', replace: 'each year', occurrence: 1 }
    ]);
    assert.equal(parsed.operations[0].target, 'The term renews annually.');
    assert.equal(parsed.operations[0].targetRef, 3);

    const operations = [parsed.operations[0], {
        type: 'highlight', target: 'Missing paragraph', textToHighlight: 'missing', color: 'yellow'
    }];
    const applied = await applyChatOperations({
        async applyOperations(requested, options) {
            assert.deepEqual(requested, operations);
            assert.equal(options.strictTargets, undefined, 'the session owns its strict-target default');
            return {
                status: 'ok',
                results: [
                    { index: 1, type: 'redline', status: 'applied' },
                    { index: 2, type: 'highlight', status: 'error', error: { message: 'Target not found.' } }
                ]
            };
        }
    }, operations, 'Test Author', 'redline');
    assert.deepEqual(applied.operationResults.map(item => item.success), [true, false],
        'the UI uses the facade receipt status instead of an undocumented hasChanges field');
    assert.equal(applied.operationResults[1].error, 'Target not found.');
    assert.match(buildOpSummaryHtml(applied.operationResults), /1\/2 operations applied/);
}

async function testDirectEditPreservesForeignRevisions() {
    const session = createDocumentSession(inputBytes);
    const initial = session.inspect();
    const priorEdit = await session.applyOperations([{
        type: 'redline',
        target: paragraphTarget(initial, 'Untouched bold sentinel'),
        modified: 'Untouched bold sentinel with a prior change'
    }], { author: 'Prior Reviewer', generateRedlines: true });
    assert.equal(priorEdit.status, 'ok', JSON.stringify(priorEdit.error || priorEdit.results));

    const directTarget = session.inspect();
    const directEdit = await session.applyOperations([{
        type: 'redline',
        target: paragraphTarget(directTarget, 'Plain paragraph before bullet list.'),
        replacements: [{ find: 'before', replace: 'ahead of' }]
    }], { author: 'Browser Session Test', generateRedlines: false });
    assert.equal(directEdit.status, 'ok', JSON.stringify(directEdit.error || directEdit.results));

    const reopened = openDocx(session.toUint8Array());
    const accepted = reopened.inspect();
    const rejected = reopened.inspect({ revisionView: 'rejected' });
    assert.ok(accepted.paragraphs[0].revisionAuthors.includes('Prior Reviewer'));
    assert.equal(rejected.paragraphs[0].exactText, 'Untouched bold sentinel');
    assert.equal(accepted.paragraphs[1].exactText, 'Plain paragraph ahead of bullet list.');
    assert.equal(accepted.paragraphs[1].hasRevisions, false,
        'direct mode applies plain text and leaves the other author’s pending revision intact');
}

async function testFailedBatchLeavesSessionBytesUntouched() {
    const session = createDocumentSession(inputBytes);
    const before = session.toUint8Array();
    const inspection = session.inspect();
    const result = await session.applyOperations([
        {
            type: 'redline',
            target: paragraphTarget(inspection, 'Plain paragraph before bullet list.'),
            modified: 'This edit must roll back.'
        },
        {
            type: 'highlight',
            target: 'A paragraph absent from this document',
            textToHighlight: 'absent',
            color: 'yellow'
        }
    ], { author: 'Browser Session Test', generateRedlines: true });

    assert.equal(result.status, 'error');
    assert.equal(result.written, false);
    assert.equal(result.rolledBack, true);
    assert.deepEqual(session.toUint8Array(), before,
        'the facade session must remain unchanged when an atomic batch is refused');
    assert.equal(session.inspect().paragraphs[1].exactText, 'Plain paragraph before bullet list.');
}

async function testKitchenSinkMarkerSeedingUsesStructuredFacadeOperation() {
    const session = createDocumentSession(inputBytes);
    const before = session.inspect();
    const markers = [
        'DEMO TEXT TARGET',
        'DEMO FORMAT TARGET',
        'DEMO LIST TARGET',
        'DEMO TABLE TARGET'
    ];
    assert.equal(before.status, 'ok');
    assert.equal(before.paragraphs.length, 24);

    const seeded = await seedParagraphMarkers(session, markers, 'Kitchen Sink Test');
    assert.deepEqual(seeded.added, markers);
    assert.equal(seeded.inspection.status, 'ok');
    assert.equal(seeded.inspection.paragraphs.length, 28,
        'one structured facade batch turns four newline-separated markers into four paragraphs');
    for (const marker of markers) {
        assert.equal(seeded.inspection.paragraphs.filter(paragraph => paragraph.exactText === marker).length, 1,
            `marker was created as its own paragraph: ${marker}`);
    }
    const beforeCounts = new Map();
    const afterCounts = new Map();
    for (const paragraph of before.paragraphs) beforeCounts.set(paragraph.exactText, (beforeCounts.get(paragraph.exactText) || 0) + 1);
    for (const paragraph of seeded.inspection.paragraphs) afterCounts.set(paragraph.exactText, (afterCounts.get(paragraph.exactText) || 0) + 1);
    for (const [text, count] of beforeCounts) {
        assert.equal(afterCounts.get(text), count, `source paragraph is preserved: ${text}`);
    }
    assert.equal(seeded.result.results[0].status, 'applied');
}

async function testKitchenSinkKeepsLegacyUnderscoreMarkerAnchors() {
    const demoSource = await readFile(new URL('../browser-demo/demo.js', import.meta.url), 'utf8');
    const source = extractProductionFunctions(demoSource, ['seedMissingDemoTargets']);
    const seedMissingDemoTargets = new Function(
        'DEMO_MARKERS', 'DEMO_MARKER_ALIASES', 'seedParagraphMarkers', 'log',
        `${source}\nreturn seedMissingDemoTargets;`
    )(
        ['DEMO TEXT TARGET', 'DEMO FORMAT TARGET', 'DEMO LIST TARGET', 'DEMO TABLE TARGET'],
        {
            'DEMO TEXT TARGET': ['DEMO_TEXT_TARGET'],
            'DEMO LIST TARGET': ['DEMO_LIST_TARGET'],
            'DEMO TABLE TARGET': ['DEMO_TABLE_TARGET']
        },
        seedParagraphMarkers,
        () => {}
    );
    const session = {
        inspect() {
            return {
                status: 'ok',
                paragraphs: [
                    'DEMO_TEXT_TARGET', 'DEMO FORMAT TARGET', 'DEMO LIST TARGET', 'DEMO TABLE TARGET'
                ].map((exactText, index) => ({ index: index + 1, exactText, inTable: false, list: null }))
            };
        },
        async applyOperations() {
            assert.fail('all legacy and current markers already exist, so seeding should not write');
        }
    };
    const resolved = await seedMissingDemoTargets(session, 'Kitchen Sink Alias Test');
    assert.equal(resolved['DEMO TEXT TARGET'], 'DEMO_TEXT_TARGET',
        'legacy underscore target text remains resolvable and is not replaced with a new marker');
}

try {
    assert.equal(typeof globalThis.Word, 'undefined', 'the session must not depend on Word globals');
    await testOpenInspectAndAtomicMixedBatchRoundTrip();
    await testDirectEditPreservesForeignRevisions();
    await testFailedBatchLeavesSessionBytesUntouched();
    await testKitchenSinkMarkerSeedingUsesStructuredFacadeOperation();
    await testKitchenSinkKeepsLegacyUnderscoreMarkerAnchors();
    await testProductionChatParserAndReceiptSummary();
    console.log('PASS: browser document session facade tests');
} catch (error) {
    console.error('FAIL:', error?.stack || error?.message || error);
    process.exitCode = 1;
}
