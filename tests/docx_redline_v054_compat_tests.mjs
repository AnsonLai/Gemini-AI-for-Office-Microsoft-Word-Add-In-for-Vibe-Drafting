import assert from 'assert';
import { createRequire } from 'module';
import { readFileSync, readdirSync, statSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath, pathToFileURL } from 'url';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

import {
    assertRedlineResult,
    prepareOperationInput,
    RedlineOperationError
} from '../src/taskpane/modules/docx-redline-js-integration/redline-result.js';

const EXPECTED_VERSION = '0.5.4';
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..');

function resolvePackageRoot() {
    if (process.env.DOCX_REDLINE_PACKAGE_ROOT) {
        return resolve(process.env.DOCX_REDLINE_PACKAGE_ROOT);
    }

    const require = createRequire(import.meta.url);
    return dirname(require.resolve('@ansonlai/docx-redline-js'));
}

function readPackageVersion(packageRoot) {
    return JSON.parse(readFileSync(join(packageRoot, 'package.json'), 'utf8')).version;
}

function documentXml(paragraphs) {
    return [
        `<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}">`,
        '<w:body>',
        ...paragraphs,
        '<w:sectPr/>',
        '</w:body>',
        '</w:document>'
    ].join('');
}

function paragraphXml(text, paraId = null) {
    const id = paraId ? ` w14:paraId="${paraId}"` : '';
    return `<w:p${id}><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function parseParagraph(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    return doc.getElementsByTagName('w:p')[0];
}

function canonicalDocumentText(api, xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    return api.getDocumentParagraphNodes(doc)
        .map(paragraph => api.getParagraphText(paragraph))
        .join('\n');
}

function walkJavaScriptFiles(root) {
    const files = [];
    for (const entry of readdirSync(root)) {
        const fullPath = join(root, entry);
        const stat = statSync(fullPath);
        if (stat.isDirectory()) files.push(...walkJavaScriptFiles(fullPath));
        else if (/\.(?:js|mjs)$/.test(entry)) files.push(fullPath);
    }
    return files;
}

function testNewCodesPassThroughConsumerBoundary() {
    const codes = [
        'INVALID_OPERATION',
        'COMMENTED_CONTENT_DELETE',
        'ANCHOR_NOT_FOUND',
        'AMBIGUOUS_ANCHOR',
        'PATCH_ROUNDTRIP_MISMATCH',
        'FOREIGN_PARAGRAPH_MARK_DELETION',
        'GENERATED_OOXML_INVALID',
        'REJECTED_INSERTION_STATE_REQUIRED',
        'UNSAFE_REVISION_BOUNDARY',
        'FUTURE_V05_ERROR'
    ];

    for (const code of codes) {
        assert.throws(
            () => assertRedlineResult({
                status: 'error',
                hasChanges: false,
                error: { code, message: 'Synthetic v0.5 failure' },
                warnings: ['preserved warning']
            }, 'v0.5.4 compatibility test'),
            error => {
                assert.ok(error instanceof RedlineOperationError);
                assert.strictEqual(error.code, code);
                assert.deepStrictEqual(error.details.warnings, ['preserved warning']);
                return true;
            }
        );
    }
}

function testConsumerBatchCallsRequireExplicitAtomicity() {
    const productionRoots = [
        join(REPO_ROOT, 'src'),
        join(REPO_ROOT, 'browser-demo'),
        join(REPO_ROOT, 'mcp', 'docx-server', 'src')
    ];

    for (const file of productionRoots.flatMap(walkJavaScriptFiles)) {
        const source = readFileSync(file, 'utf8');
        if (!source.includes('applyOperationsToDocumentXml')) continue;
        assert.match(
            source,
            /atomic\s*:\s*true/,
            `${file} uses applyOperationsToDocumentXml without an explicit atomic: true transaction policy`
        );
    }
}

async function loadV054(packageRoot) {
    const api = await import(pathToFileURL(join(packageRoot, 'index.js')).href);
    const runner = await import(pathToFileURL(join(
        packageRoot,
        'services',
        'standalone-operation-runner.js'
    )).href);

    api.configureXmlProvider({ DOMParser, XMLSerializer });
    api.configureLogger({ log: () => {}, warn: () => {}, error: () => {} });
    return { api, runner };
}

function testCanonicalTextAndFreshFingerprints(api) {
    const revisionParagraph = [
        '<w:p xmlns:w="', NS_W, '" xmlns:w14="', NS_W14, '" w14:paraId="A1B2C3D4">',
        '<w:r><w:t>Current</w:t><w:tab/><w:t>line</w:t><w:br/>',
        '<w:softHyphen/><w:noBreakHyphen/></w:r>',
        '<w:del w:id="1" w:author="Reviewer"><w:r><w:delText>deleted</w:delText></w:r></w:del>',
        '<w:moveFrom w:id="2" w:author="Reviewer"><w:r><w:t>moved</w:t></w:r></w:moveFrom>',
        '</w:p>'
    ].join('');
    const paragraph = parseParagraph(revisionParagraph);
    assert.strictEqual(api.getParagraphText(paragraph), 'Current\tline\n\u00ad\u2011');

    const originalFingerprint = api.createParagraphFingerprint(paragraph, { index: 1 });
    const changedCurrent = parseParagraph(revisionParagraph.replace('Current', 'Changed'));
    const changedDeleted = parseParagraph(revisionParagraph.replace('deleted', 'different deletion'));

    assert.match(originalFingerprint, /^fnv1a32:[0-9a-f]{8}$/);
    assert.notStrictEqual(
        api.createParagraphFingerprint(changedCurrent, { index: 1 }),
        originalFingerprint,
        'a current-view text change must invalidate the descriptor fingerprint'
    );
    assert.strictEqual(
        api.createParagraphFingerprint(changedDeleted, { index: 1 }),
        originalFingerprint,
        'deleted-view-only text must not change an accepted-view fingerprint'
    );
}

async function testStructuredOperationFailures(runner) {
    const source = documentXml([paragraphXml('repeat repeat.', 'AAAABBBB')]);
    const invalid = await runner.applyOperationToDocumentXml(source, {
        type: 'unsupported-operation',
        target: { index: 1 }
    }, 'AI Redliner');
    assert.strictEqual(invalid.status, 'error');
    assert.strictEqual(invalid.error?.code, 'INVALID_OPERATION');
    assert.strictEqual(invalid.documentXml, source);
    assert.strictEqual(invalid.hasChanges, false);

    const missingAnchor = await runner.applyOperationToDocumentXml(source, {
        type: 'comment',
        target: { index: 1 },
        textToComment: 'missing anchor',
        commentContent: 'Review this.'
    }, 'AI Redliner');
    assert.strictEqual(missingAnchor.status, 'error');
    assert.strictEqual(missingAnchor.error?.code, 'ANCHOR_NOT_FOUND');
    assert.strictEqual(missingAnchor.documentXml, source);

    const ambiguousAnchor = await runner.applyOperationToDocumentXml(source, {
        type: 'comment',
        target: { index: 1 },
        textToComment: 'repeat',
        commentContent: 'Review this.'
    }, 'AI Redliner');
    assert.strictEqual(ambiguousAnchor.status, 'error');
    assert.strictEqual(ambiguousAnchor.error?.code, 'AMBIGUOUS_ANCHOR');
    assert.strictEqual(ambiguousAnchor.documentXml, source);
}

async function testCommentedParagraphDeletionFailsClosed(runner) {
    const commented = documentXml([[ 
        '<w:p w14:paraId="CCCCDDDD">',
        '<w:commentRangeStart w:id="7"/>',
        '<w:r><w:t>Commented clause.</w:t></w:r>',
        '<w:commentRangeEnd w:id="7"/>',
        '<w:r><w:commentReference w:id="7"/></w:r>',
        '</w:p>'
    ].join('')]);
    const result = await runner.applyOperationToDocumentXml(commented, {
        type: 'delete',
        target: { index: 1 }
    }, 'AI Redliner');

    assert.strictEqual(result.status, 'error');
    assert.strictEqual(result.error?.code, 'COMMENTED_CONTENT_DELETE');
    assert.deepStrictEqual(result.error?.commentIds, ['7']);
    assert.strictEqual(result.documentXml, commented);
    assert.strictEqual(result.hasChanges, false);
}

async function testAtomicBatchRollsBackAllArtifacts(runner) {
    const source = documentXml([
        paragraphXml('First clause.', '11111111'),
        paragraphXml('Second clause.', '22222222')
    ]);
    const result = await runner.applyOperationsToDocumentXml(source, [
        {
            operationId: 'valid-first',
            type: 'redline',
            target: { index: 1, exactText: 'First clause.' },
            modified: 'Updated first clause.'
        },
        {
            operationId: 'invalid-second',
            type: 'not-supported',
            target: { index: 2 }
        }
    ], 'AI Redliner', null, {
        atomic: true,
        continueOnError: true,
        strictTargets: true
    });

    assert.strictEqual(result.status, 'error');
    assert.strictEqual(result.error?.code, 'BATCH_OPERATION_FAILED');
    assert.strictEqual(result.rolledBack, true);
    assert.strictEqual(result.documentXml, source);
    assert.strictEqual(result.hasChanges, false);
    assert.deepStrictEqual(result.numberingXmlParts, []);
    assert.strictEqual(result.commentsXml, null);
    assert.deepStrictEqual(result.authorsUsed, []);
    assert.deepStrictEqual(result.results.map(item => item.status), ['applied', 'error']);
    assert.strictEqual(result.results[1].error?.code, 'INVALID_OPERATION');
    assert.strictEqual(result.receipts[0].finalDisposition, 'rolled_back');
    assert.strictEqual(result.receipts[0].committed, false);
    assert.strictEqual(result.receipts[1].finalDisposition, 'refused');
    assert.strictEqual(result.receipts[1].committed, false);
}

async function testSameAuthorMergeAndForeignAuthorRefusal(api, runner) {
    const source = documentXml([paragraphXml('Original clause.', 'ABCD1234')]);
    const first = await runner.applyOperationToDocumentXml(source, {
        type: 'redline',
        target: { index: 1, exactText: 'Original clause.' },
        modified: 'First proposal.'
    }, 'AI Redliner', null, { strictTargets: true });
    assert.strictEqual(first.status, 'ok');
    assert.strictEqual(first.hasChanges, true);

    const merged = await runner.applyOperationToDocumentXml(first.documentXml, {
        type: 'redline',
        target: { index: 1, exactText: 'First proposal.' },
        modified: 'Second proposal.'
    }, 'AI Redliner', null, { strictTargets: true });
    assert.strictEqual(merged.status, 'ok');
    assert.strictEqual(canonicalDocumentText(api, merged.documentXml), 'Second proposal.');

    const foreign = await runner.applyOperationToDocumentXml(first.documentXml, {
        type: 'redline',
        target: { index: 1, exactText: 'First proposal.' },
        modified: 'Foreign proposal.'
    }, 'Another Reviewer', null, { strictTargets: true });
    assert.strictEqual(foreign.status, 'error');
    assert.strictEqual(foreign.error?.code, 'EXISTING_REVISIONS');
    assert.strictEqual(foreign.documentXml, first.documentXml);
}

async function testRevisionLifecycleAndStructuralText(api, runner) {
    const structuralText = 'Keep\tline\nsoft\u00adhard\u2011hyphen';
    const originalText = `Link clause.\n${structuralText}`;
    const modifiedText = `Link updated clause.\n${structuralText}`;
    const source = documentXml([[ 
        '<w:p w14:paraId="FACE1234">',
        '<w:hyperlink r:id="rId5" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<w:r><w:t>Link</w:t></w:r>',
        '</w:hyperlink>',
        '<w:r><w:t xml:space="preserve"> clause.</w:t></w:r>',
        '</w:p>'
    ].join(''), [
        '<w:p w14:paraId="FACE5678"><w:r><w:t>Keep</w:t><w:tab/><w:t>line</w:t>',
        '<w:br/><w:t>soft</w:t><w:softHyphen/><w:t>hard</w:t>',
        '<w:noBreakHyphen/><w:t>hyphen</w:t></w:r></w:p>'
    ].join('')]);
    assert.strictEqual(canonicalDocumentText(api, source), originalText);

    const changed = await runner.applyOperationToDocumentXml(source, {
        type: 'redline',
        target: { index: 1, exactText: 'Link clause.' },
        modified: 'Link updated clause.'
    }, 'AI Redliner', null, { strictTargets: true });
    assert.notStrictEqual(
        changed.status,
        'error',
        `structural-text redline failed: ${JSON.stringify(changed.error)}`
    );
    assert.strictEqual(canonicalDocumentText(api, changed.documentXml), modifiedText);
    assert.ok(changed.documentXml.includes('r:id="rId5"'), 'unchanged hyperlink relationship must survive');

    const revisionIds = Array.from(changed.documentXml.matchAll(
        /<w:(?:ins|del|pPrChange|rPrChange)\b[^>]*\bw:id="([^"]+)"/g
    ), match => match[1]);
    assert.ok(revisionIds.length > 0);
    assert.strictEqual(new Set(revisionIds).size, revisionIds.length);

    const accepted = api.acceptTrackedChangesInOoxml(changed.documentXml, { allAuthors: true });
    const rejected = api.rejectTrackedChangesInOoxml(changed.documentXml, { allAuthors: true });
    assert.notStrictEqual(accepted.status, 'error');
    assert.notStrictEqual(rejected.status, 'error');
    assert.strictEqual(canonicalDocumentText(api, accepted.oxml), modifiedText);
    assert.strictEqual(canonicalDocumentText(api, rejected.oxml), originalText);
}

async function testConsumerSanitizationPolicy(api, runner) {
    const source = documentXml([paragraphXml('Original clause.', '9999AAAA')]);
    const operation = {
        type: 'redline',
        target: { index: 1, exactText: 'Original clause.' },
        modified: 'Here is the redline:\nPay $1,000 under $term$.'
    };
    const prepared = prepareOperationInput(operation, true);
    assert.strictEqual(prepared.sanitized, true);
    assert.strictEqual(prepared.operation.modified, 'Pay $1,000 under $term$.');

    const sanitized = await runner.applyOperationToDocumentXml(
        source,
        prepared.operation,
        'AI Redliner',
        null,
        { generateRedlines: false, sanitizeInput: true, strictTargets: true }
    );
    const literal = await runner.applyOperationToDocumentXml(
        source,
        operation,
        'AI Redliner',
        null,
        { generateRedlines: false, sanitizeInput: false, strictTargets: true }
    );

    assert.strictEqual(sanitized.status, 'ok');
    assert.strictEqual(canonicalDocumentText(api, sanitized.documentXml), 'Pay $1,000 under $term$.');
    assert.strictEqual(literal.status, 'ok');
    assert.strictEqual(
        canonicalDocumentText(api, literal.documentXml),
        'Here is the redline:\nPay $1,000 under $term$.'
    );
}

async function run() {
    testNewCodesPassThroughConsumerBoundary();
    testConsumerBatchCallsRequireExplicitAtomicity();

    const packageRoot = resolvePackageRoot();
    const version = readPackageVersion(packageRoot);
    if (version !== EXPECTED_VERSION) {
        console.log(
            `PASS: v0.5.4 consumer guards; SKIP: package behavior requires ${EXPECTED_VERSION} `
            + `(resolved ${version}). Set DOCX_REDLINE_PACKAGE_ROOT to a v0.5.4 package root to run it before WP2.`
        );
        return;
    }

    const { api, runner } = await loadV054(packageRoot);
    testCanonicalTextAndFreshFingerprints(api);
    await testStructuredOperationFailures(runner);
    await testCommentedParagraphDeletionFailsClosed(runner);
    await testAtomicBatchRollsBackAllArtifacts(runner);
    await testSameAuthorMergeAndForeignAuthorRefusal(api, runner);
    await testRevisionLifecycleAndStructuralText(api, runner);
    await testConsumerSanitizationPolicy(api, runner);
    console.log('PASS: docx-redline-js v0.5.4 compatibility tests');
}

run().catch(error => {
    console.error('FAIL:', error?.stack || error);
    process.exit(1);
});
