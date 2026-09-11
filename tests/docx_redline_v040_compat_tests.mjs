import './setup-xml-provider.mjs';

import assert from 'assert';
import {
    applyRedlineToOxml,
    ingestWordOoxmlToPlainTextResult
} from '@ansonlai/docx-redline-js';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function paragraph(text) {
    return `<w:p xmlns:w="${NS}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

async function testStructuredTargetError() {
    const result = await applyRedlineToOxml(
        paragraph('Actual source'),
        'Missing source',
        'Updated source'
    );

    assert.strictEqual(result.status, 'error');
    assert.strictEqual(result.error?.code, 'TARGET_NOT_FOUND');
    assert.strictEqual(result.hasChanges, false);
}

async function testSanitizationIsExplicit() {
    const original = 'Original clause.';
    const proposed = 'Here is the redline:\nUpdated clause.';
    const unsanitized = await applyRedlineToOxml(
        paragraph(original),
        original,
        proposed,
        { generateRedlines: false, sanitizeInput: false }
    );
    const sanitized = await applyRedlineToOxml(
        paragraph(original),
        original,
        proposed,
        { generateRedlines: false, sanitizeInput: true }
    );

    assert.ok(unsanitized.oxml.includes('Here is the redline:'));
    assert.ok(!sanitized.oxml.includes('Here is the redline:'));
    assert.strictEqual(ingestWordOoxmlToPlainTextResult(sanitized.oxml).text, 'Updated clause.');
    assert.ok(sanitized.warnings?.some(warning => warning.includes('sanitized')));
}

async function testLiteralDollarAndWhitespaceContentSurvives() {
    const original = 'Fee due.';
    const proposed = '  Pay $1,000 under $term$.  ';
    const result = await applyRedlineToOxml(
        paragraph(original),
        original,
        proposed,
        { generateRedlines: false, sanitizeInput: true }
    );

    assert.notStrictEqual(result.status, 'error');
    assert.ok(result.oxml.includes(proposed));
    assert.ok(result.oxml.includes('xml:space="preserve"'));
}

async function run() {
    await testStructuredTargetError();
    await testSanitizationIsExplicit();
    await testLiteralDollarAndWhitespaceContentSurvives();
    console.log('PASS: docx-redline-js v0.4.0 compatibility tests');
}

run().catch(error => {
    console.error('FAIL:', error?.stack || error);
    process.exit(1);
});
