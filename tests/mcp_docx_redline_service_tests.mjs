import assert from 'assert';
import { XMLSerializer } from '@xmldom/xmldom';
import {
    deriveParagraphAcceptedText,
    reconcileParagraphEdit
} from '../mcp/docx-server/src/services/docx-redline-js-service.mjs';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function paragraph(text) {
    return `<w:p xmlns:w="${NS}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

async function testNoOpRemainsNoOp() {
    const xml = paragraph('Unchanged');
    const result = await reconcileParagraphEdit({
        paragraphXml: xml,
        paragraphText: 'Unchanged',
        modifiedText: 'Unchanged',
        generateRedlines: false
    });

    assert.strictEqual(result.hasChanges, false);
    assert.deepStrictEqual(result.replacementNodes, []);
}

async function testStructuredErrorKeepsCode() {
    await assert.rejects(
        reconcileParagraphEdit({
            paragraphXml: paragraph('Actual source'),
            paragraphText: 'Missing source',
            modifiedText: 'Replacement',
            generateRedlines: false
        }),
        error => error.code === 'TARGET_NOT_FOUND'
            && error.message.includes('Original text was not found')
    );
}

async function testCallerContentIsNotImplicitlySanitized() {
    const result = await reconcileParagraphEdit({
        paragraphXml: paragraph('Original'),
        paragraphText: 'Original',
        modifiedText: 'Here is the agreement:\nPay $1,000 under $term$.',
        generateRedlines: false
    });

    assert.strictEqual(result.hasChanges, true);
    const acceptedText = result.replacementNodes
        .map(node => deriveParagraphAcceptedText(new XMLSerializer().serializeToString(node)))
        .join('\n');
    assert.ok(acceptedText.includes('Here is the agreement:'));
    assert.ok(acceptedText.includes('$1,000'));
    assert.ok(acceptedText.includes('$term$'));
}

async function run() {
    await testNoOpRemainsNoOp();
    await testStructuredErrorKeepsCode();
    await testCallerContentIsNotImplicitlySanitized();
    console.log('PASS: MCP docx-redline service tests');
}

run().catch(error => {
    console.error('FAIL:', error?.stack || error);
    process.exit(1);
});
