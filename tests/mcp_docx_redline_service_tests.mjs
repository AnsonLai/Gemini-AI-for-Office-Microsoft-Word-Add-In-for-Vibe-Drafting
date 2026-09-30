import assert from 'node:assert/strict';
import {
    applyDocumentOperations,
    createDocument,
    editDocumentParagraph,
    listDocumentParagraphs
} from '../mcp/docx-server/src/services/docx-document-service.mjs';

async function testNoOpRemainsNoOp() {
    const doc = await createDocument('Unchanged');
    const handle = listDocumentParagraphs(doc).items[0].id;
    const result = await editDocumentParagraph(doc, handle, 'Unchanged', { generateRedlines: false });
    assert.equal(result.changed, false);
    assert.equal(result.updatedText, 'Unchanged');
}

async function testStructuredErrorKeepsCode() {
    const doc = await createDocument('Actual source');
    await assert.rejects(
        applyDocumentOperations(doc, [{
            type: 'replace', target: { exactText: 'Missing source' }, modified: 'Replacement'
        }]),
        error => error.code === 'TARGET_NOT_FOUND' && error.details.rolledBack === true
    );
    assert.equal(doc.inspect().paragraphs[0].text, 'Actual source');
}

async function testCallerContentIsNotImplicitlySanitized() {
    const doc = await createDocument('Original');
    const handle = listDocumentParagraphs(doc).items[0].id;
    const result = await editDocumentParagraph(doc, handle, 'Here is the agreement:\nPay $1,000 under $term$.', { generateRedlines: false });
    assert.equal(result.changed, true);
    const accepted = doc.inspect().paragraphs.map(item => item.text).join('\n');
    assert.ok(accepted.includes('Here is the agreement:'));
    assert.ok(accepted.includes('$1,000'));
    assert.ok(accepted.includes('$term$'));
}

async function run() {
    await testNoOpRemainsNoOp();
    await testStructuredErrorKeepsCode();
    await testCallerContentIsNotImplicitlySanitized();
    console.log('PASS: MCP docx document service tests');
}

run().catch(error => {
    console.error('FAIL:', error?.stack || error);
    process.exit(1);
});
