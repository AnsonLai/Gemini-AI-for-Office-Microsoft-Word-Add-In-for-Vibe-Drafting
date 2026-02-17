import './setup-xml-provider.mjs';

import assert from 'assert';
import { getParagraphText } from '../src/taskpane/modules/reconciliation/standalone.js';
import { applyOperationToDocumentXml } from '../src/taskpane/modules/reconciliation/services/standalone-operation-runner.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parseXmlStrict(xmlText, label) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown'}`);
    }
    return xmlDoc;
}

function buildDocumentXml(text) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function buildNumberedListDocumentXml(items, numId = '77') {
    const paragraphs = items
        .map(item => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${item}</w:t></w:r></w:p>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${paragraphs}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

async function testRedlineOperation() {
    const sourceText = 'Alpha target text.';
    const modifiedText = 'Alpha target text updated.';
    const inputXml = buildDocumentXml(sourceText);
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: sourceText,
            modified: modifiedText
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: false
        }
    );

    assert.strictEqual(result.hasChanges, true, 'redline operation should report changes');
    const resultDoc = parseXmlStrict(result.documentXml, 'redline output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const firstParagraphText = getParagraphText(paragraphs[0]).trim();
    assert.strictEqual(firstParagraphText, modifiedText, 'redline operation should rewrite paragraph text');
}

async function testCommentOperation() {
    const sourceText = 'Comment target paragraph.';
    const inputXml = buildDocumentXml(sourceText);
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'comment',
            target: sourceText,
            textToComment: 'target',
            commentContent: 'Please review this term.'
        },
        'StandaloneRunnerTest'
    );

    assert.strictEqual(result.hasChanges, true, 'comment operation should report changes');
    assert.ok(result.commentsXml && result.commentsXml.includes('Please review this term.'), 'comment operation should emit comments xml');
}

async function testRangeListRedlineDoesNotDuplicateExistingItems() {
    const existingItems = [
        'Business plans, strategies, financial information, pricing, and marketing data.',
        'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.',
        'Information concerning the Disclosing Party\'s employees, contractors, customers, and suppliers.',
        'Any notes, analyses, compilations, studies, or other materials prepared by the Receiving Party that contain, reflect, or are derived from the foregoing.'
    ];
    const insertedItem = 'Photographs, videos, and other recordings of prototypes and physical hardware.';
    const modifiedText = [
        `1. ${existingItems[0]}`,
        `2. ${insertedItem}`,
        `3. ${existingItems[1]}`,
        `4. ${existingItems[2]}`,
        `5. ${existingItems[3]}`
    ].join('\n');
    const inputXml = buildNumberedListDocumentXml(existingItems);
    const logs = [];
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: existingItems[0],
            targetRef: 'P1',
            targetEndRef: 'P4',
            modified: modifiedText
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            onInfo: message => logs.push(String(message)),
            onWarn: message => logs.push(String(message))
        }
    );

    assert.strictEqual(result.hasChanges, true, 'range list redline should report changes');
    assert.strictEqual(
        logs.some(message => message.includes('Applying explicit-range insertion-only heuristic')),
        true,
        'range list redline should use explicit-range insertion-only heuristic'
    );

    const resultDoc = parseXmlStrict(result.documentXml, 'range list redline output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const revisionDeletes = resultDoc.getElementsByTagNameNS(NS_W, 'del').length;
    const revisionInserts = resultDoc.getElementsByTagNameNS(NS_W, 'ins').length;
    const paragraphTexts = paragraphs.map(paragraph => getParagraphText(paragraph)).filter(Boolean);
    assert.strictEqual(
        paragraphTexts.length,
        5,
        'range list redline should produce five list paragraphs without duplicate tail items'
    );
    assert.strictEqual(
        paragraphTexts.filter(text => text === existingItems[1]).length,
        1,
        'existing item #2 should not be duplicated'
    );
    assert.strictEqual(
        paragraphTexts[1],
        insertedItem,
        'new confidentiality bullet should be inserted at position #2'
    );
    assert.strictEqual(
        revisionDeletes,
        0,
        'surgical insertion should not emit delete revisions for untouched list items'
    );
    assert.strictEqual(
        revisionInserts,
        1,
        'surgical insertion should emit exactly one inserted revision for the new list item'
    );

    const listNumIds = paragraphs.map(paragraph => {
        const numPr = paragraph.getElementsByTagNameNS(NS_W, 'numPr')[0];
        const numIdNode = numPr ? numPr.getElementsByTagNameNS(NS_W, 'numId')[0] : null;
        return numIdNode ? (numIdNode.getAttribute('w:val') || numIdNode.getAttribute('val') || null) : null;
    });
    assert.strictEqual(
        listNumIds.filter(numId => numId === '77').length,
        5,
        'surgical insertion should preserve original list numId across all items'
    );
}

async function testSingleParagraphListConcatenationUsesSurgicalInsertion() {
    const existingItems = [
        'Business plans, strategies, financial information, pricing, and marketing data.',
        'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.',
        'Information concerning the Disclosing Party\'s employees, contractors, customers, and suppliers.',
        'Any notes, analyses, compilations, studies, or other materials prepared by the Receiving Party that contain, reflect, or are derived from the foregoing.'
    ];
    const insertedItem = 'Photographs, videos, and other recordings of prototypes and physical hardware.';
    const inputXml = buildNumberedListDocumentXml(existingItems);
    const logs = [];
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: existingItems[1],
            targetRef: 'P2',
            modified: `${insertedItem}${existingItems[1]}`
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            onInfo: message => logs.push(String(message)),
            onWarn: message => logs.push(String(message))
        }
    );

    assert.strictEqual(result.hasChanges, true, 'single-paragraph concatenation edit should report changes');
    assert.strictEqual(
        logs.some(message => message.includes('single-paragraph list adjacency insertion heuristic')),
        true,
        'single-paragraph concatenation edit should use adjacency insertion heuristic'
    );

    const resultDoc = parseXmlStrict(result.documentXml, 'single paragraph list insertion output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const revisionDeletes = resultDoc.getElementsByTagNameNS(NS_W, 'del').length;
    const revisionInserts = resultDoc.getElementsByTagNameNS(NS_W, 'ins').length;
    const paragraphTexts = paragraphs.map(paragraph => getParagraphText(paragraph)).filter(Boolean);

    assert.strictEqual(paragraphTexts.length, 5, 'single-paragraph concatenation should become one inserted list item');
    assert.strictEqual(paragraphTexts[1], insertedItem, 'new item should be inserted directly before original target item');
    assert.strictEqual(
        paragraphTexts.filter(text => text === existingItems[1]).length,
        1,
        'original target item should remain a single untouched list item'
    );
    assert.strictEqual(revisionDeletes, 0, 'adjacency insertion should not emit delete revisions');
    assert.strictEqual(revisionInserts, 1, 'adjacency insertion should emit exactly one inserted revision');
}

async function run() {
    await testRedlineOperation();
    await testCommentOperation();
    await testRangeListRedlineDoesNotDuplicateExistingItems();
    await testSingleParagraphListConcatenationUsesSurgicalInsertion();
    console.log('PASS: standalone operation runner tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
