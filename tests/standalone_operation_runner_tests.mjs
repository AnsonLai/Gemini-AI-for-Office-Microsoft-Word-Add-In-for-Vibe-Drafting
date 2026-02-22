import './setup-xml-provider.mjs';

import assert from 'assert';
import { buildTargetReferenceSnapshot, getParagraphText } from '../src/taskpane/modules/reconciliation/standalone.js';
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

function buildTwoColumnTitleTableDocumentXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Title:</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Title:</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
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

async function testSingleParagraphListConcatenationWithInlineMarkersDoesNotInsertExtraItem() {
    const existingItems = [
        'The Disclosing Party possesses certain confidential, proprietary, and trade secret information.',
        'The Parties desire to enter into a potential business relationship or transaction (the “Purpose”), which requires the Disclosing Party to disclose certain Confidential Information (as defined below) to the Receiving Party.',
        'The Receiving Party agrees to receive and treat such Confidential Information in confidence, subject to the terms and conditions of this Agreement.'
    ];
    const inputXml = buildNumberedListDocumentXml(existingItems);
    const malformedMergedEdit = [
        'A. The Disclosing Party possesses certain confidential, proprietary, and trade secret information.',
        'B. The Parties desire to enter into a potential business relationship or transaction (the “Purpose”), which requires the Disclosing Party to disclose certain Confidential Information (as defined below) to the Receiving Party.',
        'C. The Receiving Party agrees to receive and treat such Confidential Information in confidence, subject to the terms and conditions of this Agreement.',
        existingItems[0]
    ].join(' ');

    const logs = [];
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: existingItems[0],
            targetRef: 'P1',
            modified: malformedMergedEdit
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            onInfo: message => logs.push(String(message)),
            onWarn: message => logs.push(String(message))
        }
    );

    assert.strictEqual(result.hasChanges, true, 'malformed single-paragraph list edit should still report changes');
    assert.strictEqual(
        logs.some(message => message.includes('single-paragraph list adjacency insertion heuristic')),
        false,
        'malformed inline list markers should not trigger adjacency insertion heuristic'
    );

    const resultDoc = parseXmlStrict(result.documentXml, 'single paragraph malformed list edit output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const paragraphTexts = paragraphs.map(paragraph => getParagraphText(paragraph)).filter(Boolean);
    assert.strictEqual(
        paragraphTexts.length,
        3,
        'malformed single-paragraph list edit should not insert an extra list paragraph'
    );
}

async function testPlainParagraphInsertionBeforeTargetCreatesSeparateParagraph() {
    const original = 'NON-DISCLOSURE AGREEMENT';
    const insertedMarkdown = '**INSTRUCTIONS:** Please review this Non-Disclosure Agreement carefully.';
    const insertedPlain = 'INSTRUCTIONS: Please review this Non-Disclosure Agreement carefully.';
    const logs = [];
    const inputXml = buildDocumentXml(original);
    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: original,
            targetRef: 'P1',
            modified: `${insertedMarkdown}\n${original}`
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true,
            onInfo: message => logs.push(String(message)),
            onWarn: message => logs.push(String(message))
        }
    );

    assert.strictEqual(result.hasChanges, true, 'plain insertion-before-target shape should report changes');
    assert.strictEqual(
        logs.some(message => message.includes('plain adjacency insertion heuristic')),
        true,
        'plain insertion-before-target shape should use plain adjacency insertion heuristic'
    );

    const resultDoc = parseXmlStrict(result.documentXml, 'plain insertion-before-target output');
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const paragraphTexts = paragraphs.map(paragraph => getParagraphText(paragraph).trim()).filter(Boolean);

    assert.strictEqual(
        paragraphTexts.length,
        2,
        'plain insertion-before-target should produce two paragraphs'
    );
    assert.strictEqual(
        paragraphTexts[0],
        insertedPlain,
        'plain insertion-before-target should insert new paragraph text before original target paragraph'
    );
    assert.strictEqual(
        paragraphTexts[1],
        original,
        'plain insertion-before-target should preserve original target paragraph as separate paragraph'
    );
    assert.ok(
        result.documentXml.includes('<w:b'),
        'plain insertion-before-target should parse markdown formatting via existing formatter pipeline'
    );
    assert.ok(
        !paragraphTexts[0].includes('**'),
        'plain insertion-before-target should not leave raw markdown markers in output text'
    );
}

async function testFormatOnlyRedlineWithTrackedWrapperStillApplies() {
    const originalText = 'These instructions are for the user to fill out the document. Please replace all bracketed information (e.g., "[Name of Disclosing Party]") with the appropriate details. Ensure all necessary signatures are obtained. NON-DISCLOSURE AGREEMENT';
    const modifiedText = 'These instructions are for the user to fill out the document. Please replace all bracketed information (e.g., "[Name of Disclosing Party]") with the appropriate details. Ensure all necessary signatures are obtained. ++NON-DISCLOSURE AGREEMENT++';
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:ins w:id="1" w:author="Prior" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>These instructions are for the user to fill out the document. Please replace all bracketed information (e.g., "[Name of Disclosing Party]") with the appropriate details. Ensure all necessary signatures are obtained. </w:t></w:r>
      </w:ins>
      <w:r><w:t>NON-DISCLOSURE AGREEMENT</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            target: originalText,
            targetRef: 'P1',
            modified: modifiedText
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should still apply when paragraph contains tracked-change wrappers'
    );
    assert.ok(
        result.documentXml.includes('<w:u'),
        'format-only redline should apply underline markup to the target run'
    );
}

async function testTextToTableWithoutHeaderSeparatorPreservesAllRows() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>Disclosing Party: [Name of Disclosing Party]</w:t></w:r></w:p>
    <w:p><w:r><w:t>And</w:t></w:r></w:p>
    <w:p><w:r><w:t>Receiving Party: [Name of Receiving Party]</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const modifiedText = [
        '| Disclosing Party: | [Name of Disclosing Party] [Address of Disclosing Party] (the "Disclosing Party") |',
        '| Receiving Party: | [Name of Receiving Party] [Address of Receiving Party] (the "Receiving Party") |'
    ].join('\n');

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            targetEndRef: 'P3',
            target: 'Disclosing Party: [Name of Disclosing Party]',
            modified: modifiedText
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(result.hasChanges, true, 'text-to-table redline should report changes');
    const resultDoc = parseXmlStrict(result.documentXml, 'text-to-table no-header output');
    const tables = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'tbl'));
    assert.strictEqual(tables.length, 1, 'text-to-table redline should produce one table');
    const tableText = Array.from(tables[0].getElementsByTagNameNS(NS_W, 't'))
        .map(node => String(node.textContent || ''))
        .join(' ');
    assert.ok(
        tableText.includes('Disclosing Party:'),
        'table should include first markdown row text when no header separator is provided'
    );
    assert.ok(
        tableText.includes('Receiving Party:'),
        'table should include second markdown row text'
    );
}

async function testFormatOnlyRedlineSupportsNonWPrefixOoxml() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:document xmlns:x="${NS_W}">
  <x:body>
    <x:p><x:r><x:t>By</x:t></x:r></x:p>
    <x:sectPr/>
  </x:body>
</x:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'By',
            modified: '**By**'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should apply even when OOXML uses a non-w namespace prefix'
    );
    const resultDoc = parseXmlStrict(result.documentXml, 'non-w-prefix format-only output');
    assert.ok(
        resultDoc.getElementsByTagNameNS(NS_W, 'b').length > 0,
        'format-only redline should emit bold run properties on non-w-prefix OOXML'
    );
}

async function testFormatOnlyRedlineWithAllTextInsideInsertionWrapper() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:ins w:id="9" w:author="Prior" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>By: [Name]</w:t></w:r>
      </w:ins>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'By: [Name]',
            modified: '**By**: [Name]'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should apply when all text is nested inside insertion wrapper'
    );
    const resultDoc = parseXmlStrict(result.documentXml, 'format-only insertion wrapper output');
    assert.ok(
        resultDoc.getElementsByTagNameNS(NS_W, 'b').length > 0,
        'format-only redline should emit bold formatting even when text is wrapped in w:ins'
    );
}

async function testFormatOnlyRedlineRematchesWhenRefParagraphDrifts() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t></w:t></w:r></w:p>
    <w:p><w:r><w:t>By: [Name]</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'By: [Name]',
            modified: '**By**: [Name]'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should rematch by text when targetRef paragraph has drifted'
    );
    const resultDoc = parseXmlStrict(result.documentXml, 'ref drift rematch output');
    assert.ok(
        resultDoc.getElementsByTagNameNS(NS_W, 'b').length > 0,
        'format-only redline should apply bold formatting after ref-drift rematch'
    );
}

async function testFormatOnlyRedlineFallsBackToOoxmlWhenNoSpansAreExtractable() {
    const inputXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> MERGEFIELD  SignatureLine </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(
        inputXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: 'By: [Name]',
            modified: '**By:** [Name]'
        },
        'StandaloneRunnerTest',
        null,
        {
            generateRedlines: true
        }
    );

    assert.strictEqual(
        result.hasChanges,
        true,
        'format-only redline should not no-op when no text spans are extractable from OOXML'
    );
    const resultDoc = parseXmlStrict(result.documentXml, 'no-span format-only output');
    assert.ok(
        resultDoc.getElementsByTagNameNS(NS_W, 'b').length > 0,
        'format-only redline should still emit bold formatting via OOXML fallback'
    );
    const paragraphs = Array.from(resultDoc.getElementsByTagNameNS(NS_W, 'p'));
    const outputText = paragraphs.map(paragraph => getParagraphText(paragraph)).join('\n');
    assert.ok(
        outputText.includes('By: [Name]'),
        'format-only OOXML fallback should preserve rendered target text'
    );
}

async function testDuplicateTableStructuralOpsAreDedupedPerTurn() {
    const inputXml = buildTwoColumnTitleTableDocumentXml();
    const snapshotDoc = parseXmlStrict(inputXml, 'table dedupe snapshot');
    const runtimeContext = {
        targetRefSnapshot: buildTargetReferenceSnapshot(snapshotDoc),
        tableStructuralRedlineKeys: new Set()
    };

    const opA = {
        type: 'redline',
        targetRef: 'P1',
        target: 'Title:',
        modified: 'Title:\nDate:'
    };
    const opB = {
        type: 'redline',
        targetRef: 'P2',
        target: 'Title:',
        modified: 'Title:\nDate:'
    };

    const stepA = await applyOperationToDocumentXml(
        inputXml,
        opA,
        'StandaloneRunnerTest',
        runtimeContext,
        {
            generateRedlines: true
        }
    );
    assert.strictEqual(stepA.hasChanges, true, 'first table-structural op should apply');

    const stepB = await applyOperationToDocumentXml(
        stepA.documentXml,
        opB,
        'StandaloneRunnerTest',
        runtimeContext,
        {
            generateRedlines: true
        }
    );
    assert.strictEqual(
        stepB.hasChanges,
        false,
        'duplicate table-structural op in the same turn should be skipped'
    );
    assert.strictEqual(
        (stepB.warnings || []).some(w => String(w).includes('duplicate table-structural redline')),
        true,
        'dedupe skip should emit a warning'
    );
}

async function run() {
    await testRedlineOperation();
    await testCommentOperation();
    await testRangeListRedlineDoesNotDuplicateExistingItems();
    await testSingleParagraphListConcatenationUsesSurgicalInsertion();
    await testSingleParagraphListConcatenationWithInlineMarkersDoesNotInsertExtraItem();
    await testPlainParagraphInsertionBeforeTargetCreatesSeparateParagraph();
    await testFormatOnlyRedlineWithTrackedWrapperStillApplies();
    await testTextToTableWithoutHeaderSeparatorPreservesAllRows();
    await testFormatOnlyRedlineSupportsNonWPrefixOoxml();
    await testFormatOnlyRedlineWithAllTextInsideInsertionWrapper();
    await testFormatOnlyRedlineRematchesWhenRefParagraphDrifts();
    await testFormatOnlyRedlineFallsBackToOoxmlWhenNoSpansAreExtractable();
    await testDuplicateTableStructuralOpsAreDedupedPerTurn();
    console.log('PASS: standalone operation runner tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
