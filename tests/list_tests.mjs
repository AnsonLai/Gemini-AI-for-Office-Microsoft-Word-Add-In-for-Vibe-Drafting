import './setup-xml-provider.mjs';

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { ReconciliationPipeline } from '../src/taskpane/modules/reconciliation/pipeline/pipeline.js';
import { NumberingService } from '../src/taskpane/modules/reconciliation/services/numbering-service.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/pipeline/ingestion.js';
import { serializeToOoxml } from '../src/taskpane/modules/reconciliation/pipeline/serialization.js';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/engine/oxml-engine.js';
import {
    resolveSingleLineListFallbackNumberingAction,
    recordSingleLineListFallbackExplicitSequence,
    clearSingleLineListFallbackExplicitSequence,
    enforceListBindingOnParagraphNodes
} from '../src/taskpane/modules/reconciliation/standalone.js';
import {
    planListInsertionOnlyEdit,
    stripRedundantLeadingListMarkers
} from '../src/taskpane/modules/reconciliation/core/list-targeting.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// --- Test 1: List Expansion with Soft Break (from repro_list_issue.js) ---
async function testListExpansionWithSoftBreak() {
    console.log('\n=== Test: List Expansion with Soft Break ===');
    const pipeline = new ReconciliationPipeline({
        generateRedlines: true,
        author: 'AI',
        numberingService: new NumberingService()
    });

    // Original paragraph with a soft break (<w:br/>)
    const originalOoxml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Phase 1 text.</w:t><w:br/><w:t>More text.</w:t></w:r></w:p>';
    const modifiedText = 'A. Point 1\nB. Point 2\nC. Point 3';

    try {
        const result = await pipeline.execute(originalOoxml, modifiedText);
        const pTags = (result.ooxml.match(/<w:p\b/g) || []);
        const pCount = pTags.length;

        // Allow 3 or 4 (3 list items + optional trailing empty paragraph)
        if (pCount === 3 || pCount === 4) {
            console.log('✅ SUCCESS: Correctly expanded into 3 paragraphs (plus optional trailing).');
        } else {
            console.error('❌ FAILED: Expected 3-4 paragraphs, got ' + pCount);
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Test 2: Surgical List (from test-list-surgical.mjs) ---
async function testSurgicalList() {
    console.log('\n=== Test: Surgical List Operations ===');

    // Sub-test 1: Multi-paragraph ingestion
    console.log('--- Sub-test: Multi-Paragraph Ingestion ---');
    const testOoxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>Item One</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>Item Two</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>Item Three</w:t></w:r></w:p>
  </w:body>
</w:document>
`;

    const { runModel, acceptedText, pPr } = ingestOoxml(testOoxml);

    // Sub-test 2: Round-trip serialization
    console.log('--- Sub-test: Round-Trip Serialization ---');
    const serialized = serializeToOoxml(runModel, pPr, []);
    const pCount = (serialized.match(/<w:p>/g) || []).length;
    if (pCount === 3) {
        console.log('✅ PASS: Serialization maintained paragraph count.');
    } else {
        console.log('❌ FAIL: Serialization lost paragraphs.');
    }

    // Sub-test 3: Surgical diff on existing list
    console.log('--- Sub-test: Surgical Diff ---');
    const pipeline = new ReconciliationPipeline({ author: 'Test', generateRedlines: true });
    // Modify only the third item
    const modifiedText = 'Item One\nItem Two\nItem Three MODIFIED';

    try {
        const result = await pipeline.execute(testOoxml, modifiedText);

        // Check that deletion of "Item One" is NOT present
        const hasUnwantedDeletion = result.ooxml.includes('<w:delText xml:space="preserve">Item One</w:delText>');
        // Check that insertion of "MODIFIED" IS present
        const hasExpectedInsertion = result.ooxml.includes('MODIFIED');

        if (!hasUnwantedDeletion && hasExpectedInsertion) {
            console.log('✅ SUCCESS: Surgical diff is working!');
        } else {
            console.log('❌ FAILURE: Still doing full replacement or missing edit');
        }
    } catch (e) {
        console.error('Pipeline error:', e);
    }
}

// --- Test 3: Repro List Issue (from repro_list_issue.mjs) ---
async function testReproListIssue() {
    console.log('\n=== Test: Repro List Issue (Numbering Reset) ===');
    const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

    try {
        let originalOoxml;
        try {
            originalOoxml = await fs.promises.readFile(DOC_PATH, 'utf-8');
        } catch (err) {
            console.log(`⚠️ SKIPPING: Could not read sample_doc at ${DOC_PATH}. This is expected if sample_doc is missing.`);
            return;
        }

        const { acceptedText: originalText } = ingestOoxml(originalOoxml);
        const targetSentence = "This copy is to be used solely for archival purposes to ensure compliance and record-keeping.";

        if (!originalText.includes(targetSentence)) {
            console.log("⚠️ SKIPPING: Target sentence not found in sample_doc.");
            return;
        }

        const insertion = "\n    * The retention of such copy must be legally required.";
        const newText = originalText.replace(targetSentence, targetSentence + insertion);

        const result = await applyRedlineToOxml(originalOoxml, originalText, newText, {
            author: 'TestUser',
            generateRedlines: true
        });

        const index = result.oxml.indexOf('must be legally required');
        if (index === -1) {
            console.log("❌ FAIL: Inserted text not found in output.");
        } else {
            console.log("✅ Inserted text found.");
        }

    } catch (e) {
        console.error("💥 Error during test:", e);
    }
}

// --- Test 4: Fixed List Conversion (from verify_fixed_list_conversion.mjs) ---
// Helper for this test
function parseMarkdownList(content) {
    if (!content) return null;
    const lines = content.trim().split('\n');
    const items = [];
    for (const line of lines) {
        if (!line.trim()) continue;
        const markerRegex = /^(\s*)((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*•])\s*)(.*)$/;
        const match = line.match(markerRegex);
        if (match) {
            const indent = match[1];
            const marker = match[2].trim();
            const text = match[3];
            const level = Math.floor(indent.length / 2);
            const isBullet = /^[-*•]$/.test(marker);
            items.push({
                type: isBullet ? 'bullet' : 'numbered',
                level,
                text: text.trim(),
                marker: marker
            });
            continue;
        }
        items.push({ type: 'text', level: 0, text: line.trim() });
    }
    if (items.length === 0) return null;
    const hasNumbered = items.some(i => i.type === 'numbered');
    const hasBullet = items.some(i => i.type === 'bullet');
    return {
        type: hasNumbered ? 'numbered' : (hasBullet ? 'bullet' : 'text'),
        items: items
    };
}

async function verifyFixedListConversion() {
    console.log('\n=== Test: Fixed List Conversion ===');

    // Test detection
    console.log('--- Sub-test: parseMarkdownList Detection ---');
    const content = "A. Item 1\nB. Item 2";
    const listData = parseMarkdownList(content);
    if (listData && listData.type === 'numbered') {
        console.log('✅ PASS: Alpha markers detected.');
    } else {
        console.log('❌ FAIL: Alpha markers missed.');
    }

    // Test conversion
    console.log('--- Sub-test: Alpha List Conversion ---');
    const pipeline = new ReconciliationPipeline({ generateRedlines: false });
    const contentAlpha = "A. Item 1\nB. Item 2\nC. Item 3";

    // Mock context
    try {
        const result = await pipeline.executeListGeneration(contentAlpha, null, null, "Original Text");
        const hasUpperLetter = result.numberingXml && result.numberingXml.includes('w:numFmt w:val="upperLetter"');
        if (hasUpperLetter) {
            console.log('✅ PASS: Alpha markers detected and numbering.xml updated.');
        } else {
            console.log('❌ FAIL: Alpha markers not properly mapped to numbering.xml.');
        }
    } catch (e) {
        console.log('⚠️ ERROR in Alpha List Conversion:', e.message);
    }

    // Test Font Inheritance
    console.log('--- Sub-test: Font Inheritance ---');
    const pipelineFont = new ReconciliationPipeline({
        generateRedlines: false,
        font: 'Calibri'
    });
    try {
        const result = await pipelineFont.executeListGeneration("1. Item 1", null, null, "Original Text");
        const hasFont = result.ooxml.includes('w:rFonts w:ascii="Calibri"');
        if (hasFont) {
            console.log('✅ PASS: Font inherited correctly.');
        } else {
            console.log('❌ FAIL: Font not found in output OOXML.');
        }
    } catch (e) {
        console.log('❌ FAIL: Execution error:', e.message);
    }
}

// --- Test 5: Verify List Fixes (from verify_list_fixes.mjs) ---
async function verifyListFixes() {
    console.log('\n=== Test: Verify List Fixes ===');

    // Sub-test 1: List Context Preservation
    console.log('--- Sub-test: Single Item Context Preservation ---');
    const originalOoxml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
            <w:numPr>
                <w:ilvl w:val="0"/>
                <w:numId w:val="1"/>
            </w:numPr>
        </w:pPr>
        <w:r><w:t>Verification of Compliance</w:t></w:r>
    </w:p>`.trim();

    const originalText = "Verification of Compliance";
    const newContent = "The retention of this copy must be legally required.";

    const result = await applyRedlineToOxml(originalOoxml, originalText, newContent, {
        author: 'TestUser',
        generateRedlines: true
    });

    const hasNumPr = result.oxml.includes('<w:numId w:val="1"/>') && result.oxml.includes('<w:ilvl w:val="0"/>');

    if (hasNumPr) {
        console.log('✅ PASS: Context preserved.');
    } else {
        console.log('❌ FAIL: Context lost.');
    }

    // Sub-test 2: Nested Item Insertion
    console.log('--- Sub-test: Nested Item Generation (1.1.1) ---');
    const originalOoxmlNested = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
            <w:numPr>
                <w:ilvl w:val="1"/>
                <w:numId w:val="1"/>
            </w:numPr>
        </w:pPr>
        <w:r><w:t>Item 1.1</w:t></w:r>
    </w:p>`.trim();

    const originalTextNested = "Item 1.1";
    const newContentNested = "Item 1.1\n  1.1.1. New nested item";

    const resultNested = await applyRedlineToOxml(originalOoxmlNested, originalTextNested, newContentNested, {
        author: 'TestUser',
        generateRedlines: false
    });

    const paragraphs = resultNested.oxml.match(/<w:p[\s\S]*?<\/w:p>/g);

    if (paragraphs && paragraphs.length >= 2) {
        const p2 = paragraphs[1];
        const hasIlvl2 = p2.includes('w:ilvl w:val="2"');
        if (hasIlvl2) {
            console.log('✅ PASS: Correct ilvl generated for nested marker.');
        } else {
            console.log('❌ FAIL: Incorrect ilvl for nested marker.');
        }
    } else {
        console.log('❌ FAIL: Second paragraph not generated.');
    }
}

// --- Test 5b: Insertion-only nested depth for sub-sub ordered intent ---
function testInsertionOnlyNestedDepthHeuristics() {
    console.log('\n=== Test: Insertion-Only Nested Depth Heuristics ===');

    const parser = new DOMParser();
    const paragraphDoc = parser.parseFromString(`
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
            <w:p>
                <w:pPr>
                    <w:numPr>
                        <w:ilvl w:val="1"/>
                        <w:numId w:val="42"/>
                    </w:numPr>
                </w:pPr>
                <w:r><w:t>This copy is to be used solely for archival purposes to ensure compliance and record-keeping.</w:t></w:r>
            </w:p>
        </w:body>
    </w:document>`, 'application/xml');
    const targetParagraph = paragraphDoc.getElementsByTagNameNS('*', 'p')[0];
    const anchorText = 'This copy is to be used solely for archival purposes to ensure compliance and record-keeping.';

    const bulletAmbiguous = `${anchorText}\n  - This archival copy must be legally required by the SEC or FCC.`;
    const bulletPlan = planListInsertionOnlyEdit(targetParagraph, bulletAmbiguous, {
        currentParagraphText: anchorText
    });
    const bulletIlvl = bulletPlan?.entries?.[0]?.ilvl ?? null;
    if (bulletPlan && bulletIlvl === 2) {
        console.log('✅ PASS: Ambiguous bullet insertion promoted to child depth (ilvl=2).');
    } else {
        console.log('❌ FAIL: Ambiguous bullet insertion did not promote to child depth.', { bulletIlvl });
    }

    const explicitComposite = `${anchorText}\n  2.2.1. This archival copy must be legally required by the SEC or FCC.`;
    const explicitPlan = planListInsertionOnlyEdit(targetParagraph, explicitComposite, {
        currentParagraphText: anchorText
    });
    const explicitIlvl = explicitPlan?.entries?.[0]?.ilvl ?? null;
    if (explicitPlan && explicitIlvl === 2) {
        console.log('✅ PASS: Explicit composite marker mapped to ilvl=2.');
    } else {
        console.log('❌ FAIL: Explicit composite marker depth mapping incorrect.', { explicitIlvl });
    }
}

// --- Test 5c: Redundant manual marker stripping for list item text ---
function testRedundantLeadingListMarkerStripping() {
    console.log('\n=== Test: Redundant Leading Marker Stripping ===');

    const cases = [
        { input: '- The Receiving Party retains one copy.', expected: 'The Receiving Party retains one copy.' },
        { input: '2.1. The Receiving Party retains one copy.', expected: 'The Receiving Party retains one copy.' },
        { input: '2.1. - The Receiving Party retains one copy.', expected: 'The Receiving Party retains one copy.' },
        { input: 'Specifically, retention must be legally required by the SEC or FCC.', expected: 'Specifically, retention must be legally required by the SEC or FCC.' }
    ];

    let failed = 0;
    for (const testCase of cases) {
        const actual = stripRedundantLeadingListMarkers(testCase.input);
        if (actual !== testCase.expected) {
            failed++;
            console.log('❌ FAIL:', { input: testCase.input, expected: testCase.expected, actual });
        }
    }

    if (failed === 0) {
        console.log('✅ PASS: Redundant list marker prefixes are stripped as expected.');
    }
}

// --- Test 5d: Explicit single-line fallback sequence action helpers ---
function testExplicitSingleLineFallbackSequenceHelpers() {
    console.log('\n=== Test: Explicit Single-Line Fallback Sequence Helpers ===');

    const state = { explicitByNumberingKey: new Map() };
    const plan1 = { numberingKey: 'numbered:decimal:single', startAt: 1 };
    const action1 = resolveSingleLineListFallbackNumberingAction(plan1, state);
    if (action1.type !== 'explicitStartNew') {
        console.log('❌ FAIL: Expected explicitStartNew for first marker.', action1);
        return;
    }

    recordSingleLineListFallbackExplicitSequence(state, action1.numberingKey, '1000', 1);
    const plan2 = { numberingKey: 'numbered:decimal:single', startAt: 2 };
    const action2 = resolveSingleLineListFallbackNumberingAction(plan2, state);
    if (!(action2.type === 'explicitReuse' && action2.numId === '1000')) {
        console.log('❌ FAIL: Expected explicitReuse with numId 1000 for second marker.', action2);
        return;
    }

    recordSingleLineListFallbackExplicitSequence(state, action2.numberingKey, action2.numId, 2);
    const restartPlan = { numberingKey: 'numbered:decimal:single', startAt: 1 };
    const restartAction = resolveSingleLineListFallbackNumberingAction(restartPlan, state);
    if (restartAction.type !== 'explicitStartNew') {
        console.log('❌ FAIL: Expected explicitStartNew for restart marker sequence.', restartAction);
        return;
    }

    clearSingleLineListFallbackExplicitSequence(state, plan1.numberingKey);
    const actionAfterClear = resolveSingleLineListFallbackNumberingAction(plan2, state);
    if (actionAfterClear.type !== 'explicitStartNew') {
        console.log('❌ FAIL: Expected explicitStartNew after clearing sequence.', actionAfterClear);
        return;
    }

    const nonExplicit = resolveSingleLineListFallbackNumberingAction({ numberingKey: 'numbered:decimal:single', startAt: null }, state);
    if (nonExplicit.type !== 'sharedByStyle') {
        console.log('❌ FAIL: Expected sharedByStyle for non-explicit plan.', nonExplicit);
        return;
    }

    console.log('✅ PASS: Explicit single-line sequence helper behavior is correct.');
}

// --- Test 5e: Deterministic list binding enforcement for paragraph fragments ---
function testEnforceListBindingOnParagraphNodes() {
    console.log('\n=== Test: Enforce List Binding On Paragraph Nodes ===');
    const parser = new DOMParser();
    const frag = parser.parseFromString(`
    <root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="2"/>
                </w:numPr>
                <w:pPrChange w:id="1">
                    <w:pPr>
                        <w:numPr>
                            <w:ilvl w:val="0"/>
                            <w:numId w:val="2"/>
                        </w:numPr>
                    </w:pPr>
                </w:pPrChange>
            </w:pPr>
            <w:r><w:t>Header</w:t></w:r>
        </w:p>
    </root>`, 'application/xml');

    const paragraph = frag.getElementsByTagNameNS('*', 'p')[0];
    const updatedCount = enforceListBindingOnParagraphNodes([paragraph], {
        numId: 1000,
        ilvl: 0
    });
    const numIdNode = paragraph.getElementsByTagNameNS('*', 'numId')[0];
    const numId = numIdNode?.getAttribute('w:val') || numIdNode?.getAttribute('val') || '';
    const hasPPrChange = paragraph.getElementsByTagNameNS('*', 'pPrChange').length > 0;

    if (updatedCount === 1 && numId === '1000' && !hasPPrChange) {
        console.log('✅ PASS: List binding enforcement rewrote numId and cleared pPrChange.');
    } else {
        console.log('❌ FAIL: List binding enforcement result unexpected.', { updatedCount, numId, hasPPrChange });
    }
}

// --- Test 6: Kitchen Sink Markdown (Headings + Lists + Table) ---
async function testKitchenSinkMarkdown() {
    console.log('\n=== Test: Kitchen Sink Markdown ===');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: false,
        author: 'KitchenSink',
        numberingService: new NumberingService(),
        font: 'Calibri'
    });

    const modifiedText = `### Liquidation Instructions
Intro line for context.
1. **Securities:** Sell all stocks.
2. *Cash/Cash Equivalents:* Consolidate balances.
| Account | Number |
|---|---|
| Brokerage | 123 |
| Savings | 456 |
#### Payout Notes
1. **Recipient Bank:** [Bank Name]
2. **ABA/Routing Number:** [Routing Number]`;

    try {
        const result = await pipeline.executeListGeneration(modifiedText, null, null, "Placeholder");
        const ooxml = result.ooxml || '';

        const hasHeading3 = ooxml.includes('w:pStyle w:val="Heading3"');
        const hasHeading4 = ooxml.includes('w:pStyle w:val="Heading4"');
        const hasTable = ooxml.includes('<w:tbl');
        const hasList = ooxml.includes('<w:numPr>');
        const hasBold = /<w:b(?:\s|\/)/.test(ooxml);
        const hasItalic = /<w:i(?:\s|\/)/.test(ooxml);
        const hasTableData = ooxml.includes('Brokerage') && ooxml.includes('Savings');

        if (hasHeading3 && hasHeading4 && hasTable && hasList && hasBold && hasItalic && hasTableData) {
            console.log('✅ PASS: Headings, lists, tables, and inline formatting rendered.');
        } else {
            console.log('❌ FAIL: Missing expected elements.', {
                hasHeading3,
                hasHeading4,
                hasTable,
                hasList,
                hasBold,
                hasItalic,
                hasTableData
            });
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Test 8: Plain text to upperAlpha ordered list (the Recitals bug) ---
// Reproduces the exact scenario from the user's bug report: paragraphs starting
// with "A. ...", "B. ...", "C. ..." that need to become a proper ordered list
// with upperLetter numbering. Previously failed because:
//   1) numbering.xml was duplicated per-paragraph, causing merge conflicts
//   2) per-paragraph insertOoxml("Replace") dropped <w:pPr> under tracked changes
async function testPlainTextToUpperAlphaList() {
    console.log('\n=== Test 8: Plain Text → upperAlpha Ordered List (Recitals Bug) ===');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: false,
        author: 'Test',
        numberingService: new NumberingService()
    });

    // Simulates the Recitals section: 3 plain paragraphs with A./B./C. markers
    const modifiedText = `A. The Disclosing Party possesses certain confidential, proprietary, and trade secret information.
B. The Parties desire to enter into a potential business relationship or transaction (the "Purpose"), which requires the Disclosing Party to disclose certain Confidential Information (as defined below) to the Receiving Party.
C. The Receiving Party agrees to receive and treat such Confidential Information in confidence, subject to the terms and conditions of this Agreement.`;

    try {
        const result = await pipeline.executeListGeneration(modifiedText, null, null, "Original Text");
        const ooxml = result.ooxml || '';

        // Check: must have numPr on every paragraph
        const numPrCount = (ooxml.match(/<w:numPr>/g) || []).length;
        // Check: must use upperLetter numbering format
        const hasUpperLetter = (result.numberingXml || '').includes('w:numFmt w:val="upperLetter"');
        // Check: text should NOT retain "A. " / "B. " / "C. " prefixes (markers stripped)
        const hasLiteralMarker = ooxml.includes('>A. The Disclosing') || ooxml.includes('>B. The Parties');

        const pass = numPrCount >= 3 && hasUpperLetter && !hasLiteralMarker;
        if (pass) {
            console.log('✅ PASS: upperAlpha list created with 3 items, markers stripped, correct numFmt.');
        } else {
            console.log('❌ FAIL:', { numPrCount, hasUpperLetter, hasLiteralMarker });
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Test 9: Decimal numbered list from plain text ---
// Ensures numbered lists with decimal format also get a proper numbering
// definition (not relying on a pre-existing numId=2 in the document).
async function testDecimalNumberedList() {
    console.log('\n=== Test 9: Decimal Numbered List ===');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: false,
        author: 'Test',
        numberingService: new NumberingService()
    });

    const modifiedText = `1. First item in the list.
2. Second item in the list.
3. Third item.
4. Fourth and final item.`;

    try {
        const result = await pipeline.executeListGeneration(modifiedText, null, null, "Original");
        const ooxml = result.ooxml || '';

        const numPrCount = (ooxml.match(/<w:numPr>/g) || []).length;
        const pCount = (ooxml.match(/<w:p\b/g) || []).length;

        if (numPrCount >= 4 && pCount >= 4) {
            console.log('✅ PASS: Decimal list with 4 items, all have numPr.');
        } else {
            console.log('❌ FAIL:', { numPrCount, pCount });
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Test 10: Bullet list creation ---
async function testBulletListCreation() {
    console.log('\n=== Test 10: Bullet List Creation ===');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: false,
        author: 'Test',
        numberingService: new NumberingService()
    });

    const modifiedText = `- Confidentiality obligations
- Non-disclosure requirements
- Permitted exceptions`;

    try {
        const result = await pipeline.executeListGeneration(modifiedText, null, null, "Original");
        const ooxml = result.ooxml || '';

        const numPrCount = (ooxml.match(/<w:numPr>/g) || []).length;
        const hasBulletFmt = (result.numberingXml || '').includes('w:numFmt w:val="bullet"');
        const pCount = (ooxml.match(/<w:p\b/g) || []).length;

        if (numPrCount >= 3 && pCount >= 3) {
            console.log('✅ PASS: Bullet list with 3 items.');
        } else {
            console.log('❌ FAIL:', { numPrCount, hasBulletFmt, pCount });
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Test 11: lowerAlpha numbered list ---
// Note: "i.", "ii.", "iii." are ambiguous between lowerAlpha and lowerRoman.
// The marker detector treats single-letter markers like "i." as lowerLetter.
// Use explicit lowerAlpha markers to test this path unambiguously.
async function testLowerAlphaList() {
    console.log('\n=== Test 11: lowerAlpha Numbered List ===');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: false,
        author: 'Test',
        numberingService: new NumberingService()
    });

    const modifiedText = `a. First obligation
b. Second obligation
c. Third obligation`;

    try {
        const result = await pipeline.executeListGeneration(modifiedText, null, null, "Original");
        const ooxml = result.ooxml || '';
        const numXml = result.numberingXml || '';

        const numPrCount = (ooxml.match(/<w:numPr>/g) || []).length;
        const hasLowerLetterFmt = numXml.includes('w:numFmt w:val="lowerLetter"');
        const pCount = (ooxml.match(/<w:p\b/g) || []).length;

        if (numPrCount >= 3 && hasLowerLetterFmt && pCount >= 3) {
            console.log('✅ PASS: lowerAlpha list with 3 items, correct numFmt.');
        } else {
            console.log('❌ FAIL:', { numPrCount, hasLowerLetterFmt, pCount });
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Test 12: List with redlines (tracked changes) ---
// The original bug was that numPr was dropped under tracked changes.
// This test ensures the reconciliation pipeline generates valid OOXML
// for list content even with generateRedlines: true.
async function testListWithRedlines() {
    console.log('\n=== Test 12: List with Redlines ===');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: true,
        author: 'AI',
        numberingService: new NumberingService()
    });

    // Original is plain text, new content is a numbered list
    const originalOoxml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Original plain paragraph text.</w:t></w:r></w:p>';
    const modifiedText = `A. First recital clause.
B. Second recital clause.`;

    try {
        const result = await pipeline.execute(originalOoxml, modifiedText);
        const ooxml = result.ooxml || '';

        // Must have list structure (numPr) in the output
        const hasNumPr = ooxml.includes('<w:numPr>') || ooxml.includes('w:numId');
        // Must have at least 2 paragraphs
        const pCount = (ooxml.match(/<w:p\b/g) || []).length;
        // Must have the text content
        const hasText = ooxml.includes('First recital') && ooxml.includes('Second recital');

        if (hasNumPr && pCount >= 2 && hasText) {
            console.log('✅ PASS: List with redlines preserves numPr and content.');
        } else {
            console.log('❌ FAIL:', { hasNumPr, pCount, hasText });
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Test 13: Single numbering definition for multi-item list ---
// Verifies that only ONE numbering definition is created for all items,
// not duplicated per-paragraph (the original bug).
async function testSingleNumberingDefinition() {
    console.log('\n=== Test 13: Single Numbering Definition (no duplicates) ===');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: false,
        author: 'Test',
        numberingService: new NumberingService()
    });

    const modifiedText = `A. Item one
B. Item two
C. Item three
D. Item four
E. Item five`;

    try {
        const result = await pipeline.executeListGeneration(modifiedText, null, null, "Original");
        const numXml = result.numberingXml || '';

        // All 5 paragraphs should reference the SAME numId.
        // The original bug caused each paragraph to get its own numbering.xml
        // with a separate numId, breaking the list continuity.
        const ooxml = result.ooxml || '';
        const numIdRefs = ooxml.match(/w:numId w:val="(\d+)"/g) || [];
        const uniqueNumIds = new Set(numIdRefs.map(m => m.match(/"(\d+)"/)[1]));

        // Key check: all items share ONE numId, and we have 5 references
        if (uniqueNumIds.size === 1 && numIdRefs.length >= 5) {
            console.log(`✅ PASS: All 5 items share same numId (${[...uniqueNumIds][0]}), no per-paragraph duplication.`);
        } else {
            console.log('❌ FAIL:', { uniqueNumIds: [...uniqueNumIds], numIdRefs: numIdRefs.length });
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Test 14: Structural list conversion must bypass no-op guard ---
// Reproduces cases where text is already marker-prefixed (A./B./C.)
// and should still be converted into a true Word list.
async function testStructuralListConversionBypassesNoOp() {
    console.log('\n=== Test 14: Structural List Conversion Bypasses No-Op ===');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: false,
        author: 'Test',
        numberingService: new NumberingService()
    });

    const originalOoxml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>A. First recital clause.</w:t><w:br/><w:t>B. Second recital clause.</w:t><w:br/><w:t>C. Third recital clause.</w:t></w:r></w:p>';
    const modifiedText = `A. First recital clause.
B. Second recital clause.
C. Third recital clause.`;

    try {
        const result = await pipeline.execute(originalOoxml, modifiedText);
        const ooxml = result.ooxml || '';

        const hasNumPr = ooxml.includes('<w:numPr>');
        const pCount = (ooxml.match(/<w:p\b/g) || []).length;
        const unchanged = ooxml === originalOoxml;

        if (hasNumPr && pCount >= 3 && !unchanged) {
            console.log('✅ PASS: Structural list conversion executed even when textual content matched.');
        } else {
            console.log('❌ FAIL:', { hasNumPr, pCount, unchanged });
        }
    } catch (e) {
        console.error('❌ ERROR:', e);
    }
}

// --- Main Runner ---
(async () => {
    console.log('STARTING LIST TESTS...');

    await testListExpansionWithSoftBreak();
    await testSurgicalList();
    await testReproListIssue();
    await verifyFixedListConversion();
    await verifyListFixes();
    testInsertionOnlyNestedDepthHeuristics();
    testRedundantLeadingListMarkerStripping();
    testExplicitSingleLineFallbackSequenceHelpers();
    testEnforceListBindingOnParagraphNodes();
    await testKitchenSinkMarkdown();
    testMixedContentParsing();

    // New tests for the Recitals/ordered-list bug
    await testPlainTextToUpperAlphaList();
    await testDecimalNumberedList();
    await testBulletListCreation();
    await testLowerAlphaList();
    await testListWithRedlines();
    await testSingleNumberingDefinition();
    await testStructuralListConversionBypassesNoOp();

    console.log('\nALL LIST TESTS COMPLETE.');
})();

// --- Test 7: Mixed Content Parsing (from test_mixed_content_parsing.mjs) ---
function testMixedContentParsing() {
    console.log('\n=== Test: Mixed Content Parsing ===');

    // Case 1: Mixed Content (Preamble + List) - The User's Bug Case
    const input1 = `If the Receiving Party is required by law...
1. provide the Disclosing Party...
2. reasonably cooperate...
3. if disclosure is ultimately required...`;

    const result1 = parseMarkdownList(input1);
    if (result1 && result1.type === 'numbered' && result1.items.length === 4 && result1.items[0].type === 'text') {
        console.log('✅ PASS: Mixed Content correctly parsed');
    } else {
        console.log('❌ FAIL: Mixed Content parsing failed', result1);
    }

    // Case 2: Pure Numbered List
    const input2 = `1. Item One
2. Item Two`;
    const result2 = parseMarkdownList(input2);
    if (result2 && result2.items[0].type === 'numbered') {
        console.log('✅ PASS: Pure Numbered List correctly parsed');
    } else {
        console.log('❌ FAIL: Pure Numbered List failed');
    }

    // Case 3: Text Only
    const input3 = `Just some text.
More text.`;
    const result3 = parseMarkdownList(input3);
    if (result3 && result3.type === 'text') {
        console.log('✅ PASS: Text Only correctly parsed');
    } else {
        console.log('❌ FAIL: Text Only parsing failed');
    }

    // Case 4: Bullet List
    const input4 = `- Bullet 1
* Bullet 2`;
    const result4 = parseMarkdownList(input4);
    if (result4 && result4.type === 'bullet') {
        console.log('✅ PASS: Bullet List correctly parsed');
    } else {
        console.log('❌ FAIL: Bullet List parsing failed');
    }
}

