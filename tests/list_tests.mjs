
import { JSDOM } from 'jsdom';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { ReconciliationPipeline } from '../src/taskpane/modules/reconciliation/pipeline.js';
import { NumberingService } from '../src/taskpane/modules/reconciliation/numbering-service.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/ingestion.js';
import { serializeToOoxml } from '../src/taskpane/modules/reconciliation/serialization.js';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/oxml-engine.js';

// --- Global Setup ---
const dom = new JSDOM('<!DOCTYPE html><html><body></body></html>');
global.window = dom.window;
global.document = dom.window.document;
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.Node = dom.window.Node;

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
        const pTags = (result.ooxml.match(/<w:p/g) || []);
        const pCount = pTags.length;

        if (pCount === 3) {
            console.log('✅ SUCCESS: Correctly expanded into 3 paragraphs.');
        } else {
            console.error('❌ FAILED: Expected 3 paragraphs, got ' + pCount);
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

// --- Main Runner ---
(async () => {
    console.log('STARTING LIST TESTS...');

    await testListExpansionWithSoftBreak();
    await testSurgicalList();
    await testReproListIssue();
    await verifyFixedListConversion();
    await verifyListFixes();
    testMixedContentParsing();

    console.log('\nALL LIST TESTS COMPLETE.');
})();

// --- Test 6: Mixed Content Parsing (from test_mixed_content_parsing.mjs) ---
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
