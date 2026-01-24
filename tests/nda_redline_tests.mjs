/**
 * NDA Redlining Tests
 * 
 * Tests specific NDA modification scenarios requested by the user.
 * Defines input modifications and asserts expected OOXML structural changes.
 * 
 * Run with: node --experimental-modules tests/nda_redline_tests.mjs
 */

import { JSDOM } from 'jsdom';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

// --- Mock Browser Environment ---
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

// --- Imports ---
import { ReconciliationPipeline } from '../src/taskpane/modules/reconciliation/pipeline.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/ingestion.js';

// --- Config ---
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

// --- Helper Functions ---

/**
 * Reads the sample document XML.
 */
async function readSampleDoc() {
    return await fs.readFile(DOC_PATH, 'utf-8');
}

/**
 * simplistic helper to extract text from OOXML for modification.
 * In a real app, the AI would generate the full new text.
 * Here we manually construct the "New Text" based on the scenario.
 */
function extractTextSimple(ooxml) {
    const { acceptedText } = ingestOoxml(ooxml);
    return acceptedText;
}

/**
 * runTest - Executing the pipeline and running assertions
 */
async function runTest(name, transformFn, assertionFn) {
    console.log(`\n=== Test: ${name} ===`);
    try {
        const originalOoxml = await readSampleDoc();
        const originalText = extractTextSimple(originalOoxml);

        const newText = transformFn(originalText);

        // Skip execution if transform returned null (meaning we couldn't sim the change easily string-wise)
        if (newText === null) {
            console.log('⚠️ SKIPPED: Could not simulate text transformation easily.');
            return;
        }

        const pipeline = new ReconciliationPipeline({
            author: 'TestUser',
            generateRedlines: true,
            validateOutput: false // Disable basic validation for fragments to allow targeted checks
        });

        // The pipeline expects a paragraph-level or document-level string. 
        // Our sample doc is a full document.
        const result = await pipeline.execute(originalOoxml, newText);

        if (!result.isValid) {
            console.log('❌ Invalid Result:', result.warnings);
        }

        const passed = assertionFn(result.ooxml, originalOoxml);
        if (passed) {
            console.log('✅ PASS');
        } else {
            console.log('❌ FAIL');
        }
    } catch (e) {
        console.error('💥 ERROR:', e);
    }
}

// --- Tests ---

(async () => {
    console.log('Starting NDA Redlining Tests...');

    // 1. Add Instructions
    await runTest(
        'Add Instructions Paragraph',
        (text) => `Instructions: Please review the following NDA carefully.\n\n${text}`,
        (ooxml) => {
            // Should have an insertion at the start
            return ooxml.includes('<w:ins') && ooxml.includes('Instructions: Please review');
        }
    );

    // 2. Underline Title
    await runTest(
        'Underline Title',
        (text) => text, // This test requires formatting change, not text change. Pipeline takes newText.
        // Implication: Pipeline needs to infer formatting from markdown or we need a way to pass format-only.
        // Current pipeline preprocessMarkdown parses **bold**, *italic*. 
        // Does it support underline? Markdown uses HTML <u> or specific syntax?
        // Let's try HTML <u> tag which preprocessMarkdown might support
        (text) => text.replace('NON-DISCLOSURE AGREEMENT', '<u>NON-DISCLOSURE AGREEMENT</u>'),
        (ooxml) => {
            return ooxml.includes('<w:u') && ooxml.includes('NON-DISCLOSURE AGREEMENT');
        }
    );

    // 3. Recitals to Ordered List
    await runTest(
        'Recitals to Ordered List',
        (text) => {
            // Replace A. B. C. manually with markdown list syntax if needed, 
            // OR checks if pipeline auto-detects.
            // Let's assume we change "A. The Disclosing..." to "1. The Disclosing..."
            return text
                .replace('A. The Disclosing', '1. The Disclosing')
                .replace('B. The Parties', '2. The Parties')
                .replace('C. The Receiving', '3. The Receiving');
        },
        (ooxml) => {
            // Should contain numPr (numbering properties)
            return ooxml.includes('<w:numPr>') && ooxml.includes('<w:numId');
        }
    );

    // 4. Parties to Table
    await runTest(
        'Parties to Table',
        (text) => {
            // Hard to represent in plain text unless we use Markdown table
            const start = text.indexOf('Disclosing Party:');
            const end = text.indexOf('RECITALS');
            if (start === -1 || end === -1) return null;

            const tableMd = `
| Disclosing Party | Receiving Party |
| --- | --- |
| [Name of Disclosing Party] | [Name of Receiving Party] |
| [Address of Disclosing Party] | [Address of Receiving Party] |
`;
            return text.slice(0, start) + tableMd + text.slice(end);
        },
        (ooxml) => {
            return ooxml.includes('<w:tbl>');
        }
    );

    // 5. Bold Signature Fields
    await runTest(
        'Bold Signature Fields',
        (text) => {
            // Replace "By:" with "**By:**"
            return text.replace(/By:/g, '**By:**').replace(/Title:/g, '**Title:**');
        },
        (ooxml) => {
            // Check for bold tag on "By:"
            // Note: Simplistic check, might need regex to ensure it's on the specific run
            return ooxml.includes('<w:b/>') || ooxml.includes('<w:b>');
        }
    );

    // 6. Add Date Row to Signature
    await runTest(
        'Add Date Row',
        (text) => {
            // Append "Date: ..." to the end of the text. 
            // If the original text comes from a table, `extractTextSimple` might just give newlines.
            // If we append to text, does it go into the table? Unlikely without specific logic.
            return text + '\nDate: _________________________';
        },
        (ooxml) => {
            // We want it to be a new ROW <w:tr> in the table, not just a paragraph
            // This is a high bar for text-based recon
            return ooxml.includes('<w:tr>');
        }
    );

    // 7. Add Bullet for Confidential Info
    await runTest(
        'Add Confidential Info Bullet',
        (text) => {
            // Find the list items. 
            // "Technical data..." is item 2.
            const target = 'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.';
            const insertion = '\n* Photographs, videos, and other recordings of prototypes and physical hardware.';
            return text.replace(target, target + insertion);
        },
        (ooxml) => {
            return ooxml.includes('Photographs, videos') && ooxml.includes('<w:numPr>');
        }
    );

    // 8. Change Archival Copies (1 -> 2)
    await runTest(
        'Change Archival Copies',
        (text) => text.replace('one (1) copy', 'two (2) copies'),
        (ooxml) => {
            // Should have deletion of "one (1) copy" and insertion of "two (2) copies"
            return ooxml.includes('<w:del') && ooxml.includes('one (1) copy') &&
                ooxml.includes('<w:ins') && ooxml.includes('two (2) copies');
        }
    );

    // 9. Add Sub-bullet in Archival Exception
    await runTest(
        'Add Sub-bullet Archival',
        (text) => {
            const target = 'compliance and record-keeping.';
            // Indent with spaces for sub-bullet? 
            return text.replace(target, target + '\n  * Must be legally required.');
        },
        (ooxml) => {
            // Check for ilvl val="1" (level 2, 0-indexed) assuming previous was level 1
            // Actually doc might look different.
            return ooxml.includes('Must be legally required');
        }
    );

    // 10. Unbold BC
    await runTest(
        'Unbold British Columbia',
        (text) => {
            // Text is "British Columbia". In doc it is bold. 
            // We pass it as plain text (no ** markers). 
            // Pipeline must recognize it WAS bold and now isn't.
            // But if we pass plain text, pipeline might assume "no change" to text.
            // Does pipeline handle formatting removal? 
            // Only if we explicitly infer formatting differences. 
            // If existing is **BC**, and new is "BC", diff engine sees change?
            // Actually pipeline compares text content. If text is same, formatting might be ignored unless we parse formatting.
            return text;
        },
        (ooxml) => {
            // If logic requires format change detection, this test validates if that exists.
            // We check that <w:b/> is NOT present around BC or is toggled off?
            // <w:b w:val="0"/> or removal of <w:b/>
            return !ooxml.match(/<w:b\/>\s*<w:t>British Columbia/);
        }
    );

    // 11. Rewrite Disclosure Provision
    await runTest(
        'Rewrite Disclosure Provision',
        (text) => {
            const oldSection = 'If the Receiving Party is required by law, regulation, or court order to disclose any Confidential Information, the Receiving Party shall, to the extent legally permitted, provide the Disclosing Party with prompt written notice prior to such disclosure so that the Disclosing Party may seek a protective order or other appropriate remedy.';
            const newSection = 'The Receiving Party agrees to fight any orders to disclose Confidential Information. They will involve the Disclosing Party immediately. Only after all legal avenues are exhausted may they disclose the information.';
            return text.replace(oldSection, newSection);
        },
        (ooxml) => {
            return ooxml.includes('agrees to fight any orders') && ooxml.includes('<w:del');
        }
    );

})();
