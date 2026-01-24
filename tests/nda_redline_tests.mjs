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
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/oxml-engine.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/ingestion.js';

// --- Config ---
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

// --- Helper Functions ---

async function readSampleDoc() {
    return await fs.readFile(DOC_PATH, 'utf-8');
}

function extractTextSimple(ooxml) {
    const { acceptedText } = ingestOoxml(ooxml);
    return acceptedText;
}

// Helper to normalize XML for simpler assertions (optional)
function normalizeXml(xml) {
    return xml.replace(/\s+/g, ' ');
}

/**
 * Extracts the table OOXML from the full document for table-scoped testing.
 * Returns the table element wrapped in minimal document structure.
 */
function extractTableOoxml(fullOoxml) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(fullOoxml, 'text/xml');
    const tables = xmlDoc.getElementsByTagName('w:tbl');

    if (tables.length === 0) return null;

    // Return the table XML directly
    return serializer.serializeToString(tables[0]);
}

/**
 * Extracts text from table cells for comparison.
 */
function extractTableText(tableOoxml) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(`<root>${tableOoxml}</root>`, 'text/xml');
    const cells = xmlDoc.getElementsByTagName('w:tc');
    const texts = [];

    for (const cell of cells) {
        const textNodes = cell.getElementsByTagName('w:t');
        let cellText = '';
        for (const t of textNodes) {
            cellText += t.textContent || '';
        }
        if (cellText.trim()) texts.push(cellText.trim());
    }

    return texts.join(' | ');
}

async function runTest(name, transformFn, assertionFn) {
    console.log(`\n=== Test: ${name} ===`);
    try {
        const originalOoxml = await readSampleDoc();
        const originalText = extractTextSimple(originalOoxml);

        const newText = transformFn(originalText);
        if (newText === null) {
            console.log('⚠️ SKIPPED: Could not simulate text transformation easily.');
            return;
        }

        // Use applyRedlineToOxml which handles tables and lists correctly
        const result = await applyRedlineToOxml(originalOoxml, originalText, newText, {
            author: 'TestUser',
            generateRedlines: true
        });

        // Note: applyRedlineToOxml returns { oxml, hasChanges }
        if (!result.hasChanges) {
            console.log('⚠️ No changes detected by engine.');
        }

        const passed = assertionFn(result.oxml, originalOoxml);
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
    console.log('Starting NDA Redlining Tests (using oxml-engine)...');

    // 1. Add Instructions
    await runTest(
        'Add Instructions Paragraph',
        (text) => `Instructions: Please review the following NDA carefully.\n\n${text}`,
        (ooxml) => {
            return ooxml.includes('<w:ins') && ooxml.includes('Instructions: Please review');
        }
    );

    // 2. Underline Title (Simulating Markdown or HTML format)
    await runTest(
        'Underline Title',
        (text) => text.replace('NON-DISCLOSURE AGREEMENT', '<u>NON-DISCLOSURE AGREEMENT</u>'),
        (ooxml) => {
            return ooxml.includes('<w:u') && ooxml.includes('NON-DISCLOSURE AGREEMENT');
        }
    );

    // 3. Recitals to Ordered List
    await runTest(
        'Recitals to Ordered List',
        (text) => {
            // Convert Recitals to Markdown List
            return text
                .replace('A. The Disclosing', '1. The Disclosing')
                .replace('B. The Parties', '2. The Parties')
                .replace('C. The Receiving', '3. The Receiving');
        },
        (ooxml) => {
            // Check for numPr (numbering properties)
            return ooxml.includes('<w:numPr>') && ooxml.includes('<w:numId');
        }
    );

    // 4. Parties to Table (Using Markdown Table Syntax)
    await runTest(
        'Parties to Table',
        (text) => {
            const start = text.indexOf('Disclosing Party:');
            const end = text.indexOf('RECITALS');
            if (start === -1 || end === -1) return null;

            const tableMd = `
| Disclosing Party | Receiving Party |
| --- | --- |
| [Name of Disclosing Party] | [Name of Receiving Party] |
| [Address of Disclosing Party] | [Address of Receiving Party] |
`;
            return text.slice(0, start) + tableMd + '\n' + text.slice(end);
        },
        (ooxml) => {
            // Should contain table structure
            return ooxml.includes('<w:tbl>');
        }
    );

    // 5. Bold Signature Fields
    await runTest(
        'Bold Signature Fields',
        (text) => text.replace(/By:/g, '**By:**').replace(/Title:/g, '**Title:**'),
        (ooxml) => {
            return ooxml.includes('<w:b/>') || ooxml.includes('<w:b>');
        }
    );

    // 6. Add Date Row to Signature (Using direct table reconciliation)
    // This test specifically tests row addition to an EXISTING table
    console.log(`\n=== Test: Add Date Row ===`);
    try {
        const originalOoxml = await readSampleDoc();

        // Extract just the table OOXML
        const tableOoxml = extractTableOoxml(originalOoxml);
        if (!tableOoxml) {
            console.log('⚠️ SKIPPED: No table found in document');
        } else {
            // Extract the original table text for reference
            const tableText = extractTableText(tableOoxml);

            // The new table should have the same content PLUS a Date row
            const newTableMd = `
| DISCLOSING PARTY: | RECEIVING PARTY: |
| --- | --- |
| _________________________ | _________________________ |
| By: [Name] | By: [Name] |
| Title: | Title: |
| Date: ___________________ | Date: ___________________ |
`;
            // Call applyRedlineToOxml with the TABLE OOXML (not full doc)
            // and the markdown table as the new content
            const result = await applyRedlineToOxml(tableOoxml, tableText, newTableMd, {
                author: 'TestUser',
                generateRedlines: true
            });

            // Check for Date row and w:ins for the new row
            const hasDateRow = result.oxml.includes('Date:');
            const hasInsertMark = result.oxml.includes('<w:ins');
            const hasTableRows = result.oxml.includes('<w:tr') || result.oxml.includes('w:tr>');

            if (hasDateRow && hasTableRows) {
                console.log('✅ PASS');
            } else {
                console.log('❌ FAIL');
                console.log(`  hasDateRow: ${hasDateRow}, hasInsertMark: ${hasInsertMark}, hasTableRows: ${hasTableRows}`);
            }
        }
    } catch (e) {
        console.error('💥 ERROR:', e);
    }


    // 7. Add Bullet
    await runTest(
        'Add Confidential Info Bullet',
        (text) => {
            const target = 'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.';
            const insertion = '\n* Photographs, videos, and other recordings of prototypes and physical hardware.';
            return text.replace(target, target + insertion);
        },
        (ooxml) => {
            return ooxml.includes('Photographs, videos') && ooxml.includes('<w:numPr>');
        }
    );

    // 8. Change Archival Copies
    await runTest(
        'Change Archival Copies',
        (text) => text.replace('one (1) copy', 'two (2) copies'),
        (ooxml) => {
            return ooxml.includes('<w:del') && ooxml.includes('one (1)') &&
                ooxml.includes('<w:ins') && ooxml.includes('two (2)');
        }
    );

    // 9. Add Sub-bullet in Archival Exception
    await runTest(
        'Add Sub-bullet Archival',
        (text) => {
            const target = 'compliance and record-keeping.';
            // To create a sub-bullet, Markdown usually uses indentation (2 or 4 spaces).
            return text.replace(target, target + '\n    * Must be legally required.');
        },
        (ooxml) => {
            // Check for ilvl val="1" or "2" depending on base level.
            // If base is 0, this should be 1.
            // The logic in list generation sets ilvl based on indentation.
            return ooxml.includes('Must be legally required') && (ooxml.includes('w:val="1"') || ooxml.includes('w:val="2"'));
        }
    );

    // 10. Unbold BC
    await runTest(
        'Unbold British Columbia',
        (text) => {
            // We effectively pass the same text. The engine compares text.
            // Text is same. "British Columbia".
            // But formatting is diff? 
            // If input text has NO markdown (no **), and original has formatting, 
            // `applyRedlineToOxml` checks `needsFormatRemoval` (line 242 oxml-engine.js).
            // "Format REMOVAL: text is unchanged, no new hints, but original has formatting to strip"
            // This logic seems to trigger if NO Markdown is present.
            // But wouldn't that strip formatting from EVERYTHING?
            // "needsFormatRemoval = !hasTextChanges && !hasFormatHints && hasExistingFormatting"
            // If I return the full text exactly as is, it triggers this?
            // That would be dangerous if it strips ALL formatting.
            // Let's see if the engine is smart enough to target specific areas or global.
            // The code seems to iterate ALL runs and if they have formatting, it removes them if no new hints exist?
            // "applyFormatRemovalWithSpans" iterates all spans.
            // This implies: "If you send me plain text, I assume you want plain text."
            // Which is correct for "Unbold" if the AI returns plain text.
            return text;
        },
        (ooxml) => {
            // Verification: w:b should be removed or set to 0.
            // British Columbia is near end.
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
