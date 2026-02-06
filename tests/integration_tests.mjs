import './setup-xml-provider.mjs';

/**
 * Integration Tests
 * 
 * Combines NDA Redlining Tests and Redline Toggle Tests.
 * Run with: node --experimental-modules tests/integration_tests.mjs
 */

import { JSDOM } from 'jsdom';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/engine/oxml-engine.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/pipeline/ingestion.js';

// --- Mock Browser Environment ---
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

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

function extractTableOoxml(fullOoxml) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(fullOoxml, 'text/xml');
    const tables = xmlDoc.getElementsByTagName('w:tbl');
    if (tables.length === 0) return null;
    return serializer.serializeToString(tables[0]);
}

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
        const result = await applyRedlineToOxml(originalOoxml, originalText, newText, {
            author: 'TestUser',
            generateRedlines: true
        });
        if (!result.hasChanges) {
            console.log('⚠️ No changes detected by engine.');
        }
        const passed = assertionFn(result.oxml, originalOoxml);
        if (passed) console.log('✅ PASS');
        else console.log('❌ FAIL');
    } catch (e) {
        console.error('💥 ERROR:', e);
    }
}

// --- NDA Redline Tests ---
async function runNdaTests() {
    console.log('Starting NDA Redlining Tests...');

    await runTest('Add Instructions Paragraph',
        (text) => `Instructions: Please review the following NDA carefully.\n\n${text}`,
        (ooxml) => ooxml.includes('<w:ins') && ooxml.includes('Instructions: Please review')
    );

    await runTest('Underline Title',
        (text) => text.replace('NON-DISCLOSURE AGREEMENT', '<u>NON-DISCLOSURE AGREEMENT</u>'),
        (ooxml) => ooxml.includes('<w:u') && ooxml.includes('NON-DISCLOSURE AGREEMENT')
    );

    await runTest('Recitals to Ordered List',
        (text) => text.replace('A. The Disclosing', '1. The Disclosing')
            .replace('B. The Parties', '2. The Parties')
            .replace('C. The Receiving', '3. The Receiving'),
        (ooxml) => ooxml.includes('<w:numPr>') && ooxml.includes('<w:numId')
    );

    await runTest('Parties to Table',
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
        (ooxml) => ooxml.includes('<w:tbl>')
    );

    await runTest('Bold Signature Fields',
        (text) => text.replace(/By:/g, '**By:**').replace(/Title:/g, '**Title:**'),
        (ooxml) => ooxml.includes('<w:b/>') || ooxml.includes('<w:b>')
    );

    // Add Date Row
    console.log(`\n=== Test: Add Date Row ===`);
    try {
        const originalOoxml = await readSampleDoc();
        const tableOoxml = extractTableOoxml(originalOoxml);
        if (!tableOoxml) {
            console.log('⚠️ SKIPPED: No table found in document');
        } else {
            const tableText = extractTableText(tableOoxml);
            const newTableMd = `
| DISCLOSING PARTY: | RECEIVING PARTY: |
| --- | --- |
| _________________________ | _________________________ |
| By: [Name] | By: [Name] |
| Title: | Title: |
| Date: ___________________ | Date: ___________________ |
`;
            const result = await applyRedlineToOxml(tableOoxml, tableText, newTableMd, {
                author: 'TestUser',
                generateRedlines: true
            });
            const hasDateRow = result.oxml.includes('Date:');
            const hasTableRows = result.oxml.includes('<w:tr') || result.oxml.includes('w:tr>');
            if (hasDateRow && hasTableRows) console.log('✅ PASS');
            else console.log('❌ FAIL');
        }
    } catch (e) {
        console.error('💥 ERROR:', e);
    }

    await runTest('Add Confidential Info Bullet',
        (text) => {
            const target = 'Technical data, specifications, designs, prototypes, software, algorithms, source code, and intellectual property.';
            const insertion = '\n* Photographs, videos, and other recordings of prototypes and physical hardware.';
            return text.replace(target, target + insertion);
        },
        (ooxml) => ooxml.includes('Photographs, videos') && ooxml.includes('<w:numPr>')
    );

    await runTest('Change Archival Copies',
        (text) => text.replace('one (1) copy', 'two (2) copies'),
        (ooxml) => ooxml.includes('<w:del') && ooxml.includes('one (1)') && ooxml.includes('<w:ins') && ooxml.includes('two (2)')
    );

    await runTest('Add Sub-bullet Archival',
        (text) => {
            const target = 'compliance and record-keeping.';
            return text.replace(target, target + '\n    * Must be legally required.');
        },
        (ooxml) => ooxml.includes('Must be legally required') && (ooxml.includes('w:val="1"') || ooxml.includes('w:val="2"'))
    );

    await runTest('Unbold British Columbia',
        (text) => text,
        (ooxml) => !ooxml.match(/<w:b\/>\s*<w:t>British Columbia/)
    );

    await runTest('Rewrite Disclosure Provision',
        (text) => {
            const oldSection = 'If the Receiving Party is required by law, regulation, or court order to disclose any Confidential Information, the Receiving Party shall, to the extent legally permitted, provide the Disclosing Party with prompt written notice prior to such disclosure so that the Disclosing Party may seek a protective order or other appropriate remedy.';
            const newSection = 'The Receiving Party agrees to fight any orders to disclose Confidential Information. They will involve the Disclosing Party immediately. Only after all legal avenues are exhausted may they disclose the information.';
            return text.replace(oldSection, newSection);
        },
        (ooxml) => ooxml.includes('agrees to fight any orders') && ooxml.includes('<w:del')
    );
}

// --- Redline Toggle Tests ---
async function runRedlineToggleTests() {
    console.log('\nStarting Redline Toggle Regression Tests...');

    const originalOoxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:t>Hello world</w:t>
            </w:r>
        </w:p>
    `;
    const originalText = "Hello world";
    const modifiedText = "Hello Gemini";

    console.log('\n--- Test 1: Redlines ENABLED ---');
    const resultEnabled = await applyRedlineToOxml(originalOoxml, originalText, modifiedText, {
        author: 'TestUser',
        generateRedlines: true
    });
    if (resultEnabled.oxml.includes('<w:ins') && resultEnabled.oxml.includes('<w:del')) {
        console.log('✅ PASS: Track changes generated in redline mode');
    } else {
        console.log('❌ FAIL: Track changes NOT generated in redline mode');
    }

    console.log('\n--- Test 2: Redlines DISABLED ---');
    const resultDisabled = await applyRedlineToOxml(originalOoxml, originalText, modifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });
    const hasInsDisabled = resultDisabled.oxml.includes('<w:ins');
    const hasDelDisabled = resultDisabled.oxml.includes('<w:del');
    const hasNewText = resultDisabled.oxml.includes('Gemini');

    if (!hasInsDisabled && !hasDelDisabled && hasNewText) {
        console.log('✅ PASS: No track changes generated when disabled');
    } else {
        console.log(`❌ FAIL: Redline toggle NOT honored. Ins: ${hasInsDisabled}, Del: ${hasDelDisabled}, NewText: ${hasNewText}`);
        console.log('Partial Output:', resultDisabled.oxml.substring(0, 200));
    }


    console.log('\n--- Test 3: List Expansion with Redlines DISABLED ---');
    const listModifiedText = "Original text\n* Item 1\n* Item 2";
    const resultListDisabled = await applyRedlineToOxml(originalOoxml, originalText, listModifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });
    if (!resultListDisabled.oxml.includes('<w:ins') && resultListDisabled.oxml.includes('Item 1')) {
        console.log('✅ PASS: List expansion honors redline toggle');
    } else {
        console.log('❌ FAIL: List expansion redline toggle issue');
    }

    console.log('\n--- Test 4: Table Reconciliation with Redlines DISABLED ---');
    const tableOriginalOoxml = `
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:tblGrid><w:gridCol/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
    `;
    const tableOriginalText = "Header 1";
    const tableModifiedText = "| Header Updated |\n| --- |";
    const resultTableDisabled = await applyRedlineToOxml(tableOriginalOoxml, tableOriginalText, tableModifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });

    // Strict check to avoid matching <w:insideH> in table properties
    const hasInsTable = /<w:ins\b/.test(resultTableDisabled.oxml);
    const hasDelTable = /<w:del\b/.test(resultTableDisabled.oxml);
    const hasUpdatedText = resultTableDisabled.oxml.includes('Header Updated');

    if (!hasInsTable && !hasDelTable && hasUpdatedText) {

        console.log('✅ PASS: Table reconciliation honors redline toggle');
    } else {
        console.log(`❌ FAIL: Table reconciliation redline toggle issue. Ins: ${hasInsTable}, Del: ${hasDelTable}, UpdatedText: ${hasUpdatedText}`);
        console.log('Output Preview:', resultTableDisabled.oxml.substring(0, 500));
        console.log('Includes Header Updated:', resultTableDisabled.oxml.includes('Header Updated'));
    }

}

// --- Main Runner ---
(async () => {
    await runNdaTests();
    await runRedlineToggleTests();
    console.log('\nALL INTEGRATION TESTS COMPLETE.');
})();

