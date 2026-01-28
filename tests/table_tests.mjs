
import { JSDOM } from 'jsdom';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/oxml-engine.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/ingestion.js';
import { parseTable } from '../src/taskpane/modules/reconciliation/pipeline.js';
import { ingestTableToVirtualGrid } from '../src/taskpane/modules/reconciliation/ingestion.js';
import { diffTablesWithVirtualGrid, serializeVirtualGridToOoxml } from '../src/taskpane/modules/reconciliation/table-reconciliation.js';

// --- Global Setup ---
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// --- Test 1: Repro Table Issue (from repro_table_issue.mjs) ---
async function testReproTableIssue() {
    console.log('\n=== Test: Repro Table Issue (Parties to Table) ===');
    const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

    try {
        const originalOoxml = await fs.readFile(DOC_PATH, 'utf-8');
        const { acceptedText: originalText } = ingestOoxml(originalOoxml);

        // Construct new text with Markdown Table for Parties
        const start = originalText.indexOf('Disclosing Party:');
        const end = originalText.indexOf('RECITALS');

        if (start === -1 || end === -1) {
            console.log('⚠️ SKIPPING: Could not find "Disclosing Party:" or "RECITALS" in sample doc.');
            return;
        }

        const tableMd = `
| Disclosing Party | Receiving Party |
| --- | --- |
| [Name of Disclosing Party] | [Name of Receiving Party] |
| [Address of Disclosing Party] | [Address of Receiving Party] |
`;
        const newText = originalText.slice(0, start) + tableMd + '\n' + originalText.slice(end);

        console.log('Applying Redline with Markdown Table...');
        const result = await applyRedlineToOxml(originalOoxml, originalText, newText, {
            author: 'DebugUser',
            generateRedlines: true
        });

        console.log('Has Changes:', result.hasChanges);
        console.log('Result OXML contains w:tbl:', result.oxml.includes('<w:tbl>'));

        // Check if the NEW table text exists
        if (result.oxml.includes('[Name of Disclosing Party]')) {
            console.log('✅ PASS: New table content found.');
        } else {
            console.log('❌ FAIL: New table content missing.');
        }

    } catch (e) {
        console.error('Test Error:', e);
    }
}

// --- Test 2: Table Diff Diagnostic (from debug_table_diff.mjs) ---
async function testTableDiffDiagnostic() {
    console.log('\n=== Test: Table Diff Diagnostic ===');
    const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

    try {
        const originalOoxml = await fs.readFile(DOC_PATH, 'utf-8');

        // Parse the full XML to find tables
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(originalOoxml, 'text/xml');
        const tables = xmlDoc.getElementsByTagName('w:tbl');

        if (tables.length === 0) {
            console.log('⚠️ SKIPPING: No tables found in document');
            return;
        }

        // Use the first (and likely only) table
        const table = tables[0];
        const oldGrid = ingestTableToVirtualGrid(table);

        // New table markdown
        const newTableMd = `
| DISCLOSING PARTY: | RECEIVING PARTY: |
| --- | --- |
| _________________________ | _________________________ |
| By: [Name] | By: [Name] |
| Title: | Title: |
| Date: ___________________ | Date: ___________________ |
`;

        const newTableData = parseTable(newTableMd);
        const operations = diffTablesWithVirtualGrid(oldGrid, newTableData);

        console.log('Total operations:', operations.length);

        // Serialize
        const reconciled = serializeVirtualGridToOoxml(oldGrid, operations, { generateRedlines: true, author: 'TestUser' });

        // Check for Date row
        const hasDate = reconciled.includes('Date:');
        const hasIns = reconciled.includes('<w:ins');

        if (hasDate && hasIns) {
            console.log('✅ PASS: Table diff applied and serialized with redlines.');
        } else {
            console.log('❌ FAIL: Table diff missing content or redlines.');
        }

    } catch (e) {
        console.error('Test Error:', e);
    }
}

// --- Test 3: Table Reconciliation Flow (from debug_table_reconciliation.mjs) ---
async function testTableReconciliationFlow() {
    console.log('\n=== Test: Table Reconciliation Flow ===');
    const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

    try {
        const originalOoxml = await fs.readFile(DOC_PATH, 'utf-8');

        // Extract just the table OOXML
        const parser = new DOMParser();
        const serializer = new XMLSerializer();
        const xmlDoc = parser.parseFromString(originalOoxml, 'text/xml');
        const tables = xmlDoc.getElementsByTagName('w:tbl');

        if (tables.length === 0) {
            console.log('⚠️ SKIPPING: No tables found');
            return;
        }

        const tableOoxml = serializer.serializeToString(tables[0]);

        // Get table text
        const textNodes = tables[0].getElementsByTagName('w:t');
        let tableText = '';
        for (const t of textNodes) {
            tableText += (t.textContent || '') + ' ';
        }

        // The new markdown table with Date row
        const newTableMd = `
| DISCLOSING PARTY: | RECEIVING PARTY: |
| --- | --- |
| _________________________ | _________________________ |
| By: [Name] | By: [Name] |
| Title: | Title: |
| Date: ___________________ | Date: ___________________ |
`;

        const result = await applyRedlineToOxml(tableOoxml, tableText.trim(), newTableMd, {
            author: 'TestUser',
            generateRedlines: true
        });

        if (result.oxml.includes('Date:') && result.oxml.includes('<w:ins')) {
            console.log('✅ PASS: applyRedlineToOxml handled table update.');
        } else {
            console.log('❌ FAIL: applyRedlineToOxml failed on table.');
        }

    } catch (e) {
        console.error('Test Error:', e);
    }
}

// --- Main Runner ---
(async () => {
    console.log('STARTING TABLE TESTS...');

    await testReproTableIssue();
    await testTableDiffDiagnostic();
    await testTableReconciliationFlow();

    console.log('\nALL TABLE TESTS COMPLETE.');
})();
