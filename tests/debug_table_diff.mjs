/**
 * Detailed diagnostic for table row insertion
 */
import { JSDOM } from 'jsdom';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

import { parseTable } from '../src/taskpane/modules/reconciliation/pipeline.js';
import { ingestTableToVirtualGrid } from '../src/taskpane/modules/reconciliation/ingestion.js';
import { diffTablesWithVirtualGrid, serializeVirtualGridToOoxml } from '../src/taskpane/modules/reconciliation/table-reconciliation.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/ingestion.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

async function runDiagnostic() {
    const originalOoxml = await fs.readFile(DOC_PATH, 'utf-8');
    const { acceptedText } = ingestOoxml(originalOoxml);

    console.log('=== Table Diff Operations Diagnostic ===\n');

    // Parse the full XML to find tables
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(originalOoxml, 'text/xml');
    const tables = xmlDoc.getElementsByTagName('w:tbl');

    console.log('Number of tables in document:', tables.length);

    if (tables.length === 0) {
        console.log('No tables found in document');
        return;
    }

    // Use the first (and likely only) table
    const table = tables[0];

    // Ingest to Virtual Grid
    const oldGrid = ingestTableToVirtualGrid(table);
    console.log('\n=== Old Grid ===');
    console.log('rowCount:', oldGrid.rowCount);
    console.log('colCount:', oldGrid.colCount);

    // Display old grid content
    for (let r = 0; r < oldGrid.rowCount; r++) {
        const rowTexts = [];
        for (let c = 0; c < oldGrid.colCount; c++) {
            const cell = oldGrid.grid[r]?.[c];
            if (cell && !cell.isMergeContinuation) {
                rowTexts.push(cell.getText().substring(0, 30));
            }
        }
        console.log(`  Row ${r}: [${rowTexts.join(' | ')}]`);
    }

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
    console.log('\n=== New Table Data ===');
    console.log('hasHeader:', newTableData.hasHeader);
    console.log('headers:', newTableData.headers);
    console.log('rows count:', newTableData.rows.length);
    newTableData.rows.forEach((row, i) => console.log(`  Row ${i}:`, row));

    // Compute diff
    const operations = diffTablesWithVirtualGrid(oldGrid, newTableData);
    console.log('\n=== Diff Operations ===');
    console.log('Total operations:', operations.length);
    operations.forEach((op, i) => {
        console.log(`  [${i}] ${op.type}:`, op.type === 'row_insert' ? op.cells : `row=${op.gridRow}, col=${op.gridCol || 'N/A'}`);
    });

    // Serialize
    console.log('\n=== Serializing... ===');
    const reconciled = serializeVirtualGridToOoxml(oldGrid, operations, { generateRedlines: true, author: 'TestUser' });

    // Check for Date row
    console.log('Has Date: in output:', reconciled.includes('Date:'));
    console.log('Has w:ins:', reconciled.includes('<w:ins'));
    console.log('Output length:', reconciled.length);

    // Find the row_insert operations specifically
    const insertOps = operations.filter(o => o.type === 'row_insert');
    console.log('\nRow insert operations:', insertOps.length);
    insertOps.forEach(op => console.log('  Insert row:', op.gridRow, 'cells:', op.cells));

    // Save output for inspection
    await fs.writeFile(path.join(__dirname, 'table_diff_output.xml'), reconciled);
    console.log('\nOutput saved to tests/table_diff_output.xml');
}

runDiagnostic().catch(console.error);
