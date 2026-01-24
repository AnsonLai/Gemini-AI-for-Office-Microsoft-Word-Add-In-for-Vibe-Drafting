/**
 * Diagnostic test for applyTableReconciliation flow
 */
import { JSDOM } from 'jsdom';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/oxml-engine.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

async function runDiagnostic() {
    console.log('=== applyTableReconciliation Flow Diagnostic ===\n');

    const originalOoxml = await fs.readFile(DOC_PATH, 'utf-8');

    // Extract just the table OOXML
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(originalOoxml, 'text/xml');
    const tables = xmlDoc.getElementsByTagName('w:tbl');

    if (tables.length === 0) {
        console.log('No tables found');
        return;
    }

    const tableOoxml = serializer.serializeToString(tables[0]);
    console.log('Table OOXML preview (first 500 chars):\n', tableOoxml.substring(0, 500));
    console.log('\n---\n');

    // The new markdown table with Date row
    const newTableMd = `
| DISCLOSING PARTY: | RECEIVING PARTY: |
| --- | --- |
| _________________________ | _________________________ |
| By: [Name] | By: [Name] |
| Title: | Title: |
| Date: ___________________ | Date: ___________________ |
`;

    // Get table text (simulate what the test does)
    const textNodes = tables[0].getElementsByTagName('w:t');
    let tableText = '';
    for (const t of textNodes) {
        tableText += (t.textContent || '') + ' ';
    }
    console.log('Table text extracted:', tableText.trim().substring(0, 100));
    console.log('\n---\n');

    console.log('Calling applyRedlineToOxml...');
    const result = await applyRedlineToOxml(tableOoxml, tableText.trim(), newTableMd, {
        author: 'TestUser',
        generateRedlines: true
    });

    console.log('hasChanges:', result.hasChanges);
    console.log('Output length:', result.oxml.length);
    console.log('Output preview:\n', result.oxml.substring(0, 1000));
    console.log('\n---\n');

    console.log('Has w:tbl:', result.oxml.includes('<w:tbl>') || result.oxml.includes('w:tbl>'));
    console.log('Has w:tr:', result.oxml.includes('<w:tr') || result.oxml.includes('w:tr>'));
    console.log('Has Date:', result.oxml.includes('Date:'));
    console.log('Has w:ins:', result.oxml.includes('<w:ins'));

    // Save output for inspection
    await fs.writeFile(path.join(__dirname, 'applyTableReconciliation_output.xml'), result.oxml);
    console.log('\nOutput saved to tests/applyTableReconciliation_output.xml');
}

runDiagnostic().catch(console.error);
