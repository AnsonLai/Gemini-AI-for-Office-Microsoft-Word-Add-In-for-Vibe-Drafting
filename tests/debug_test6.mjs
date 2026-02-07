import './setup-xml-provider.mjs';
/**
 * Diagnostic test for table row addition
 */
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/engine/oxml-engine.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/pipeline/ingestion.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

async function runDiagnostic() {
    const originalOoxml = await fs.readFile(DOC_PATH, 'utf-8');
    const { acceptedText } = ingestOoxml(originalOoxml);

    console.log('=== Add Date Row Diagnostic ===\n');

    // Build the new table with Date row added
    const newTable = `
| DISCLOSING PARTY: | RECEIVING PARTY: |
| --- | --- |
| _________________________ | _________________________ |
| By: [Name] | By: [Name] |
| Title: | Title: |
| Date: ___________________ | Date: ___________________ |
`;

    const start = acceptedText.indexOf('DISCLOSING PARTY:');
    if (start === -1) {
        console.log('ERROR: Could not find DISCLOSING PARTY in text');
        return;
    }

    console.log('Found DISCLOSING PARTY: at position', start);
    console.log('Original text around signature:');
    console.log(acceptedText.substring(start, start + 300));
    console.log('\n---\n');

    const newText = acceptedText.substring(0, start) + newTable;

    console.log('Calling applyRedlineToOxml...');
    const result = await applyRedlineToOxml(originalOoxml, acceptedText, newText, {
        author: 'TestUser',
        generateRedlines: true
    });

    console.log('hasChanges:', result.hasChanges);

    // Check for Date row
    const hasDateRow = result.oxml.includes('Date:') && result.oxml.includes('<w:tr>');
    console.log('Has Date: in output:', result.oxml.includes('Date:'));
    console.log('Has <w:tr>:', result.oxml.includes('<w:tr>'));
    console.log('Has <w:tbl>:', result.oxml.includes('<w:tbl>'));
    console.log('Has <w:ins:', result.oxml.includes('<w:ins'));

    // Save output for inspection
    await fs.writeFile(path.join(__dirname, 'test6_output.xml'), result.oxml);
    console.log('\nOutput saved to tests/test6_output.xml');
}

runDiagnostic().catch(console.error);

