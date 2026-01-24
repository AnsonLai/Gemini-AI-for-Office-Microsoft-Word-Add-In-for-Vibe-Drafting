
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
import { applyRedlineToOxml } from './src/taskpane/modules/reconciliation/oxml-engine.js';
import { ingestOoxml } from './src/taskpane/modules/reconciliation/ingestion.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, 'tests/sample_doc/word/document.xml');

async function runRepro() {
    console.log('--- Reproduction: Parties to Table ---');
    const originalOoxml = await fs.readFile(DOC_PATH, 'utf-8');
    const { acceptedText: originalText } = ingestOoxml(originalOoxml);

    // Construct new text with Markdown Table for Parties
    const start = originalText.indexOf('Disclosing Party:');
    const end = originalText.indexOf('RECITALS');

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

    // Check if the ORIGINAL table (Signature) is modified or if a NEW table is created?
    // Original table text: "DISCLOSING PARTY:"
    console.log('Result OXML contains "DISCLOSING PARTY:" (Original Table check):', result.oxml.includes('DISCLOSING PARTY:'));

    // Check if the NEW table text exists
    console.log('Result OXML contains new table content "[Name of Disclosing Party]":', result.oxml.includes('[Name of Disclosing Party]'));

    // Check if the OLD text "Disclosing Party:" (which was replaced) is gone (wrapped in w:del)
    console.log('Original Text "Disclosing Party:" deleted?', result.oxml.includes('<w:delText xml:space="preserve">Disclosing Party:</w:delText>'));

}

runRepro();
