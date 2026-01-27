/**
 * Repro Failures
 * Isolating the two failing cases from nda_redline_tests.mjs
 */

import { JSDOM } from 'jsdom';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/oxml-engine.js';
import { ingestOoxml } from '../src/taskpane/modules/reconciliation/ingestion.js';

// --- Mock Browser Environment ---
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

// --- Config ---
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

// --- Helpers ---
async function readSampleDoc() {
    return await fs.readFile(DOC_PATH, 'utf-8');
}

function extractTextSimple(ooxml) {
    const { acceptedText } = ingestOoxml(ooxml);
    return acceptedText;
}

// --- Tests ---
(async () => {
    console.log('=== Reproducing Redline Failures ===\n');

    const originalOoxml = await readSampleDoc();
    const originalText = extractTextSimple(originalOoxml);

    // 1. Change Archival Copies
    console.log('--- Test 1: Change Archival Copies ---');
    const text1 = originalText.replace('one (1) copy', 'two (2) copies');
    const result1 = await applyRedlineToOxml(originalOoxml, originalText, text1, {
        author: 'TestUser',
        generateRedlines: true
    });

    console.log('Has Changes:', result1.hasChanges);
    console.log('Contains <w:del>:', result1.oxml.includes('<w:del'));
    console.log('Contains <w:ins>:', result1.oxml.includes('<w:ins'));
    // Find the context
    const match1 = result1.oxml.match(/<w:p[^>]*>.*?legal counsel may retain.*?<\/w:p>/);
    if (match1) {
        console.log('Output Snippet:\n', match1[0]);
    } else {
        console.log('Output Snippet not found (looking for fallback context)...');
        const fallbackIndex = result1.oxml.indexOf('legal counsel may retain');
        console.log('Snippet around index:', result1.oxml.substring(fallbackIndex, fallbackIndex + 500));
    }

    // 2. Rewrite Disclosure Provision
    console.log('\n--- Test 2: Rewrite Disclosure Provision ---');
    const oldSection = 'If the Receiving Party is required by law, regulation, or court order to disclose any Confidential Information, the Receiving Party shall, to the extent legally permitted, provide the Disclosing Party with prompt written notice prior to such disclosure so that the Disclosing Party may seek a protective order or other appropriate remedy.';
    const newSection = 'The Receiving Party agrees to fight any orders to disclose Confidential Information. They will involve the Disclosing Party immediately. Only after all legal avenues are exhausted may they disclose the information.';

    if (!originalText.includes(oldSection)) {
        console.error('CRITICAL: Original text does not contain the old section to replace!');
    }

    const text2 = originalText.replace(oldSection, newSection);
    const result2 = await applyRedlineToOxml(originalOoxml, originalText, text2, {
        author: 'TestUser',
        generateRedlines: true
    });

    console.log('Has Changes:', result2.hasChanges);
    console.log('Contains <w:del>:', result2.oxml.includes('<w:del'));
    console.log('Contains "agrees to fight":', result2.oxml.includes('agrees to fight'));

    const match2 = result2.oxml.match(/<w:p[^>]*>.*?agrees to fight.*?<\/w:p>/);
    if (match2) {
        console.log('Output Snippet:\n', match2[0]);
    } else {
        console.log('Snippet not found (looking for fallback context)...');
        const fallbackIndex2 = result2.oxml.indexOf('agrees to fight');
        console.log('Snippet around index:', result2.oxml.substring(fallbackIndex2 - 100, fallbackIndex2 + 500));
    }

})();
