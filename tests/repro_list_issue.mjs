/**
 * Reproduction Test - List Numbering Issue
 * 
 * Recreates the user's scenario where adding a sub-bullet to "Archival Exception" 2.2
 * causes the numbering to reset/flatten instead of preserving hierarchy.
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

// --- Test Execution ---
(async () => {
    console.log('Running Reproduction Test: List Numbering Reset...');

    try {
        const originalOoxml = await readSampleDoc();
        const originalText = extractTextSimple(originalOoxml);

        // Simulate user's edit:
        // Existing text likely looks like:
        // 2. Archival Exception
        // 2.1 The Receiving Party's legal counsel may retain one (1) copy...
        // 2.2 This copy is to be used solely...
        //
        // User wants to add "2.2.1 ...legally required"

        // Let's identify the target section in original text to key off of
        const targetSentence = "This copy is to be used solely for archival purposes to ensure compliance and record-keeping.";

        if (!originalText.includes(targetSentence)) {
            console.error("❌ CRITICAL: Target sentence not found in sample_doc. Checking extracting content...");
            console.log("Original Text Preview:\n", originalText.substring(0, 500));
            return;
        }

        // Apply the edit: append the sub-bullet content
        // IMPORTANT: We use indentation to signal sub-bullet.
        // User wants 2.2.1.
        // If 2.2 is at level 1 (assuming 2. is level 0), then 2.2.1 should be level 2 (4 spaces or 2 tabs)
        // BUT, the user also mentioned "2.1.1" in the prompt "I was hoping to keep 2.1 and 2.2 while creating a 2.1.1"
        // Wait, the prompt says "subbullet in the archival exception after 2.2... creating a 2.1.1". 
        // 2.2 -> 2.2.1 makes sense. 2.2 -> 2.1.1 does not unless they meant 2.1.
        // Let's assume they want a child of the current item.

        // Create 'newText' by replacing the target with target + sub-bullet
        // We use 4-space indent for the sub-bullet
        const insertion = "\n    * The retention of such copy must be legally required.";
        const newText = originalText.replace(targetSentence, targetSentence + insertion);

        console.log("Original Text has newlines:", originalText.includes('\n'));
        console.log("New Text:", newText);

        const result = await applyRedlineToOxml(originalOoxml, originalText, newText, {
            author: 'TestUser',
            generateRedlines: true
        });


        console.log(`\nReconciliation Result (Valid: ${result.isValid})`);

        await fs.writeFile(path.join(__dirname, 'repro_output.xml'), result.oxml);
        console.log("Written output to tests/repro_output.xml");

        // Analyze the output XML around the change
        const index = result.oxml.indexOf('must be legally required');
        if (index === -1) {
            console.log("❌ FAIL: Inserted text not found in output.");
        } else {
            console.log("✅ Inserted text found.");
        }

    } catch (e) {
        console.error("💥 Error during test:", e);
    }
})();
