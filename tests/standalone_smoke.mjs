import './setup-xml-provider.mjs';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import assert from 'assert';

import {
    applyRedlineToOxml,
    parseOoxml,
    ingestOoxml
} from '../src/taskpane/modules/reconciliation/standalone.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOC_PATH = path.join(__dirname, 'sample_doc/word/document.xml');

async function run() {
    const originalXml = await fs.readFile(DOC_PATH, 'utf8');
    const { acceptedText } = ingestOoxml(originalXml);
    assert(acceptedText && acceptedText.length > 0, 'Accepted text must not be empty');

    const modifiedText = `${acceptedText}\nStandalone smoke insertion.`;
    const result = await applyRedlineToOxml(originalXml, acceptedText, modifiedText, {
        author: 'StandaloneSmoke',
        generateRedlines: true
    });

    assert(result.hasChanges, 'Expected changes to be detected');
    assert(result.oxml && result.oxml.length > 0, 'Expected OOXML output');
    assert(result.oxml.includes('w:ins') || result.oxml.includes('w:del'), 'Expected track-change markers');

    const parsed = parseOoxml(result.oxml);
    const parseError = parsed.getElementsByTagName('parsererror')[0];
    assert(!parseError, `Output XML parse error: ${parseError?.textContent || ''}`);

    console.log('PASS: standalone smoke test');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
