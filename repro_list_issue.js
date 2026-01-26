import { ReconciliationPipeline } from './src/taskpane/modules/reconciliation/pipeline.js';
import { NumberingService } from './src/taskpane/modules/reconciliation/numbering-service.js';
import { JSDOM } from 'jsdom';
import fs from 'fs';

// Setup environment for Node.js
const dom = new JSDOM('<!DOCTYPE html><html><body></body></html>');
global.window = dom.window;
global.document = dom.window.document;
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.Node = dom.window.Node;

async function testListExpansionWithSoftBreak() {
    const pipeline = new ReconciliationPipeline({
        generateRedlines: true,
        author: 'AI',
        numberingService: new NumberingService()
    });

    // Original paragraph with a soft break (<w:br/>)
    const originalOoxml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Phase 1 text.</w:t><w:br/><w:t>More text.</w:t></w:r></w:p>';
    const modifiedText = 'A. Point 1\nB. Point 2\nC. Point 3';

    console.log('--- Testing Expansion from Paragraph with Soft Break ---');
    const result = await pipeline.execute(originalOoxml, modifiedText);

    const pTags = (result.ooxml.match(/<w:p/g) || []);
    const pCount = pTags.length;

    let output = `Result p-count=${pCount}, isValid=${result.isValid}\n`;
    output += `--- RESULT OOXML ---\n${result.ooxml}`;

    fs.writeFileSync('repro_output.txt', output);
    console.log('Output written to repro_output.txt');

    if (pCount === 3) {
        console.log('SUCCESS: Correctly expanded into 3 paragraphs despite soft break in original.');
    } else {
        console.error('FAILED: Expected 3 paragraphs, got ' + pCount);
    }
}

testListExpansionWithSoftBreak().catch(console.error);
