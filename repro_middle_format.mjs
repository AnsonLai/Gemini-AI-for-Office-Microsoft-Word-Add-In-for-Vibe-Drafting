import { applyRedlineToOxml } from './src/taskpane/modules/reconciliation/oxml-engine.js';
import { JSDOM } from 'jsdom';

const dom = new JSDOM();
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;

async function testMiddleFormat() {
    console.log("--- Test: Multiple Formatting Changes in Middle of Paragraph ---");

    // One run containing the whole text
    const originalOxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:t>The quick brown fox jumps.</w:t>
            </w:r>
        </w:p>
    `;
    const originalText = "The quick brown fox jumps.";

    // We want: The **quick** *brown* fox jumps.
    // "quick" is at index 4-9
    // "brown" is at index 10-15
    const modifiedText = "The **quick** *brown* fox jumps.";

    const result = await applyRedlineToOxml(originalOxml, originalText, modifiedText, {
        author: 'Tester',
        generateRedlines: true
    });

    console.log("Has Changes:", result.hasChanges);
    console.log("Result OOXML:\n", result.oxml);

    const hasBold = result.oxml.includes('<w:b/>') || (result.oxml.includes('<w:b ') && !result.oxml.includes('w:val="0"'));
    const hasItalic = result.oxml.includes('<w:i/>') || (result.oxml.includes('<w:i ') && !result.oxml.includes('w:val="0"'));

    console.log("Has Bold (quick):", hasBold);
    console.log("Has Italic (brown):", hasItalic);

    if (hasBold && hasItalic) {
        console.log("PASS: Both bold and italic applied.");
    } else if (hasBold) {
        console.log("FAIL: Only bold applied. Italic was likely swallowed by run split.");
    } else {
        console.log("FAIL: Neither formatting was applied.");
    }
}

testMiddleFormat().catch(console.error);
