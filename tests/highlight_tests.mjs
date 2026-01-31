import { JSDOM } from 'jsdom';
import { applyHighlightToOoxml } from '../src/taskpane/ooxml-formatting-removal.js';
import assert from 'assert';

// --- Global Setup ---
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;
global.Node = dom.window.Node;

async function testHighlight() {
    console.log("STARTING HIGHLIGHT TESTS...");

    const originalOoxml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:t>Hello world</w:t>
        </w:r>
    </w:p>`;

    console.log("--- Test: Apply Yellow Highlight ---");
    const result1 = applyHighlightToOoxml(originalOoxml, "Hello", "yellow");
    console.log("Result contains highlight:", result1.includes('w:highlight w:val="yellow"'));
    assert(result1.includes('w:highlight w:val="yellow"'), "Highlight tag not found");
    assert(result1.includes('<w:t>Hello world</w:t>'), "Text content corrupted");

    console.log("--- Test: Apply Green Highlight to Existing rPr ---");
    const ooxmlWithRPr = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>Hello world</w:t>
        </w:r>
    </w:p>`;
    const result2 = applyHighlightToOoxml(ooxmlWithRPr, "Hello", "green");
    console.log("Result contains highlight:", result2.includes('w:highlight w:val="green"'));
    console.log("Result contains bold:", result2.includes('<w:b/>'));
    assert(result2.includes('w:highlight w:val="green"'), "Green highlight not found");
    assert(result2.includes('<w:b/>'), "Bold property lost");

    console.log("--- Test: Case Insensitivity of Color ---");
    const result3 = applyHighlightToOoxml(originalOoxml, "Hello", "Cyan");
    assert(result3.includes('w:highlight w:val="cyan"'), "Cyan highlight (case insensitive) not found");

    console.log("✅ ALL HIGHLIGHT TESTS PASSED");
}

testHighlight().catch(err => {
    console.error("❌ TEST FAILED:", err);
    process.exit(1);
});
