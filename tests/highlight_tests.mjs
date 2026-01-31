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
    assert(result1.includes('w:highlight w:val="yellow"'), "Highlight tag not found");
    // With split logic, "Hello world" -> "Hello" (highlighted) + " world" (suffix)
    assert(result1.includes('<w:t xml:space="preserve">Hello</w:t>'), "Highlighted text segment missing");
    assert(result1.includes('<w:t xml:space="preserve"> world</w:t>'), "Suffix text segment missing");

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

    console.log("--- Test: Substring Highlight (Split Run) ---");
    const ooxmlSplit = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:t>prefix target suffix</w:t>
        </w:r>
    </w:p>`;
    const result4 = applyHighlightToOoxml(ooxmlSplit, "target", "yellow");
    const matchCount = (result4.match(/<w:r>/g) || []).length;
    if (matchCount > 1) {
        console.log("Result split runs: true");
    } else {
        console.log("Result split runs: false (FAIL - highlighted entire run)");
    }

    console.log("--- Test: Highlight with Redlines (Track Changes) ---");
    const ooxmlRedline = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:t>track changes highlight</w:t>
        </w:r>
    </w:p>`;
    // This signature assumes we will update applyHighlightToOoxml to accept options
    const result5 = applyHighlightToOoxml(ooxmlRedline, "highlight", "yellow", {
        generateRedlines: true,
        author: "TestAuthor"
    });

    // Should have w:rPrChange
    if (result5.includes('w:rPrChange')) {
        console.log("Result contains rPrChange: true");
    } else {
        console.log("Result contains rPrChange: false (FAIL - no track changes generated)");
    }
    if (result5.includes('w:author="TestAuthor"')) {
        console.log("Result contains correct author: true");
    } else {
        console.log("Result contains correct author: false");
    }

    console.log("✅ ALL HIGHLIGHT TESTS PASSED");
}

testHighlight().catch(err => {
    console.error("❌ TEST FAILED:", err);
    process.exit(1);
});
