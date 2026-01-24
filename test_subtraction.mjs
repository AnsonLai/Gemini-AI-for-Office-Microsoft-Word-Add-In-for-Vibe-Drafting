import { applyRedlineToOxml } from './src/taskpane/modules/reconciliation/oxml-engine.js';
import { JSDOM } from 'jsdom';

// Mock DOMParser and XMLSerializer for Node.js
const dom = new JSDOM();
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;

async function runTests() {
    console.log("--- Test: Formatting Subtraction (Unbolding) ---");
    const originalOxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:rPr><w:b/></w:rPr>
                <w:t>Bold Text</w:t>
            </w:r>
        </w:p>
    `;
    const originalText = "Bold Text";
    const modifiedText = "Bold Text"; // Removed markdown markers -> should unbold

    const result = await applyRedlineToOxml(originalOxml, originalText, modifiedText, {
        author: 'Tester',
        generateRedlines: true
    });

    console.log("Has Changes:", result.hasChanges);
    if (result.oxml.includes('w:b w:val="0"') || result.oxml.includes('w:b val="0"')) {
        console.log("PASS: Found explicit unbold (val=0)");
    } else {
        console.log("FAIL: Explicit unbold not found");
        console.log("Result OOXML:", result.oxml);
    }

    if (result.oxml.includes('w:rPrChange')) {
        console.log("PASS: Found rPrChange redline");
    } else {
        console.log("FAIL: rPrChange redline not found");
    }

    console.log("\n--- Test: Partial Subtraction (B+I -> I) ---");
    const originalOxml2 = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:rPr><w:b/><w:i/></w:rPr>
                <w:t>Bold Italic</w:t>
            </w:r>
        </w:p>
    `;
    const originalText2 = "Bold Italic";
    const modifiedText2 = "*Bold Italic*"; // Italic only -> should unbold

    const result2 = await applyRedlineToOxml(originalOxml2, originalText2, modifiedText2, {
        author: 'Tester',
        generateRedlines: true
    });

    if (result2.oxml.includes('w:b w:val="0"') && result2.oxml.includes('w:i') && !result2.oxml.includes('w:i w:val="0"')) {
        console.log("PASS: Bold turned off, Italic kept on");
    } else {
        console.log("FAIL: Incorrect partial subtraction");
        console.log("Result OOXML:", result2.oxml);
    }

    console.log("\n--- Test: Underline Removal (U -> Plain) ---");
    const originalOxml3 = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:rPr><w:u w:val="single"/></w:rPr>
                <w:t>Underlined</w:t>
            </w:r>
        </w:p>
    `;
    const originalText3 = "Underlined";
    const modifiedText3 = "Underlined";

    const result3 = await applyRedlineToOxml(originalOxml3, originalText3, modifiedText3, {
        author: 'Tester',
        generateRedlines: true
    });

    if (result3.oxml.includes('w:u w:val="none"') || result3.oxml.includes('w:u val="none"')) {
        console.log("PASS: Underline turned off (val=none)");
    } else {
        console.log("FAIL: Underline subtraction failed");
        console.log("Result OOXML:", result3.oxml);
    }
}

runTests().catch(console.error);
