/**
 * Verification Test: List Context Preservation and Nested Markers
 * 
 * Tests the logic for preserving list numbering when editing single items
 * and generating proper levels for hierarchical markers (1.1, 1.1.1).
 */

import { JSDOM } from 'jsdom';
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;

import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/oxml-engine.js';

async function runTests() {
    console.log('=== Running List Logic Verification ===\n');

    // Test 1: Preserve list context when editing a single item with plain text
    await testListContextPreservation();

    // Test 2: Surgical insertion of nested item (1.1.1)
    await testNestedItemInsertion();

    console.log('\n=== Verification Complete ===');
}

async function testListContextPreservation() {
    console.log('--- Test 1: Single Item Context Preservation ---');

    // Original OOXML for a list item (numId 1, ilvl 0)
    const originalOoxml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
            <w:numPr>
                <w:ilvl w:val="0"/>
                <w:numId w:val="1"/>
            </w:numPr>
        </w:pPr>
        <w:r><w:t>Verification of Compliance</w:t></w:r>
    </w:p>`.trim();

    const originalText = "Verification of Compliance";
    const newContent = "The retention of this copy must be legally required."; // Plain text, no markers

    const result = await applyRedlineToOxml(originalOoxml, originalText, newContent, {
        author: 'TestUser',
        generateRedlines: true
    });

    console.log('Applied changes to list item.');

    // Check if output still has numPr
    const hasNumPr = result.oxml.includes('<w:numId w:val="1"/>') && result.oxml.includes('<w:ilvl w:val="0"/>');
    console.log('Preserved numPr (numId=1, ilvl=0):', hasNumPr);

    // Check if text is correctly marked with track changes
    const hasDeletion = result.oxml.includes('<w:del');
    const hasInsertion = result.oxml.includes('<w:ins');
    console.log('Has track changes (del/ins):', hasDeletion && hasInsertion);

    if (hasNumPr && hasDeletion && hasInsertion) {
        console.log('✅ PASS: Context preserved and redlines applied.');
    } else {
        console.log('❌ FAIL: Context lost or redlines missing.');
    }
}

async function testNestedItemInsertion() {
    console.log('\n--- Test 2: Nested Item Generation (1.1.1) ---');

    // Original OOXML for a list (item 1.1)
    const originalOoxml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
            <w:numPr>
                <w:ilvl w:val="1"/>
                <w:numId w:val="1"/>
            </w:numPr>
        </w:pPr>
        <w:r><w:t>Item 1.1</w:t></w:r>
    </w:p>`.trim();

    const originalText = "Item 1.1";
    // Adding a sub-item after this one via surgical insertion (newline in content)
    const newContent = "Item 1.1\n  1.1.1. New nested item";

    const result = await applyRedlineToOxml(originalOoxml, originalText, newContent, {
        author: 'TestUser',
        generateRedlines: false // Simple check for structure
    });

    // Check if second paragraph was created with ilvl 2
    const paragraphs = result.oxml.match(/<w:p[\s\S]*?<\/w:p>/g);
    console.log('Resulting paragraphs:', paragraphs?.length);

    if (paragraphs && paragraphs.length >= 2) {
        const p2 = paragraphs[1];
        const hasIlvl2 = p2.includes('w:ilvl w:val="2"');
        console.log('Second paragraph OOXML snippet:', p2.substring(0, 150));
        console.log('Second paragraph has ilvl 2:', hasIlvl2);

        if (hasIlvl2) {
            console.log('✅ PASS: Correct ilvl generated for nested marker.');
        } else {
            console.log('❌ FAIL: Incorrect ilvl for nested marker.');
        }
    } else {
        console.log('❌ FAIL: Second paragraph not generated.');
    }
}

runTests().catch(console.error);
