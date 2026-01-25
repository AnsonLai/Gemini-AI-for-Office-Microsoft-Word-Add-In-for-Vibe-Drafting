/**
 * Redline Toggle Regression Test
 * 
 * Verifies that the redline setting is honored by applyRedlineToOxml.
 * 
 * Run with: node --experimental-modules tests/redline_toggle_test.mjs
 */

import { JSDOM } from 'jsdom';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/oxml-engine.js';

// --- Mock Browser Environment ---
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

async function runTests() {
    console.log('Starting Redline Toggle Regression Tests...');

    const originalOoxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:t>Hello world</w:t>
            </w:r>
        </w:p>
    `;
    const originalText = "Hello world";
    const modifiedText = "Hello Gemini";

    // 1. Test with generateRedlines: true (Expected: w:ins and w:del present)
    console.log('\n--- Test 1: Redlines ENABLED ---');
    const resultEnabled = await applyRedlineToOxml(originalOoxml, originalText, modifiedText, {
        author: 'TestUser',
        generateRedlines: true
    });

    const hasIns = resultEnabled.oxml.includes('<w:ins');
    const hasDel = resultEnabled.oxml.includes('<w:del');

    if (hasIns && hasDel) {
        console.log('✅ PASS: Track changes (w:ins/w:del) generated in redline mode');
    } else {
        console.log('❌ FAIL: Track changes NOT generated in redline mode');
        console.log('Output length:', resultEnabled.oxml.length);
        console.log('Output:', resultEnabled.oxml);
    }

    // 2. Test with generateRedlines: false (Expected: NO w:ins or w:del, text updated directly)
    console.log('\n--- Test 2: Redlines DISABLED ---');
    const resultDisabled = await applyRedlineToOxml(originalOoxml, originalText, modifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });

    const hasInsDisabled = resultDisabled.oxml.includes('<w:ins');
    const hasDelDisabled = resultDisabled.oxml.includes('<w:del');
    // Strip tags for text comparison
    const plainText = resultDisabled.oxml.replace(/<[^>]+>/g, '');
    const hasNewText = plainText.includes('Hello Gemini');

    console.log('Result OOXML length:', resultDisabled.oxml.length);
    console.log('Result OOXML contains <w:ins:', hasInsDisabled);
    console.log('Result OOXML contains <w:del:', hasDelDisabled);
    console.log('Cleaned text:', plainText);
    console.log('Result OOXML contains new text:', hasNewText);

    if (!hasInsDisabled && !hasDelDisabled && hasNewText) {
        console.log('✅ PASS: No track changes generated when disabled, text updated directly');
    } else {
        console.log('❌ FAIL: Redline toggle NOT honored');
        if (hasInsDisabled) console.log('- Found w:ins in output');
        if (hasDelDisabled) console.log('- Found w:del in output');
        if (!hasNewText) console.log('- New text NOT found in output');
        console.log('Output:', resultDisabled.oxml);
    }

    // 3. Test List Expansion with redlines disabled
    console.log('\n--- Test 3: List Expansion with Redlines DISABLED ---');
    const listModifiedText = "Original text\n* Item 1\n* Item 2";
    const resultListDisabled = await applyRedlineToOxml(originalOoxml, originalText, listModifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });

    const hasInsList = resultListDisabled.oxml.includes('<w:ins');
    const hasDelList = resultListDisabled.oxml.includes('<w:del');
    const hasBullets = resultListDisabled.oxml.includes('Item 1') && resultListDisabled.oxml.includes('Item 2');

    // In list expansion, we currently still generate w:numPr which is correct.
    // We want to ensure NO w:ins or w:del.

    if (!hasInsList && !hasDelList && hasBullets) {
        console.log('✅ PASS: List expansion honors redline toggle');
    } else {
        console.log('❌ FAIL: List expansion redline toggle issue');
        if (hasInsList) console.log('- Found w:ins in output');
        if (hasDelList) console.log('- Found w:del in output');
        console.log('Output:', resultListDisabled.oxml);
    }

    // 4. Test Table Reconciliation with redlines disabled
    console.log('\n--- Test 4: Table Reconciliation with Redlines DISABLED ---');
    const tableOriginalOoxml = `
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:tblGrid><w:gridCol/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
    `;
    const tableOriginalText = "Header 1";
    const tableModifiedText = "| Header Updated |\n| --- |";

    // This triggers applyTableReconciliation
    const resultTableDisabled = await applyRedlineToOxml(tableOriginalOoxml, tableOriginalText, tableModifiedText, {
        author: 'TestUser',
        generateRedlines: false
    });

    const hasInsTable = resultTableDisabled.oxml.includes('<w:ins');
    const hasDelTable = resultTableDisabled.oxml.includes('<w:del');
    const plainTextTable = resultTableDisabled.oxml.replace(/<[^>]+>/g, '');
    const hasUpdatedText = plainTextTable.includes('Header Updated');

    if (!hasInsTable && !hasDelTable && hasUpdatedText) {
        console.log('✅ PASS: Table reconciliation honors redline toggle');
    } else {
        console.log('❌ FAIL: Table reconciliation redline toggle issue');
        if (hasInsTable) console.log('- Found w:ins in output');
        if (hasDelTable) console.log('- Found w:del in output');
        if (!hasUpdatedText) {
            console.log('- Updated text NOT found in output');
            console.log('  Cleaned text:', plainTextTable);
        }
        console.log('Output:', resultTableDisabled.oxml);
    }
}

runTests().catch(console.error);
