/**
 * Reproduction Test: List Conversion with Alpha Markers and Font Inheritance
 */

import { JSDOM } from 'jsdom';
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;

import { ReconciliationPipeline } from '../src/taskpane/modules/reconciliation/pipeline.js';
import fs from 'fs';
import path from 'path';

// Mimic the parseMarkdownList from taskpane.js since it's not exported
function parseMarkdownList(content) {
    if (!content) return null;

    const lines = content.trim().split('\n');
    const items = [];

    for (const line of lines) {
        if (!line.trim()) continue;

        // Unified marker regex matching NumberingService and Pipeline
        const markerRegex = /^(\s*)((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*•])\s*)(.*)$/;
        const match = line.match(markerRegex);

        if (match) {
            const indent = match[1];
            const marker = match[2].trim();
            const text = match[3];
            const level = Math.floor(indent.length / 2); // 2 spaces per level

            // Determine type based on marker
            const isBullet = /^[-*•]$/.test(marker);
            items.push({
                type: isBullet ? 'bullet' : 'numbered',
                level,
                text: text.trim(),
                marker: marker
            });
            continue;
        }

        // If line doesn't match list pattern, still include as text
        items.push({ type: 'text', level: 0, text: line.trim() });
    }

    if (items.length === 0) return null;

    // Determine primary type (numbered or bullet)
    const hasNumbered = items.some(i => i.type === 'numbered');
    const hasBullet = items.some(i => i.type === 'bullet');

    return {
        type: hasNumbered ? 'numbered' : (hasBullet ? 'bullet' : 'text'),
        items: items
    };
}

async function runTests() {
    console.log('=== Running List Conversion Verification ===\n');

    testParseMarkdownList();
    await testAlphaListConversion();
    await testFontInheritance();

    console.log('\n=== Verification Complete ===');
}

function testParseMarkdownList() {
    console.log('--- Test 0: parseMarkdownList Detection ---');
    const content = "A. Item 1\nB. Item 2";
    const listData = parseMarkdownList(content);

    console.log('Detected list type:', listData.type);
    if (listData.type === 'numbered') {
        console.log('✅ PASS: Alpha markers detected by parseMarkdownList.');
    } else {
        console.log('❌ FAIL: Alpha markers missed by parseMarkdownList (type: ' + listData.type + ').');
    }
}

async function testAlphaListConversion() {
    console.log('\n--- Test 1: Alpha List Conversion (A, B, C) ---');

    const pipeline = new ReconciliationPipeline({ generateRedlines: false });
    const content = "A. Item 1\nB. Item 2\nC. Item 3";

    // We mock the context which has no numbering
    const result = await pipeline.executeListGeneration(content, null, null, "Original Text");

    // Check if the generated OOXML has the correct numId or if the numberingXml contains upperLetter
    const hasUpperLetter = result.numberingXml && result.numberingXml.includes('w:numFmt w:val="upperLetter"');
    console.log('Generated numbering.xml includes upperLetter:', hasUpperLetter);

    if (hasUpperLetter) {
        console.log('✅ PASS: Alpha markers detected and numbering.xml updated.');
    } else {
        console.log('❌ FAIL: Alpha markers not properly mapped to numbering.xml.');
    }
}

async function testFontInheritance() {
    console.log('\n--- Test 2: Font Inheritance ---');

    const pipeline = new ReconciliationPipeline({
        generateRedlines: false,
        font: 'Calibri'
    });

    const content = "1. Item 1";

    try {
        const result = await pipeline.executeListGeneration(content, null, null, "Original Text");

        const hasFont = result.ooxml.includes('w:rFonts w:ascii="Calibri"');
        console.log('Generated OOXML includes Calibri font:', hasFont);

        if (hasFont) {
            console.log('✅ PASS: Font inherited correctly.');
        } else {
            console.log('❌ FAIL: Font not found in output OOXML.');
        }
    } catch (e) {
        console.log('❌ FAIL: Execution error:', e.message);
    }
}

runTests().catch(console.error);
