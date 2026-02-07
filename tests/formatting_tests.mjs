import './setup-xml-provider.mjs';

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/engine/oxml-engine.js';
import { wordsToChars, charsToWords } from '../src/taskpane/modules/reconciliation/pipeline/diff-engine.js';
import { preprocessMarkdown } from '../src/taskpane/modules/reconciliation/pipeline/markdown-processor.js';
import { diff_match_patch } from 'diff-match-patch';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// --- Test 1: Repro Formatting (from repro_format.mjs) ---
async function testReproFormatting() {
    console.log('\n=== Test: Repro Formatting Issues ===');

    const initialXmlPartial = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:p>
                <w:r>
                  <w:rPr>
                     <w:rStyle w:val="Normal"/>
                  </w:rPr>
                  <w:t>Hello World</w:t>
                </w:r>
              </w:p>
            </w:body>
          </w:document>
        </pkg:xmlData>
      </pkg:part>
    </pkg:package>`;

    const initialXmlOff = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:p>
                <w:r>
                  <w:rPr>
                     <w:b w:val="0"/>
                  </w:rPr>
                  <w:t>Hello World</w:t>
                </w:r>
              </w:p>
            </w:body>
          </w:document>
        </pkg:xmlData>
      </pkg:part>
    </pkg:package>`;

    console.log("--- Sub-test: Partial Formatting (Split + Order) ---");
    const partialText = "Hello **World**";
    try {
        const res = await applyRedlineToOxml(initialXmlPartial, "Hello World", partialText);

        if (!res.oxml.includes('<w:b/>') && !res.oxml.includes('<w:b>')) {
            console.log("❌ FAIL: Bold tag missing in partial update");
        } else {
            console.log("✅ PASS: Bold tag present.");
        }
    } catch (e) {
        console.error("Test Error:", e);
    }

    console.log("--- Sub-test: Overriding 'Off' Property ---");
    const boldText = "**Hello World**";
    try {
        const res = await applyRedlineToOxml(initialXmlOff, "Hello World", boldText);
        const matches = [...res.oxml.matchAll(/<w:b(?: [^>]*)?\/>/g)];
        const tags = matches.map(m => m[0]);
        const hasEnable = tags.some(t => !t.includes('w:val="0"') && !t.includes('w:val="false"'));

        if (hasEnable) {
            console.log("✅ PASS: Bold tag added/updated despite existing disable tag.");
        } else {
            console.log("❌ FAIL: Only found existing disable-bold tag.");
        }
    } catch (e) {
        console.error("Test Error:", e);
    }
}

// --- Test 2: Formatting Subtraction (from test_subtraction.mjs) ---
async function testFormattingSubtraction() {
    console.log('\n=== Test: Formatting Subtraction ===');

    console.log("--- Sub-test: Unbolding ---");
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

    if (result.oxml.includes('w:b w:val="0"') || result.oxml.includes('w:b val="0"')) {
        console.log("✅ PASS: Found explicit unbold (val=0)");
    } else {
        console.log("❌ FAIL: Explicit unbold not found");
    }

    console.log("--- Sub-test: Partial Subtraction (B+I -> I) ---");
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
        console.log("✅ PASS: Bold turned off, Italic kept on");
    } else {
        console.log("❌ FAIL: Incorrect partial subtraction");
    }
}

// --- Test 3: Markdown Processor (from test_processor.mjs) ---
function testMarkdownProcessor() {
    console.log('\n=== Test: Markdown Processor ===');

    const testCase1 = "~~NON-DISCLOSURE AGREEMENT~~";
    const res1 = preprocessMarkdown(testCase1);
    if (res1.cleanText === "NON-DISCLOSURE AGREEMENT") {
        console.log("✅ PASS: Strikethrough cleaned");
    } else {
        console.log("❌ FAIL: Strikethrough NOT cleaned");
    }

    const testCase4 = "++*NON-DISCLOSURE AGREEMENT*++";
    const res4 = preprocessMarkdown(testCase4);
    if (res4.cleanText === "NON-DISCLOSURE AGREEMENT" && res4.formatHints.length === 2) {
        console.log("✅ PASS: Nested formatting cleaned and hints captured");
    } else {
        console.log("❌ FAIL: Nested formatting issue persists");
    }
}

// --- Test 4: Diff Engine Granularity (from repro_diff.mjs) ---
function testDiffGranularity() {
    console.log('\n=== Test: Diff Granularity ===');
    const text1 = "British Columbia";
    const text2 = "the State of California";

    const dmp = new diff_match_patch();

    console.log('--- Word Level Tokenization Check ---');
    const { chars1, chars2, wordArray } = wordsToChars(text1, text2);
    const charDiffs = dmp.diff_main(chars1, chars2);
    dmp.diff_cleanupSemantic(charDiffs);
    const wordDiffs = charsToWords(charDiffs, wordArray);

    // We expect word level diffs, so "British" "Columbia" vs "the" "State" "of" "California"
    // Instead of char diffs like B-r-i-t...

    // Just checking it runs and returns array is a good start, but let's check length
    if (Array.isArray(wordDiffs) && wordDiffs.length > 0) {
        console.log('✅ PASS: Diff engine returned results.');
    } else {
        console.log('❌ FAIL: Diff engine returned empty or invalid.');
    }
}

// --- Main Runner ---
(async () => {
    console.log('STARTING FORMATTING TESTS...');

    await testReproFormatting();
    await testFormattingSubtraction();
    testMarkdownProcessor();
    testDiffGranularity();
    await testMiddleFormat();

    console.log('\nALL FORMATTING TESTS COMPLETE.');
})();

// --- Test 5: Middle Paragraph Formatting (from repro_middle_format.mjs) ---
async function testMiddleFormat() {
    console.log('\n=== Test: Middle Paragraph Formatting ===');

    const originalOxml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:t>The quick brown fox jumps.</w:t>
            </w:r>
        </w:p>
    `;
    const originalText = "The quick brown fox jumps.";
    const modifiedText = "The **quick** *brown* fox jumps.";

    const result = await applyRedlineToOxml(originalOxml, originalText, modifiedText, {
        author: 'Tester',
        generateRedlines: true
    });

    const hasBold = result.oxml.includes('<w:b/>') || (result.oxml.includes('<w:b ') && !result.oxml.includes('w:val="0"'));
    const hasItalic = result.oxml.includes('<w:i/>') || (result.oxml.includes('<w:i ') && !result.oxml.includes('w:val="0"'));

    if (hasBold && hasItalic) {
        console.log("✅ PASS: Both bold and italic applied.");
    } else {
        console.log("❌ FAIL: Formatting missing. Bold:", hasBold, "Italic:", hasItalic);
    }
}

