import './setup-xml-provider.mjs';

import { JSDOM } from 'jsdom';
import { applyRedlineToOxml } from '../src/taskpane/modules/reconciliation/engine/oxml-engine.js';

// Global Setup
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;
global.Node = dom.window.Node;

async function testSurgicalPureFormats() {
    console.log('\n=== Test: Surgical Pure Formats (Addition & Removal) ===');

    const baseXml = (content) => `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            ${content}
        </w:p>
    `;

    // --- Helpers ---
    async function check(caseName, oxmlFragment, originalText, modifiedText, expectedChecks) {
        console.log(`\n--- ${caseName} ---`);
        const fullXml = baseXml(oxmlFragment);
        try {
            const result = await applyRedlineToOxml(fullXml, originalText, modifiedText, {
                author: 'Tester',
                generateRedlines: true
            });

            const oxml = result.oxml;
            const hasDel = oxml.includes('<w:del ') || oxml.includes('<w:del>');
            const hasIns = oxml.includes('<w:ins ') || oxml.includes('<w:ins>');
            const hasRPrChange = oxml.includes('<w:rPrChange ') || oxml.includes('<w:rPrChange>');

            if (!hasDel && !hasIns) {
                console.log("✅ PASS: No w:del/w:ins text replacements found.");
            } else {
                console.log("❌ FAIL: Found w:del/w:ins text replacements.");
            }

            if (hasRPrChange) {
                console.log("✅ PASS: Found w:rPrChange tracking.");
            } else {
                console.log("❌ FAIL: Missing w:rPrChange tracking.");
            }

            for (const check of expectedChecks) {
                if (oxml.includes(check.str)) {
                    console.log(`✅ PASS: Found expected string: ${check.desc}`);
                } else {
                    console.log(`❌ FAIL: Missing expected string: ${check.desc} ('${check.str}')`);
                }
            }
        } catch (e) {
            console.error(`ERROR in ${caseName}:`, e);
        }
    }

    // --- Test Cases ---

    // 1. ADD BOLD
    await check(
        "Add Bold",
        `<w:r><w:t>Hello</w:t></w:r>`,
        "Hello",
        "**Hello**",
        [{ str: 'w:b w:val="1"', desc: "Bold tag" }]
    );

    // 2. REMOVE BOLD
    await check(
        "Remove Bold",
        `<w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r>`,
        "Hello",
        "Hello", // Removed markdown -> unbold
        [{ str: 'w:b w:val="0"', desc: "Unbold tag" }]
    );

    // 3. ADD ITALIC
    await check(
        "Add Italic",
        `<w:r><w:t>Hello</w:t></w:r>`,
        "Hello",
        "*Hello*",
        [{ str: 'w:i w:val="1"', desc: "Italic tag" }]
    );

    // 4. REMOVE ITALIC
    await check(
        "Remove Italic",
        `<w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r>`,
        "Hello",
        "Hello",
        [{ str: 'w:i w:val="0"', desc: "Unitalic tag" }]
    );

    // 5. ADD STRIKETHROUGH
    await check(
        "Add Strikethrough",
        `<w:r><w:t>Hello</w:t></w:r>`,
        "Hello",
        "~~Hello~~",
        [{ str: 'w:strike w:val="1"', desc: "Strike tag" }]
    );

    // 6. REMOVE STRIKETHROUGH
    await check(
        "Remove Strikethrough",
        `<w:r><w:rPr><w:strike/></w:rPr><w:t>Hello</w:t></w:r>`,
        "Hello",
        "Hello",
        [{ str: 'w:strike w:val="0"', desc: "Unstrike tag" }]
    );

    // 7. ADD UNDERLINE (NOTE: Markdown doesn't have standard underline, assuming '++' or context specific, but checking generic format hint application if triggered)
    // Actually, our engine might map specific markdown to underline?
    // Let's assume input comes as format hints directly or we use a made-up syntax if supported.
    // Wait, standard markdown doesn't do underline.
    // I will check specific tag application logic by mocking the internal call if needed, but for now let's verify if '++' maps to underline?
    // Checking previous files... markdown-processor.js might have clues.
    // Assuming ++ is underline for now based on some common extensions or just testing the engine capability.

}

(async () => {
    await testSurgicalPureFormats();
})();

