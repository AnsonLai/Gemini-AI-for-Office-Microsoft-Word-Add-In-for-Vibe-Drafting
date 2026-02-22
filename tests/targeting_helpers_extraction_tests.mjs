import './setup-xml-provider.mjs';

import assert from 'assert';

import {
    resolveParagraphRangeByRefs as resolveParagraphRangeByRefsStandalone,
    inferTableReplacementParagraphBlock as inferTableReplacementParagraphBlockStandalone,
    isLikelyStructuredTableSourceParagraph as isLikelyStructuredTableSourceParagraphStandalone,
    parseOoxml
} from '../src/taskpane/modules/reconciliation/standalone.js';

import {
    resolveParagraphRangeByRefs
} from '../src/taskpane/modules/reconciliation/core/paragraph-targeting.js';

import {
    inferTableReplacementParagraphBlock,
    isLikelyStructuredTableSourceParagraph
} from '../src/taskpane/modules/reconciliation/core/table-targeting.js';

function testStandaloneReexportsMatchCore() {
    assert.strictEqual(resolveParagraphRangeByRefsStandalone, resolveParagraphRangeByRefs);
    assert.strictEqual(inferTableReplacementParagraphBlockStandalone, inferTableReplacementParagraphBlock);
    assert.strictEqual(isLikelyStructuredTableSourceParagraphStandalone, isLikelyStructuredTableSourceParagraph);
}

function testParagraphRangeResolution() {
    const xml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>One</w:t></w:r></w:p>
    <w:p><w:r><w:t>Two</w:t></w:r></w:p>
    <w:p><w:r><w:t>Three</w:t></w:r></w:p>
  </w:body>
</w:document>`.trim();
    const doc = parseOoxml(xml);
    const range = resolveParagraphRangeByRefs(doc, 'P1', 'P3', {});
    assert.ok(Array.isArray(range));
    assert.strictEqual(range.length, 3);
}

function testTableHeuristicHelpers() {
    assert.strictEqual(isLikelyStructuredTableSourceParagraph('and'), true);
    assert.strictEqual(isLikelyStructuredTableSourceParagraph('This is a full sentence.'), false);

    const xml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Title:</w:t></w:r></w:p>
    <w:p><w:r><w:t>[Name]</w:t></w:r></w:p>
    <w:p><w:r><w:t>Done.</w:t></w:r></w:p>
  </w:body>
</w:document>`.trim();
    const doc = parseOoxml(xml);
    const paragraphs = Array.from(doc.getElementsByTagNameNS('*', 'p'));
    const block = inferTableReplacementParagraphBlock(paragraphs[0], {});
    assert.ok(Array.isArray(block));
    assert.strictEqual(block.length, 2);
}

function run() {
    testStandaloneReexportsMatchCore();
    testParagraphRangeResolution();
    testTableHeuristicHelpers();
    console.log('PASS: targeting helpers extraction');
}

try {
    run();
} catch (err) {
    console.error('FAIL:', err.message);
    process.exit(1);
}
