import './setup-xml-provider.mjs';

import assert from 'assert';

import {
    createDynamicNumberingIdState as createDynamicNumberingIdStateStandalone,
    reserveNextNumberingIdPair as reserveNextNumberingIdPairStandalone,
    mergeNumberingXmlBySchemaOrder as mergeNumberingXmlBySchemaOrderStandalone
} from '@gsd/docx-reconciliation/index.js';

import {
    createDynamicNumberingIdState,
    reserveNextNumberingIdPair,
    mergeNumberingXmlBySchemaOrder
} from '@gsd/docx-reconciliation/services/numbering-helpers.js';

function testStandaloneReexportsMatchNumberingHelpers() {
    assert.strictEqual(createDynamicNumberingIdStateStandalone, createDynamicNumberingIdState);
    assert.strictEqual(reserveNextNumberingIdPairStandalone, reserveNextNumberingIdPair);
    assert.strictEqual(mergeNumberingXmlBySchemaOrderStandalone, mergeNumberingXmlBySchemaOrder);
}

function testCreateAndReserveIds() {
    const state = createDynamicNumberingIdState('', { minId: 7 });
    const pair = reserveNextNumberingIdPair(state);
    assert.deepStrictEqual(pair, { numId: 7, abstractNumId: 7 });
}

function testSchemaMergeKeepsOrderAndAddsNewNodes() {
    const existingXml = `
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1"/>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`.trim();

    const incomingXml = `
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="2"/>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`.trim();

    const merged = mergeNumberingXmlBySchemaOrder(existingXml, incomingXml);
    const indexAbstract2 = merged.indexOf('w:abstractNumId="2"');
    const indexNum2 = merged.indexOf('w:numId="2"');
    assert.ok(indexAbstract2 >= 0, 'Expected abstractNumId=2 to be present');
    assert.ok(indexNum2 >= 0, 'Expected numId=2 to be present');
    assert.ok(indexAbstract2 < indexNum2, 'Expected abstract numbering node before num node');
}

function run() {
    testStandaloneReexportsMatchNumberingHelpers();
    testCreateAndReserveIds();
    testSchemaMergeKeepsOrderAndAddsNewNodes();
    console.log('PASS: numbering helpers extraction');
}

try {
    run();
} catch (err) {
    console.error('FAIL:', err.message);
    process.exit(1);
}


