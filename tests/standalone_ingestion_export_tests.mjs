import './setup-xml-provider.mjs';

import assert from 'assert';
import {
    ingestWordOoxmlToPlainText,
    ingestWordOoxmlToMarkdown
} from '../src/taskpane/modules/reconciliation/standalone.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function buildDocumentXml(bodyInnerXml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${bodyInnerXml}
    <w:sectPr/>
  </w:body>
</w:document>`;
}

function testPlainTextReadableStructure() {
    const xml = buildDocumentXml(`
      <w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t xml:space="preserve"> world</w:t></w:r></w:p>
      <w:p><w:r><w:t>Second line</w:t></w:r></w:p>
      <w:p><w:r><w:t>Third line</w:t></w:r></w:p>
    `);

    const result = ingestWordOoxmlToPlainText(xml);
    assert.strictEqual(typeof result, 'string');
    assert.strictEqual(result, 'Hello world\n\nSecond line\n\nThird line');
}

function testMarkdownHeadingsListsAndRunFormatting() {
    const xml = buildDocumentXml(`
      <w:p>
        <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
        <w:r><w:t>Overview</w:t></w:r>
      </w:p>
      <w:p>
        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
        <w:r><w:t>Bullet item</w:t></w:r>
      </w:p>
      <w:p>
        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
        <w:r><w:t>Numbered item</w:t></w:r>
      </w:p>
      <w:p>
        <w:r><w:t>This is </w:t></w:r>
        <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
        <w:r><w:t xml:space="preserve"> and </w:t></w:r>
        <w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>
      </w:p>
    `);

    const result = ingestWordOoxmlToMarkdown(xml);
    assert.strictEqual(typeof result, 'string');
    assert.ok(result.includes('## Overview'));
    assert.ok(result.includes('- Bullet item'));
    assert.ok(result.includes('1. Numbered item'));
    assert.ok(result.includes('This is **bold** and *italic*'));
}

function testMarkdownDetectsStrongRunStyleAsBold() {
    const xml = buildDocumentXml(`
      <w:p>
        <w:r>
          <w:rPr><w:rStyle w:val="Strong"/></w:rPr>
          <w:t>British Columbia</w:t>
        </w:r>
      </w:p>
    `);

    const result = ingestWordOoxmlToMarkdown(xml);
    assert.strictEqual(result, '**British Columbia**');
}

function testInvalidInputIsNonThrowing() {
    const badXml = '';

    const plain = ingestWordOoxmlToPlainText(badXml);
    assert.strictEqual(plain, '');

    const md = ingestWordOoxmlToMarkdown(badXml);
    assert.strictEqual(md, '');
}

function run() {
    testPlainTextReadableStructure();
    testMarkdownHeadingsListsAndRunFormatting();
    testMarkdownDetectsStrongRunStyleAsBold();
    testInvalidInputIsNonThrowing();
    console.log('PASS: standalone ingestion export tests');
}

try {
    run();
} catch (error) {
    console.error('FAIL:', error?.message || error);
    process.exit(1);
}
