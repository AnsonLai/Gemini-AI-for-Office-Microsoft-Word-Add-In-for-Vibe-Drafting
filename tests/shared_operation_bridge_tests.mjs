import './setup-xml-provider.mjs';

import assert from 'assert';
import {
    applySharedOperationToParagraphOoxml,
    applySharedOperationToScopeOoxml
} from '../src/taskpane/modules/commands/shared-operation-bridge.js';
import {
    extractReplacementNodesFromOoxml,
    getParagraphText
} from '../src/taskpane/modules/reconciliation/standalone.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function buildParagraphXml(text) {
    return `<w:p xmlns:w="${NS_W}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function getDirectWordChild(element, localName) {
    if (!element) return null;
    return Array.from(element.childNodes || []).find(
        node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === localName
    ) || null;
}

async function testHighlightBridge() {
    const paragraphXml = buildParagraphXml('Alpha target span.');
    const result = await applySharedOperationToParagraphOoxml(
        paragraphXml,
        {
            type: 'highlight',
            targetRef: 'P1',
            target: 'Alpha target span.',
            textToHighlight: 'target',
            color: 'yellow'
        },
        {
            author: 'BridgeTest',
            generateRedlines: false
        }
    );

    assert.strictEqual(result.hasChanges, true, 'highlight bridge should report changes');
    assert.ok(result.paragraphOoxml.includes('w:highlight'), 'highlight bridge should return highlighted paragraph OOXML');
    assert.ok(result.packageOoxml && result.packageOoxml.includes('pkg:package'), 'highlight bridge should return package OOXML for Word insertion');
    assert.strictEqual(result.commentsXml, null, 'highlight bridge should not emit comments XML');
}

async function testCommentBridge() {
    const paragraphXml = buildParagraphXml('Comment target span.');
    const result = await applySharedOperationToParagraphOoxml(
        paragraphXml,
        {
            type: 'comment',
            targetRef: 'P1',
            target: 'Comment target span.',
            textToComment: 'target',
            commentContent: 'Please verify this term.'
        },
        {
            author: 'BridgeTest',
            generateRedlines: true
        }
    );

    assert.strictEqual(result.hasChanges, true, 'comment bridge should report changes');
    assert.ok(result.paragraphOoxml.includes('commentRangeStart'), 'comment bridge should return paragraph OOXML with comment markers');
    assert.ok(result.commentsXml && result.commentsXml.includes('Please verify this term.'), 'comment bridge should emit comments XML');
    assert.ok(result.packageOoxml && result.packageOoxml.includes('/word/comments.xml'), 'comment bridge should return package OOXML for Word insertion');
}

async function testRedlineScopeBridge() {
    const scopeXml = `<w:document xmlns:w="${NS_W}"><w:body>${buildParagraphXml('Clause one.').replace(` xmlns:w="${NS_W}"`, '')}${buildParagraphXml('Clause two.').replace(` xmlns:w="${NS_W}"`, '')}</w:body></w:document>`;
    const result = await applySharedOperationToScopeOoxml(
        scopeXml,
        {
            type: 'redline',
            targetRef: 'P1',
            targetEndRef: 'P2',
            target: 'Clause one.',
            modified: 'Clause one updated.\nClause two updated.'
        },
        {
            author: 'BridgeTest',
            generateRedlines: false
        }
    );

    assert.strictEqual(result.hasChanges, true, 'scope redline bridge should report changes');
    assert.ok(result.packageOoxml && result.packageOoxml.includes('pkg:package'), 'scope redline bridge should return package OOXML');
    const extracted = extractReplacementNodesFromOoxml(result.packageOoxml);
    const paragraphTexts = (extracted.replacementNodes || []).map(node => getParagraphText(node).trim()).filter(Boolean);
    assert.ok(paragraphTexts.includes('Clause one updated.'), 'scope redline bridge package should contain updated scope text');
    assert.ok(paragraphTexts.includes('Clause two updated.'), 'scope redline bridge package should contain second paragraph update');
}

async function testRedlineScopeBridgePreservesTableNodes() {
    const scopeXml = `<w:document xmlns:w="${NS_W}"><w:body>${buildParagraphXml('Party block start').replace(` xmlns:w="${NS_W}"`, '')}${buildParagraphXml('Disclosing Party details').replace(` xmlns:w="${NS_W}"`, '')}${buildParagraphXml('Receiving Party details').replace(` xmlns:w="${NS_W}"`, '')}</w:body></w:document>`;
    const result = await applySharedOperationToScopeOoxml(
        scopeXml,
        {
            type: 'redline',
            targetRef: 'P1',
            targetEndRef: 'P3',
            target: 'Party block start',
            modified: '| Role | Details |\n|---|---|\n| Disclosing Party | [Name] |\n| Receiving Party | [Name] |'
        },
        {
            author: 'BridgeTest',
            generateRedlines: false
        }
    );

    assert.strictEqual(result.hasChanges, true, 'scope redline bridge table transform should report changes');
    assert.ok(result.packageOoxml && result.packageOoxml.includes('pkg:package'), 'scope redline bridge table transform should return package OOXML');

    const extracted = extractReplacementNodesFromOoxml(result.packageOoxml);
    const topLevelNames = (extracted.replacementNodes || []).map(node => node.localName);
    assert.ok(topLevelNames.includes('tbl'), 'scope redline bridge package should preserve a top-level table node');
}

async function testParagraphRedlineBridgePreservesDirectListBinding() {
    const paragraphXml = `
<w:p xmlns:w="${NS_W}">
  <w:pPr>
    <w:pStyle w:val="ListParagraph"/>
    <w:pPrChange w:id="10" w:author="Prior" w:date="2026-01-01T00:00:00Z">
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="1"/>
          <w:numId w:val="99"/>
        </w:numPr>
      </w:pPr>
    </w:pPrChange>
  </w:pPr>
  <w:r><w:t>The Receiving Party's legal counsel may retain one (1) copy.</w:t></w:r>
</w:p>`.trim();

    const result = await applySharedOperationToParagraphOoxml(
        paragraphXml,
        {
            type: 'redline',
            targetRef: 'P1',
            target: "The Receiving Party's legal counsel may retain one (1) copy.",
            modified: "The Receiving Party's legal counsel may retain two (2) copies."
        },
        {
            author: 'BridgeTest',
            generateRedlines: true
        }
    );

    assert.strictEqual(result.hasChanges, true, 'paragraph redline bridge should report changes');
    const doc = new DOMParser().parseFromString(result.paragraphOoxml, 'application/xml');
    const paragraph = doc.getElementsByTagNameNS(NS_W, 'p')[0];
    assert.ok(paragraph, 'paragraph redline bridge should return paragraph OOXML');
    const pPr = getDirectWordChild(paragraph, 'pPr');
    assert.ok(pPr, 'output paragraph should include pPr');
    const numPr = getDirectWordChild(pPr, 'numPr');
    assert.ok(numPr, 'output paragraph should contain direct w:numPr list binding');
    const ilvl = getDirectWordChild(numPr, 'ilvl');
    const numId = getDirectWordChild(numPr, 'numId');
    assert.strictEqual(ilvl?.getAttribute('w:val') || ilvl?.getAttribute('val'), '1', 'output list binding should preserve ilvl');
    assert.strictEqual(numId?.getAttribute('w:val') || numId?.getAttribute('val'), '99', 'output list binding should preserve numId');
    const pPrChange = getDirectWordChild(pPr, 'pPrChange');
    assert.strictEqual(pPrChange, null, 'output paragraph should remove stale pPrChange list metadata');
}

async function run() {
    await testHighlightBridge();
    await testCommentBridge();
    await testRedlineScopeBridge();
    await testRedlineScopeBridgePreservesTableNodes();
    await testParagraphRedlineBridgePreservesDirectListBinding();
    console.log('PASS: shared operation bridge tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
