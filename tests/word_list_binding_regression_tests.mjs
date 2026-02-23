import './setup-xml-provider.mjs';

import assert from 'assert';
import { applySharedOperationToParagraphOoxml } from '../src/taskpane/modules/reconciliation-integration/word-operation-runner.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function getDirectWordChild(element, localName) {
    if (!element) return null;
    return Array.from(element.childNodes || []).find(
        node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === localName
    ) || null;
}

async function testPprChangeOnlyListMetadataDoesNotForceDirectBinding() {
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
            author: 'RegressionTest',
            generateRedlines: true
        }
    );

    assert.strictEqual(result.hasChanges, true, 'operation should apply');
    const doc = new DOMParser().parseFromString(result.paragraphOoxml, 'application/xml');
    const paragraph = doc.getElementsByTagNameNS(NS_W, 'p')[0];
    assert.ok(paragraph, 'output paragraph should exist');

    const pPr = getDirectWordChild(paragraph, 'pPr');
    assert.ok(pPr, 'output should contain pPr');
    const directNumPr = getDirectWordChild(pPr, 'numPr');
    assert.strictEqual(
        Boolean(directNumPr),
        false,
        'should not force direct numPr when source had list metadata only inside pPrChange'
    );
}

async function run() {
    await testPprChangeOnlyListMetadataDoesNotForceDirectBinding();
    console.log('PASS: word list binding regression tests');
}

run().catch(error => {
    console.error('FAIL:', error?.message || error);
    process.exit(1);
});

