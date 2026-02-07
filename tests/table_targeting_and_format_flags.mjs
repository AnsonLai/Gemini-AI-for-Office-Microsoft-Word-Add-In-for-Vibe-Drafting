import assert from 'assert';
import './setup-xml-provider.mjs';
import { createParser } from '../src/taskpane/modules/reconciliation/adapters/xml-adapter.js';
import { detectTableCellContext } from '../src/taskpane/modules/reconciliation/engine/table-cell-context.js';
import { extractFormatFromRPr } from '../src/taskpane/modules/reconciliation/engine/rpr-helpers.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

function testTableDuplicateTextTargetsByParaId() {
    const xml = `
<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p w14:paraId="AAA11111"><w:r><w:t>By: [Name]</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p w14:paraId="BBB22222"><w:r><w:t>By: [Name]</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`.trim();

    const parser = createParser();
    const doc = parser.parseFromString(xml, 'text/xml');
    const ctx = detectTableCellContext(doc, 'By: [Name]', { targetParagraphId: 'BBB22222' });

    assert.ok(ctx.targetParagraph, 'Expected a target paragraph');
    assert.strictEqual(ctx.targetParagraph.getAttribute('w14:paraId'), 'BBB22222');
}

function testFormatExtractionRespectsExplicitOffValues() {
    const xml = `
<w:rPr xmlns:w="${NS_W}">
  <w:b w:val="0"/>
  <w:i w:val="false"/>
  <w:u w:val="none"/>
  <w:strike w:val="off"/>
</w:rPr>`.trim();

    const parser = createParser();
    const doc = parser.parseFromString(xml, 'text/xml');
    const rPr = doc.documentElement;
    const format = extractFormatFromRPr(rPr);

    assert.strictEqual(format.bold, false);
    assert.strictEqual(format.italic, false);
    assert.strictEqual(format.underline, false);
    assert.strictEqual(format.strikethrough, false);
    assert.strictEqual(format.hasFormatting, false);
}

testTableDuplicateTextTargetsByParaId();
testFormatExtractionRespectsExplicitOffValues();
console.log('PASS: table targeting + format off-value handling');
