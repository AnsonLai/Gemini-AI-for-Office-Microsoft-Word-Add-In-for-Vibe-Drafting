// Reproduces: literal underscore runs ("_____" signature lines) are parsed as Markdown bold/italic
// and collapsed in the Accept All text. A table-reconciliation whose modified Markdown repeats a
// table's cells verbatim (a no-change reconcile) turns "_________________________" into "_".
// The same parser also collapses underscores in a plain paragraph rewrite with
// structuredContent: false. Exits 0 while the defect reproduces and 1 once fixed.
//
//   node docs/library-issues/2026-10-12-table-reconciliation-underscore-collapse-reproducer.mjs
import '../../tests/setup-xml-provider.mjs';
import { acceptTrackedChangesInOoxml, inspectDocumentParts } from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const line = '_'.repeat(25);
const cell = text => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = cells => `<w:tr>${cells.map(cell).join('')}</w:tr>`;
const documentXml = `<w:document ${W}><w:body>`
  + `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>${row(['DISCLOSING PARTY:', 'RECEIVING PARTY:'])}${row([line, line])}</w:tbl>`
  + `<w:p><w:r><w:t>Signed: ${line}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const source = inspectDocumentParts({ documentXml }).paragraphs;
const strongTarget = p => ({ index: p.index, exactText: p.exactText, paragraphId: p.paragraphId, fingerprint: p.fingerprint, inTable: p.inTable });

async function acceptedTexts(operation) {
  const result = await applyOperationsToDocumentXml(documentXml, [operation], 'Reproducer', null,
    { atomic: true, strictTargets: true, generateRedlines: true });
  if (result.status !== 'ok') throw new Error(`unexpected ${result.status}: ${JSON.stringify(result.error ?? result.results)}`);
  const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
  return inspectDocumentParts({ documentXml: accepted }).paragraphs.map(p => p.exactText);
}

// 1. No-change table reconcile: the modified Markdown is the table's own text.
const reconcile = await acceptedTexts({
  type: 'table-reconciliation', target: strongTarget(source[0]), structuredContent: true,
  modified: `| DISCLOSING PARTY: | RECEIVING PARTY: |\n| --- | --- |\n| ${line} | ${line} |`
});
// 2. Paragraph rewrite with structuredContent: false ("plain text") that only appends words.
const rewriteTarget = source.find(p => p.exactText.startsWith('Signed:'));
const rewrite = await acceptedTexts({
  type: 'redline', target: strongTarget(rewriteTarget), structuredContent: false,
  modified: `${rewriteTarget.exactText} (print name)`
});

const reconcileExpected = source.slice(0, 4).map(p => p.exactText);
const reconcileCollapsed = reconcile.slice(0, 4).join('|') !== reconcileExpected.join('|');
const rewriteCollapsed = !rewrite.at(-1).includes(line);
console.log(`table-reconciliation (no change) accepted cells: ${JSON.stringify(reconcile.slice(2, 4))} expected ${JSON.stringify(reconcileExpected.slice(2, 4))}`);
console.log(`paragraph rewrite (structuredContent:false) accepted: ${JSON.stringify(rewrite.at(-1))} expected ${JSON.stringify(`${rewriteTarget.exactText} (print name)`)}`);
const reproduced = reconcileCollapsed || rewriteCollapsed;
console.log(reproduced
  ? `REPRODUCED: literal underscores collapse in Accept All (table-reconciliation: ${reconcileCollapsed ? 'yes' : 'no'}, paragraph rewrite: ${rewriteCollapsed ? 'yes' : 'no'}).`
  : `NOT REPRODUCED (paragraph rewrite collapsed: ${rewriteCollapsed ? 'yes' : 'no'}).`);
process.exit(reproduced ? 0 : 1);
