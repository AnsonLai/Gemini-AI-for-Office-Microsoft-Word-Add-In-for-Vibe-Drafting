// Reproduces the two installed-package defects behind the "turn the recitals into
// a proper ordered list? A, B, C..." failure. Exits 0 while BOTH defects reproduce
// against the installed @ansonlai/docx-redline-js, and 1 once either is fixed.
//
//   node docs/library-issues/completed/2026-10-01-soft-break-list-conversion-reproducer.mjs
import '../../../tests/setup-xml-provider.mjs';
import { applyRedlineToOxml, buildListMarkdown, normalizeListItemsWithLevels } from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const LETTERED = ['A. First recital.', 'B. Second recital.', 'C. Third recital.'];
const listMarkdown = buildListMarkdown(
  normalizeListItemsWithLevels(LETTERED.map(line => line.slice(3)), { indentSpaces: 4 }),
  'numbered',
  'upperAlpha'
);
// One paragraph whose items are separated by soft line breaks (Shift+Enter).
const softBreakXml = `<w:document ${W}><w:body><w:p>`
  + `<w:r><w:t>${LETTERED[0]}</w:t></w:r>`
  + `<w:r><w:br/><w:t>${LETTERED[1]}</w:t></w:r>`
  + `<w:r><w:br/><w:t>${LETTERED[2]}</w:t></w:r>`
  + `</w:p><w:sectPr/></w:body></w:document>`;
const options = { author: 'Reproducer', generateRedlines: true };

console.log('Generated list Markdown equals the manual source text:', listMarkdown === LETTERED.join('\n'));

// Defect 1: Word's Paragraph.text reports w:br as "\v"; the target check reads
// only w:t text ("...recital.B. Second...") and refuses the edit.
const fromWordText = await applyRedlineToOxml(softBreakXml, LETTERED.join('\u000b'), listMarkdown, {
  ...options,
  explicitStructuredContent: true
});
const defect1 = fromWordText.error?.code === 'TARGET_NOT_FOUND';
console.log(`1. "\\v"-separated host text: status=${fromWordText.status} error=${fromWordText.error?.code ?? 'none'}`);

// Defect 2: identical manual markers + explicit structured content is a no-op,
// on both the OOXML range entrypoint and the document-operation runner.
const explicitRange = await applyRedlineToOxml(softBreakXml, LETTERED.join('\n'), listMarkdown, {
  ...options,
  explicitStructuredContent: true
});
const runner = await applyOperationsToDocumentXml(softBreakXml, [{
  type: 'redline',
  target: { index: 1, exactText: LETTERED.join('\n') },
  modified: listMarkdown,
  structuredContent: true
}], 'Reproducer', null, { atomic: true, generateRedlines: true, structuredContent: true });
const defect2 = explicitRange.hasChanges === false && runner.hasChanges === false;
console.log(`2. explicit list request, identical text: range hasChanges=${explicitRange.hasChanges}; runner hasChanges=${runner.hasChanges} (${runner.results?.[0]?.status ?? runner.status})`);

console.log(defect1 && defect2 ? 'REPRODUCED: both defects present in the installed package.' : 'NOT REPRODUCED: at least one defect is fixed.');
process.exit(defect1 && defect2 ? 0 : 1);
