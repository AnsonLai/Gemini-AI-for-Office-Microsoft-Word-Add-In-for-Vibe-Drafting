// Reproduces: a `format` operation cannot select a later occurrence of
// `textToFormat` inside a strongly targeted paragraph, because the only
// occurrence input (target.occurrence) is also consumed by paragraph targeting.
// Exits 0 while the defect reproduces against the installed package.
//
//   node docs/library-issues/completed/2026-10-01-format-occurrence-targeting-reproducer.mjs
import '../../../tests/setup-xml-provider.mjs';
import { inspectDocumentParts } from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const xml = `<w:document ${W}><w:body>`
  + '<w:p><w:r><w:t>Other paragraph.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t xml:space="preserve">laws of British Columbia, courts of </w:t></w:r>'
  + '<w:r><w:rPr><w:b/></w:rPr><w:t>British Columbia</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>'
  + '<w:sectPr/></w:body></w:document>';
const paragraph = inspectDocumentParts({ documentXml: xml }).paragraphs[1];
const target = { index: 2, exactText: paragraph.exactText, paragraphId: paragraph.paragraphId, fingerprint: paragraph.fingerprint };

// Expected: unbold the SECOND "British Columbia" (the bold one) in P2.
// Resolved in 0.8.4 by the dedicated `textOccurrence` field (target.occurrence
// still selects a paragraph, by design).
const result = await applyOperationsToDocumentXml(xml, [{
  type: 'format',
  target,
  textOccurrence: 2,
  textToFormat: 'British Columbia',
  properties: { bold: false }
}], 'Reproducer', null, { atomic: true, strictTargets: true, generateRedlines: true });

const formattedSecond = result.status === 'ok'
  && /laws of British Columbia, courts of <\/w:t>/.test(result.documentXml)
  && /<w:b w:val="0"\/>[\s\S]*?<w:t>British Columbia<\/w:t>/.test(result.documentXml);
const reproduced = !formattedSecond;
console.log(`format textOccurrence 2 in a strongly targeted paragraph: status=${result.status} error=${result.error?.code ?? 'none'}`);
console.log(reproduced ? 'REPRODUCED: later occurrences cannot be formatted.' : 'NOT REPRODUCED: later occurrence was formatted.');
process.exit(reproduced ? 0 : 1);
