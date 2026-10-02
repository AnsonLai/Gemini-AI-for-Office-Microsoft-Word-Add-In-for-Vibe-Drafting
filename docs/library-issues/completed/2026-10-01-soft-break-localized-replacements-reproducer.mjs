// Reproduces: localized `replacements` are refused (INVALID_OPERATION) for any
// paragraph containing a soft line break (w:br), even when the find/replace
// spans text on a single line. Exits 0 while the defect reproduces.
//
//   node docs/library-issues/completed/2026-10-01-soft-break-localized-replacements-reproducer.mjs
import '../../../tests/setup-xml-provider.mjs';
import { inspectDocumentParts } from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const documentXml = `<w:document ${W}><w:body><w:p>`
  + '<w:r><w:t>A. First recital.</w:t></w:r>'
  + '<w:r><w:br/><w:t>B. The Parties desire a potential business relationship.</w:t></w:r>'
  + '</w:p><w:sectPr/></w:body></w:document>';
const paragraph = inspectDocumentParts({ documentXml }).paragraphs[0];

const result = await applyOperationsToDocumentXml(documentXml, [{
  type: 'redline',
  target: { index: 1, exactText: paragraph.exactText, paragraphId: paragraph.paragraphId, fingerprint: paragraph.fingerprint },
  replacements: [{ find: 'a potential business relationship', replace: 'project Titan' }]
}], 'Reproducer', null, { atomic: true, generateRedlines: true });

const code = result.results?.[0]?.error?.code ?? result.error?.code;
const reproduced = result.status === 'error' && code === 'INVALID_OPERATION';
console.log(`exactText has a line break: ${/\n/.test(paragraph.exactText)}; status=${result.status}; code=${code ?? 'none'}`);
console.log(reproduced ? 'REPRODUCED: replacements refused in a soft-break paragraph.' : 'NOT REPRODUCED: replacement applied.');
process.exit(reproduced ? 0 : 1);
