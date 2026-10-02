// Reproduces: a localized replacement in a soft-break paragraph targeted only by
// index + exactText (no paragraphId/fingerprint) reports "ok" but tracks the
// WHOLE paragraph as deleted and inserts nothing. With the full strong target
// the same edit is correct. Exits 0 while the defect reproduces.
//
//   node docs/library-issues/2026-10-02-weak-target-soft-break-deletion-reproducer.mjs
import '../../tests/setup-xml-provider.mjs';
import { acceptTrackedChangesInOoxml, inspectDocumentParts } from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const documentXml = `<w:document ${W}><w:body><w:p>`
  + '<w:r><w:t>First recital.</w:t></w:r><w:r><w:br/><w:t>Parties desire a potential deal.</w:t></w:r>'
  + '</w:p><w:sectPr/></w:body></w:document>';
const paragraph = inspectDocumentParts({ documentXml }).paragraphs[0];

async function acceptedText(target) {
  const result = await applyOperationsToDocumentXml(documentXml, [{
    type: 'redline', target, structuredContent: false,
    replacements: [{ find: 'a potential deal', replace: 'project Titan' }]
  }], 'Reproducer', null, { atomic: true, generateRedlines: true });
  const accepted = acceptTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
  const text = [...accepted.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)]
    .map(match => match[0].replace(/<w:br\/>/g, '\n').replace(/<[^>]+>/g, '')).join('\n');
  return { status: result.status, text };
}

const weak = await acceptedText({ index: 1, exactText: paragraph.exactText });
const strong = await acceptedText({ index: 1, exactText: paragraph.exactText, paragraphId: paragraph.paragraphId, fingerprint: paragraph.fingerprint });
console.log(`weak target:   status=${weak.status} accepted=${JSON.stringify(weak.text)}`);
console.log(`strong target: status=${strong.status} accepted=${JSON.stringify(strong.text)}`);
const reproduced = weak.status === 'ok' && !weak.text.includes('Titan') && strong.text.includes('Titan');
console.log(reproduced ? 'REPRODUCED: weak target silently deletes the paragraph.' : 'NOT REPRODUCED.');
process.exit(reproduced ? 0 : 1);
