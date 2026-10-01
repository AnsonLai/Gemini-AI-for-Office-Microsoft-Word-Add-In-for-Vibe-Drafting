import '../../tests/setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { inspectDocumentParts } from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const documentXml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Header</w:t></w:r></w:p><w:sectPr/></w:body></w:document>';
const result = await applyOperationsToDocumentXml(
  documentXml,
  [{ type: 'redline', target: { index: 1, exactText: 'Header' }, modified: 'A. Header' }],
  'Standalone capability reproducer',
  {},
  { atomic: true, strictTargets: true, generateRedlines: true }
);
const inspected = inspectDocumentParts({ documentXml: result.documentXml });

assert.equal(result.status, 'ok', JSON.stringify(result.error));
assert.equal(result.hasChanges, true);
assert.equal(result.numberingXmlParts?.length ?? 0, 0);
assert.equal(inspected.paragraphs[0]?.text, 'A. Header');
assert.equal(inspected.paragraphs[0]?.list ?? null, null);

console.log(JSON.stringify({
  status: result.status,
  hasChanges: result.hasChanges,
  numberingParts: result.numberingXmlParts?.length ?? 0,
  acceptedText: inspected.paragraphs[0].text,
  listBinding: inspected.paragraphs[0].list ?? null
}, null, 2));
