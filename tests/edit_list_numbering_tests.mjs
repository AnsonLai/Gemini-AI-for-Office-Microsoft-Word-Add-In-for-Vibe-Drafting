// edit_list passes the document's real numbering to applyRedlineToOxml so a
// generated list never reuses an existing numbering instance (0.8.4 option).
import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { applyRedlineToOxml, buildListMarkdown, normalizeListItemsWithLevels } from '@ansonlai/docx-redline-js';
import { readPackageNumberingXml } from '../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const numberingXml = `<w:numbering xmlns:w="${W}">`
  + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const documentXml = `<w:document xmlns:w="${W}"><w:body><w:p>`
  + '<w:r><w:t>A. First recital.</w:t></w:r><w:r><w:br/><w:t>B. Second recital.</w:t></w:r>'
  + '</w:p></w:body></w:document>';
const part = (name, type, xml) => `<pkg:part pkg:name="${name}" pkg:contentType="${type}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
const flatOpc = `<?xml version="1.0" standalone="yes"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`
  + part('/_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
  + part('/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', documentXml)
  + part('/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', numberingXml)
  + '</pkg:package>';

// 1. The helper reads the numbering part from a Word package (and nothing else).
const read = readPackageNumberingXml(flatOpc);
assert.match(read, /<w:num w:numId="1"/);
assert.equal(readPackageNumberingXml(documentXml), null);
assert.equal(readPackageNumberingXml(null), null);

// 2. With the document numbering supplied, the generated list gets a fresh ID.
const listMarkdown = buildListMarkdown(normalizeListItemsWithLevels(['First recital.', 'Second recital.'], { indentSpaces: 4 }), 'numbered', 'upperAlpha');
const result = await applyRedlineToOxml(flatOpc, 'A. First recital.\u000bB. Second recital.', listMarkdown, {
  author: 'Editor', generateRedlines: true, explicitStructuredContent: true, numberingXml: read
});
assert.equal(result.status, 'ok', JSON.stringify(result.error));
assert.equal(result.hasChanges, true);
const listIds = [...result.oxml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map(match => match[1]);
assert.equal(listIds.length, 2, 'both recitals become list paragraphs');
assert.ok(listIds.every(id => id !== '1'), `generated list must not reuse existing numId 1 (got ${listIds})`);

console.log('edit_list_numbering_tests passed');
