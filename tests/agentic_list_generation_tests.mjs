import assert from 'node:assert/strict';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml, configureXmlProvider } from '@ansonlai/docx-redline-js';

configureXmlProvider({ DOMParser, XMLSerializer });

const source = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Beta</w:t></w:r></w:p></w:body></w:document>';

for (const markdown of ['- Alpha\n- Beta', '1. Alpha\n2. Beta']) {
    const result = await applyRedlineToOxml(source, 'Alpha\nBeta', markdown, {
        author: 'List Compatibility Test',
        generateRedlines: true
    });
    assert.equal(result.status, 'ok', JSON.stringify(result.error));
    assert.equal(result.hasChanges, true, `Expected structural list change for ${markdown}`);
    assert.match(result.oxml, /<w:numPr\b/, `Expected native Word numbering for ${markdown}`);
    assert.match(result.oxml, /<w:ins\b/, `Expected tracked list insertion for ${markdown}`);
}

console.log('PASS: package list generation covers same-text structural conversion');
