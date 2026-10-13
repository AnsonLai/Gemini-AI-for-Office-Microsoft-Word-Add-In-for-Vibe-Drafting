import './setup-xml-provider.mjs';

import assert from 'assert';
import { parseXml, serializeXml } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';

// The default test lane configures no XML provider and installs no DOM globals:
// the library must parse and serialize on its own @xmldom/xmldom fallback in Node.
function run() {
    if ((process.env.DOCX_TEST_XML_PROVIDER || 'xmldom').toLowerCase() === 'xmldom') {
        assert.equal(globalThis.DOMParser, undefined, 'default lane must not install a global DOMParser');
        assert.equal(globalThis.document, undefined, 'default lane must not install a global document');
    }

    const doc = parseXml('<root><item>A</item><item>B</item></root>');
    const texts = [];
    for (const node of doc.getElementsByTagName('item')) texts.push(node.textContent);
    assert.deepStrictEqual(texts, ['A', 'B'], 'Expected an iterable element list with 2 items');
    assert.match(serializeXml(doc), /<root><item>A<\/item><item>B<\/item><\/root>/);
    console.log('PASS: library XML fallback smoke');
}

try {
    run();
} catch (err) {
    console.error('FAIL:', err.message);
    process.exit(1);
}
