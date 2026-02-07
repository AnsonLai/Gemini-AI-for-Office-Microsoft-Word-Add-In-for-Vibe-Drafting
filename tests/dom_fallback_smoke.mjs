import './setup-xml-provider.mjs';

import assert from 'assert';

function run() {
    assert(global.DOMParser, 'DOMParser should be available');
    assert(global.XMLSerializer, 'XMLSerializer should be available');
    assert(global.document, 'document should be available');

    const parser = new global.DOMParser();
    const doc = parser.parseFromString('<root><item>A</item><item>B</item></root>', 'text/xml');
    const items = doc.getElementsByTagName('item');

    // Ensure NodeList is iterable in both jsdom and xmldom fallback mode.
    let count = 0;
    for (const node of items) {
        count++;
        assert(node.textContent === 'A' || node.textContent === 'B', 'Unexpected node content');
    }

    assert.strictEqual(count, 2, 'Expected iterable NodeList with 2 items');
    console.log('PASS: dom fallback smoke');
}

try {
    run();
} catch (err) {
    console.error('FAIL:', err.message);
    process.exit(1);
}
