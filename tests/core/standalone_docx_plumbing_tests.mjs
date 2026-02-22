import '../setup-xml-provider.mjs';

import assert from 'assert';
import {
    extractReplacementNodesFromOoxml,
    normalizeBodySectionOrderStandalone,
    sanitizeNestedParagraphsInTables,
    ensureNumberingArtifactsInZip,
    ensureCommentsArtifactsInZip,
    validateDocxPackage
} from '../../src/taskpane/modules/reconciliation/index.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const BASE_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${NS_CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const BASE_DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_RELS}"></Relationships>`;

const BASE_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

function parseXmlStrict(xml, label) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');
    const parseError = doc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`Parse failure (${label}): ${parseError.textContent || 'Unknown'}`);
    }
    return doc;
}

async function createBaseZip() {
    const zip = createMemoryZip();
    zip.file('word/document.xml', BASE_DOCUMENT_XML);
    zip.file('[Content_Types].xml', BASE_CONTENT_TYPES);
    zip.file('word/_rels/document.xml.rels', BASE_DOC_RELS);
    return zip;
}

function createMemoryZip() {
    const store = new Map();
    return {
        file(path, value) {
            if (typeof value === 'string') {
                store.set(path, value);
                return this;
            }
            if (!store.has(path)) return null;
            return {
                async(type) {
                    if (type !== 'string') {
                        throw new Error(`Unsupported async type for memory zip: ${type}`);
                    }
                    return Promise.resolve(store.get(path));
                }
            };
        }
    };
}

async function testExtractReplacementNodes() {
    const fragment = `<w:p xmlns:w="${NS_W}"><w:r><w:t>Fragment</w:t></w:r></w:p>`;
    const fragmentResult = extractReplacementNodesFromOoxml(fragment);
    assert.strictEqual(fragmentResult.replacementNodes.length, 1, 'fragment should yield one node');
    assert.strictEqual(fragmentResult.numberingXml, null, 'fragment numbering should be null');

    const documentXml = `<w:document xmlns:w="${NS_W}"><w:body><w:p><w:r><w:t>Doc</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
    const documentResult = extractReplacementNodesFromOoxml(documentXml);
    assert.strictEqual(documentResult.replacementNodes.length, 1, 'document should exclude sectPr');

    const packageXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" xmlns:w="${NS_W}">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document>
        <w:body>
          <w:p><w:r><w:t>Pkg</w:t></w:r></w:p>
          <w:sectPr/>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
    <pkg:xmlData>
      <w:numbering><w:abstractNum w:abstractNumId="1"/><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    const packageResult = extractReplacementNodesFromOoxml(packageXml);
    assert.strictEqual(packageResult.replacementNodes.length, 1, 'package should extract body children');
    assert.ok(packageResult.numberingXml && packageResult.numberingXml.includes('<w:numbering'), 'package should extract numberingXml');
}

function testNormalizeAndSanitize() {
    const doc = parseXmlStrict(
        `<w:document xmlns:w="${NS_W}"><w:body><w:p/><w:sectPr/><w:p/></w:body></w:document>`,
        'normalize'
    );
    normalizeBodySectionOrderStandalone(doc);
    const body = doc.getElementsByTagNameNS('*', 'body')[0];
    const children = Array.from(body.childNodes).filter(node => node.nodeType === 1);
    assert.strictEqual(children[children.length - 1].localName, 'sectPr', 'sectPr must be last body element');

    const nested = parseXmlStrict(
        `<w:document xmlns:w="${NS_W}"><w:body><w:tbl><w:tr><w:tc><w:p><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
        'sanitize'
    );
    const fixed = sanitizeNestedParagraphsInTables(nested);
    assert.ok(fixed > 0, 'sanitize should report fixed nested paragraph(s)');
    const tc = nested.getElementsByTagNameNS('*', 'tc')[0];
    const directParagraphs = Array.from(tc.childNodes).filter(
        node => node.nodeType === 1 && node.localName === 'p'
    );
    assert.ok(directParagraphs.length >= 2, 'inner paragraph should be promoted to tc child');
}

async function testArtifactsAndValidation() {
    const zip = await createBaseZip();
    const numberingXml = `<w:numbering xmlns:w="${NS_W}"><w:abstractNum w:abstractNumId="10"/><w:num w:numId="10"><w:abstractNumId w:val="10"/></w:num></w:numbering>`;
    const commentsXml1 = `<w:comments xmlns:w="${NS_W}"><w:comment w:id="1"><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:comment></w:comments>`;
    const commentsXml2 = `<w:comments xmlns:w="${NS_W}"><w:comment w:id="2"><w:p><w:r><w:t>c2</w:t></w:r></w:p></w:comment></w:comments>`;

    await ensureNumberingArtifactsInZip(zip, numberingXml);
    await ensureCommentsArtifactsInZip(zip, commentsXml1);
    await ensureCommentsArtifactsInZip(zip, commentsXml2);
    await validateDocxPackage(zip);

    const mergedComments = await zip.file('word/comments.xml')?.async('string');
    assert.ok(mergedComments && mergedComments.includes('w:id="1"') && mergedComments.includes('w:id="2"'), 'comments should merge by appending');
    const rels = await zip.file('word/_rels/document.xml.rels')?.async('string');
    assert.ok(rels && rels.includes('relationships/comments') && rels.includes('relationships/numbering'), 'document relationships should include comments and numbering');
}

async function run() {
    await testExtractReplacementNodes();
    testNormalizeAndSanitize();
    await testArtifactsAndValidation();
    console.log('PASS: standalone docx plumbing tests');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});


