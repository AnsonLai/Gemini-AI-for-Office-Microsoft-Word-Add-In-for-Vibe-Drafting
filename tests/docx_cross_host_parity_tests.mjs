import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { configureLogger, openDocx, validateDocxPackage } from '@ansonlai/docx-redline-js';
import { MemoryZip, unzipDocx, zipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { createParser, createSerializer } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';
import { prepareCanonicalBatch } from '../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';
import { executePureOoxmlBatch } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { applyDocumentOperations } from '../mcp/docx-server/src/services/docx-document-service.mjs';
import { createDocumentSession } from '../browser-demo/document-session.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS = {
    document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
    numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
    comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
    commentsExtended: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
    commentsIds: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml',
    commentsExtensible: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml'
};
const author = 'Cross Host Reviewer';
const options = { author, generateRedlines: true };

configureLogger({ log: () => {}, warn: () => {}, error: () => {} }, { level: 'silent' });

function fixtureBytes() {
    const contentTypes = `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${NS.document}"/><Override PartName="/word/styles.xml" ContentType="${NS.styles}"/><Override PartName="/word/numbering.xml" ContentType="${NS.numbering}"/></Types>`;
    const rootRels = `<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`;
    const documentRels = `<Relationships xmlns="${PR}"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdNumbering" Type="${R}/numbering" Target="numbering.xml"/><Relationship Id="rIdHyperlink" Type="${R}/hyperlink" Target="https://example.org/review" TargetMode="External"/></Relationships>`;
    const documentXml = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="A1B2C3D4"><w:pPr><w:pStyle w:val="ReviewList"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:hyperlink r:id="rIdHyperlink"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>beta</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> gamma.</w:t></w:r></w:p><w:p w14:paraId="B1B2C3D4"><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Untouched bold sentinel.</w:t></w:r></w:p><w:p w14:paraId="C1B2C3D4"><w:r><w:t>Foreign lead</w:t></w:r><w:ins w:id="901" w:author="Foreign Reviewer" w:date="2026-09-01T12:00:00Z"><w:r><w:t xml:space="preserve"> external tail</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p><w:p w14:paraId="D1B2C3D4"><w:r><w:t>Untouched plain paragraph.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ReviewList"><w:name w:val="Review List"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/></w:style></w:styles>`;
    const numberingXml = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="10"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:num w:numId="11"><w:abstractNumId w:val="10"/></w:num></w:numbering>`;
    return zipDocx(new Map([
        ['[Content_Types].xml', contentTypes],
        ['_rels/.rels', rootRels],
        ['word/_rels/document.xml.rels', documentRels],
        ['word/document.xml', documentXml],
        ['word/styles.xml', stylesXml],
        ['word/numbering.xml', numberingXml]
    ]));
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const parser = createParser();
const serializer = createSerializer();

function parseXml(xml, label) {
    const document = parser.parseFromString(xml, 'application/xml');
    assert.ok(document.documentElement, `${label} should have a root element`);
    assert.equal(document.getElementsByTagName('parsererror').length, 0, `${label} should parse`);
    return document;
}

function xmlPart(entries, name) {
    const bytes = entries.get(name);
    assert.ok(bytes, `Missing DOCX part ${name}`);
    return decoder.decode(bytes);
}

function getRelationships(xml) {
    const document = parseXml(xml, 'relationships');
    return Array.from(document.getElementsByTagNameNS(PR, 'Relationship')).map(relationship => ({
        id: relationship.getAttribute('Id'),
        type: relationship.getAttribute('Type'),
        target: relationship.getAttribute('Target'),
        targetMode: relationship.getAttribute('TargetMode')
    }));
}

function packageToFlatOpc(bytes) {
    const entries = unzipDocx(bytes);
    const types = new Map([
        ['[Content_Types].xml', 'application/xml'],
        ['_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml'],
        ['word/_rels/document.xml.rels', 'application/vnd.openxmlformats-package.relationships+xml'],
        ['word/document.xml', NS.document],
        ['word/styles.xml', NS.styles],
        ['word/numbering.xml', NS.numbering]
    ]);
    const parts = Array.from(entries, ([name, bytes]) => {
        let xml = decoder.decode(bytes).replace(/^\s*<\?xml[^?]*\?>\s*/, '');
        assert.ok(!xml.includes('<pkg:binaryData'), `Fixture part ${name} should be XML`);
        const contentType = types.get(name) || 'application/xml';
        return `<pkg:part pkg:name="/${name}" pkg:contentType="${contentType}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8"?><pkg:package xmlns:pkg="${PKG}">${parts}</pkg:package>`;
}

function flatOpcParts(payload) {
    const packageDoc = parseXml(payload, 'Word insertion package');
    return Array.from(packageDoc.getElementsByTagNameNS(PKG, 'part')).map(part => {
        const name = (part.getAttribute('pkg:name') || part.getAttribute('name')).replace(/^\/+/, '');
        const contentType = part.getAttribute('pkg:contentType') || part.getAttribute('contentType');
        const xmlData = part.getElementsByTagNameNS(PKG, 'xmlData')[0];
        const root = Array.from(xmlData?.childNodes || []).find(node => node.nodeType === 1);
        assert.ok(root, `Flat OPC part ${name} should contain XML data`);
        return { name, contentType, xml: serializer.serializeToString(root) };
    });
}

function addContentTypeOverrides(entries, parts) {
    let contentTypesXml = xmlPart(entries, '[Content_Types].xml');
    const contentTypesDoc = parseXml(contentTypesXml, 'content types');
    const overrides = Array.from(contentTypesDoc.getElementsByTagNameNS(CT, 'Override'));
    for (const part of parts) {
        const partName = `/${part.name}`;
        if (!part.contentType || overrides.some(node => node.getAttribute('PartName') === partName)) continue;
        const override = contentTypesDoc.createElementNS(CT, 'Override');
        override.setAttribute('PartName', partName);
        override.setAttribute('ContentType', part.contentType);
        contentTypesDoc.documentElement.appendChild(override);
    }
    contentTypesXml = serializer.serializeToString(contentTypesDoc.documentElement);
    entries.set('[Content_Types].xml', encoder.encode(contentTypesXml));
}

function materializeWordInsertion(sourceBytes, insertionPayload) {
    const entries = unzipDocx(sourceBytes);
    const parts = flatOpcParts(insertionPayload);
    for (const part of parts) entries.set(part.name, encoder.encode(part.xml));
    addContentTypeOverrides(entries, parts);
    return zipDocx(entries);
}

function assertUncompressedPackageUnchanged(beforeBytes, afterBytes, label) {
    const before = unzipDocx(beforeBytes);
    const after = unzipDocx(afterBytes);
    assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort(), `${label}: no package parts added or removed`);
    for (const [name, bytes] of before) {
        assert.deepEqual(after.get(name), bytes, `${label}: ${name} should not be mutated`);
    }
}

function assertOriginalRelationshipsPreserved(beforeBytes, afterBytes, label) {
    const before = getRelationships(xmlPart(unzipDocx(beforeBytes), 'word/_rels/document.xml.rels'));
    const after = getRelationships(xmlPart(unzipDocx(afterBytes), 'word/_rels/document.xml.rels'));
    for (const relationship of before) {
        assert.ok(after.some(candidate => JSON.stringify(candidate) === JSON.stringify(relationship)),
            `${label}: preserve ${relationship.type} relationship ${relationship.id}`);
    }
}

function paragraphText(paragraph) {
    return Array.from(paragraph.getElementsByTagNameNS(W, 't')).map(node => node.textContent).join('');
}

function assertPreservedStructure(bytes, sourceBytes, label) {
    const entries = unzipDocx(bytes);
    const sourceEntries = unzipDocx(sourceBytes);
    assert.deepEqual(entries.get('word/styles.xml'), sourceEntries.get('word/styles.xml'), `${label}: style definitions stay unchanged`);
    assert.deepEqual(entries.get('word/numbering.xml'), sourceEntries.get('word/numbering.xml'), `${label}: numbering definitions stay unchanged`);
    assertOriginalRelationshipsPreserved(sourceBytes, bytes, label);

    const document = parseXml(xmlPart(entries, 'word/document.xml'), `${label} document.xml`);
    const paragraphs = Array.from(document.getElementsByTagNameNS(W, 'body')[0].childNodes)
        .filter(node => node.nodeType === 1 && node.namespaceURI === W && node.localName === 'p');
    assert.equal(paragraphs.length, 4, `${label}: paragraph count stays fixed`);

    const firstPPr = paragraphs[0].getElementsByTagNameNS(W, 'pPr')[0];
    assert.equal(firstPPr.getElementsByTagNameNS(W, 'pStyle')[0]?.getAttributeNS(W, 'val'), 'ReviewList', `${label}: paragraph style survives`);
    assert.equal(firstPPr.getElementsByTagNameNS(W, 'numId')[0]?.getAttributeNS(W, 'val'), '11', `${label}: numbering binding survives`);
    const hyperlinkIds = Array.from(paragraphs[0].getElementsByTagNameNS(W, 'hyperlink'))
        .map(node => node.getAttributeNS(R, 'id'));
    assert.ok(hyperlinkIds.length > 0 && hyperlinkIds.every(id => id === 'rIdHyperlink'),
        `${label}: hyperlink boundaries still refer to the original relationship`);
    assert.ok(paragraphs[1].getElementsByTagNameNS(W, 'b').length > 0, `${label}: untouched bold formatting survives`);

    const foreignInsertion = Array.from(document.getElementsByTagNameNS(W, 'ins')).find(
        node => node.getAttributeNS(W, 'author') === 'Foreign Reviewer'
    );
    assert.ok(foreignInsertion, `${label}: foreign revision remains present`);
    assert.equal(paragraphText(foreignInsertion), ' external tail', `${label}: foreign revision text remains attributed`);

    const relationships = getRelationships(xmlPart(entries, 'word/_rels/document.xml.rels'));
    assert.ok(relationships.some(relationship => relationship.type === `${R}/comments` && relationship.target === 'comments.xml'),
        `${label}: added comment part has a document relationship`);
    for (const marker of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
        assert.equal(document.getElementsByTagNameNS(W, marker).length, 1, `${label}: comment ${marker} is anchored once`);
    }
}

async function acceptedAndRejectedTexts(bytes, label) {
    const accepted = openDocx(bytes);
    const rejected = openDocx(bytes);
    const acceptedResult = await accepted.resolveRevisions('accept', { allAuthors: true });
    const rejectedResult = await rejected.resolveRevisions('reject', { allAuthors: true });
    assert.equal(acceptedResult.status, 'ok', `${label} accept: ${JSON.stringify(acceptedResult.error)}`);
    assert.equal(rejectedResult.status, 'ok', `${label} reject: ${JSON.stringify(rejectedResult.error)}`);
    const expectedAccepted = [
        'Alpha BETA gamma.',
        'Untouched bold sentinel.',
        'Foreign lead external tail.',
        'Untouched plain paragraph.'
    ];
    const expectedRejected = [
        'Alpha beta gamma.',
        'Untouched bold sentinel.',
        'Foreign lead.',
        'Untouched plain paragraph.'
    ];
    assert.deepEqual(accepted.inspect().paragraphs.map(item => item.exactText), expectedAccepted, `${label}: exact Accept All text`);
    assert.deepEqual(rejected.inspect().paragraphs.map(item => item.exactText), expectedRejected, `${label}: exact Reject All text`);
    for (const [resolvedDoc, view] of [[accepted, 'Accept All'], [rejected, 'Reject All']]) {
        assert.equal(resolvedDoc.inspect().comments.length, 1, `${label}: comment survives ${view}`);
        assert.equal(resolvedDoc.inspect().comments[0].text, 'Review the bold sentinel.');
        assert.equal(resolvedDoc.inspect().comments[0].author, author);
    }
}

function batchOperations() {
    return [
        {
            type: 'replace',
            target: { exactText: 'Alpha beta gamma.' },
            replacements: [{ find: 'beta', replace: 'BETA' }]
        },
        {
            type: 'comment',
            target: { exactText: 'Untouched bold sentinel.' },
            textToComment: 'bold',
            commentContent: 'Review the bold sentinel.'
        }
    ];
}

function getErrorCodes(result) {
    if (!result || typeof result !== 'object') return [];
    const codes = new Set();
    const visit = (value, depth = 0) => {
        if (!value || typeof value !== 'object' || depth > 5) return;
        if (typeof value.code === 'string') codes.add(value.code);
        if (value.error) visit(value.error, depth + 1);
        if (Array.isArray(value.results)) {
            for (const item of value.results) visit(item, depth + 1);
        }
        if (value.details) visit(value.details, depth + 1);
        if (Array.isArray(value)) {
            for (const item of value) visit(item, depth + 1);
        }
    };
    visit(result);
    return [...codes];
}

async function captureFailure(run) {
    try {
        return { result: await run(), error: null };
    } catch (error) {
        return { result: null, error };
    }
}

async function runSuccessfulParity(sourceBytes) {
    const operations = batchOperations();
    const flatSource = packageToFlatOpc(sourceBytes);
    const core = await prepareCanonicalBatch(flatSource, operations, options);
    assert.equal(core.status, 'ready', `portable core: ${JSON.stringify(core.result?.error || core.result?.results)}`);
    const coreBytes = materializeWordInsertion(sourceBytes, core.insertionPayload);

    let wordInsertionCount = 0;
    let wordInsertionPayload = null;
    const wordScope = {
        getOoxml() { return { value: packageToFlatOpc(sourceBytes) }; },
        insertOoxml(payload, mode) {
            wordInsertionCount++;
            wordInsertionPayload = payload;
            assert.ok(mode, 'Word adapter supplies an insertion mode');
        }
    };
    const word = await executePureOoxmlBatch({ async sync() {} }, wordScope, operations, options);
    assert.equal(word.status, 'ok', `Word adapter: ${JSON.stringify(word.error || word.results)}`);
    assert.equal(word.written, true);
    assert.equal(wordInsertionCount, 1, 'Word adapter performs one insertion for the batch');
    const wordBytes = materializeWordInsertion(sourceBytes, wordInsertionPayload);

    const mcpDoc = openDocx(sourceBytes);
    const mcp = await applyDocumentOperations(mcpDoc, operations, options);
    assert.equal(mcp.written, true);
    const mcpBytes = mcpDoc.toUint8Array();

    const browser = createDocumentSession(sourceBytes);
    const browserResult = await browser.applyOperations(operations, options);
    assert.equal(browserResult.status, 'ok', `browser facade: ${JSON.stringify(browserResult.error || browserResult.results)}`);
    const browserBytes = browser.toUint8Array();

    const outputs = [
        ['portable core', coreBytes],
        ['Word batch adapter', wordBytes],
        ['MCP document service', mcpBytes],
        ['browser document facade', browserBytes]
    ];
    for (const [label, bytes] of outputs) {
        await validateDocxPackage(new MemoryZip(unzipDocx(bytes)));
        assertPreservedStructure(bytes, sourceBytes, label);
        await acceptedAndRejectedTexts(bytes, label);
    }
}

async function runAtomicRefusalParity(sourceBytes) {
    const badOperations = [
        {
            type: 'replace',
            target: { exactText: 'Alpha beta gamma.' },
            replacements: [{ find: 'beta', replace: 'MUST NOT COMMIT' }]
        },
        { type: 'replace', target: { exactText: 'missing source paragraph' }, modified: 'also refused' }
    ];
    const flatSource = packageToFlatOpc(sourceBytes);
    const core = await prepareCanonicalBatch(flatSource, badOperations, options);
    assert.equal(core.status, 'refused');
    assert.equal(core.result?.error?.code, 'BATCH_OPERATION_FAILED', 'portable core retains the batch wrapper');
    assert.ok(getErrorCodes(core.result).includes('TARGET_NOT_FOUND'), 'portable core preserves the engine error code');
    assert.equal(core.insertionPayload, undefined, 'refused portable batches have no insertion payload');

    let wordWrites = 0;
    const wordScope = {
        getOoxml() { return { value: packageToFlatOpc(sourceBytes) }; },
        insertOoxml() { wordWrites++; }
    };
    const word = await executePureOoxmlBatch({ async sync() {} }, wordScope, badOperations, options);
    assert.equal(word.error?.code, 'BATCH_OPERATION_FAILED', 'Word adapter retains the batch wrapper');
    assert.ok(getErrorCodes(word).includes('TARGET_NOT_FOUND'), 'Word adapter preserves the engine error code');
    assert.equal(word.written, false);
    assert.equal(word.writeAttempted, false);
    assert.equal(wordWrites, 0, 'failed Word batch never inserts');

    const mcpDoc = openDocx(sourceBytes);
    const mcpBefore = mcpDoc.toUint8Array();
    const mcpFailure = await captureFailure(() => applyDocumentOperations(mcpDoc, badOperations, options));
    assert.equal(mcpFailure.error?.code, 'TARGET_NOT_FOUND', 'MCP service exposes the structured operation code');
    assert.ok(getErrorCodes(mcpFailure.error).includes('TARGET_NOT_FOUND'), 'MCP facade preserves a generic structured target error');
    assert.ok(mcpFailure.error?.message);
    assertUncompressedPackageUnchanged(mcpBefore, mcpDoc.toUint8Array(), 'MCP refused batch');

    const browser = createDocumentSession(sourceBytes);
    const browserBefore = browser.toUint8Array();
    const browserFailure = await captureFailure(() => browser.applyOperations(badOperations, options));
    const browserFailureResult = browserFailure.result;
    assert.equal(browserFailureResult?.error?.code, 'BATCH_OPERATION_FAILED', 'browser facade retains the batch wrapper');
    assert.ok(getErrorCodes(browserFailure.error || browserFailureResult).includes('TARGET_NOT_FOUND'), 'browser facade preserves the engine error code');
    assert.ok(browserFailure.error?.message || browserFailureResult?.error?.message);
    assertUncompressedPackageUnchanged(browserBefore, browser.toUint8Array(), 'browser refused batch');

    assertUncompressedPackageUnchanged(sourceBytes, mcpDoc.toUint8Array(), 'shared source after MCP refusal');
    assertUncompressedPackageUnchanged(sourceBytes, browser.toUint8Array(), 'shared source after browser refusal');
}

async function run() {
    const sourceBytes = fixtureBytes();
    const source = openDocx(sourceBytes);
    assert.equal(source.inspect().status, 'ok');
    assert.deepEqual(source.inspect().paragraphs.map(item => item.exactText), [
        'Alpha beta gamma.',
        'Untouched bold sentinel.',
        'Foreign lead external tail.',
        'Untouched plain paragraph.'
    ], 'all hosts receive the same fixture bytes and source view');

    await runSuccessfulParity(sourceBytes);
    await runAtomicRefusalParity(sourceBytes);
    console.log('PASS: cross-host DOCX operation parity');
}

run().catch(error => {
    console.error('FAIL:', error?.stack || error);
    process.exit(1);
});
