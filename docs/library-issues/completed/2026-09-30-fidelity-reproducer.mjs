/**
 * Minimal DOCX fidelity reproducer for @ansonlai/docx-redline-js v0.8.1.
 * Run: node docs/library-issues/completed/2026-09-30-fidelity-reproducer.mjs
 * Optional: DOCX_REDLINE_SOURCE_ROOT points to a source checkout (uses index.js,
 * never dist). The script exits 1 when either independent correctness oracle fails.
 */
import { pathToFileURL } from 'node:url';
import { resolve } from 'node:path';

const sourceRoot = process.env.DOCX_REDLINE_SOURCE_ROOT;
const api = await import(sourceRoot ? pathToFileURL(resolve(sourceRoot, 'index.js')).href : '@ansonlai/docx-redline-js');
const zipApi = await import(sourceRoot ? pathToFileURL(resolve(sourceRoot, 'document/zip-archive.js')).href : '@ansonlai/docx-redline-js/document/zip-archive.js');
const { openDocx } = api;
const { zipDocx, unzipDocx } = zipApi;
const enc = new TextEncoder(), dec = new TextDecoder();
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

function fixture(paragraphBody, externalLink = false) {
    return zipDocx(new Map([
        ['[Content_Types].xml', enc.encode(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
        ['_rels/.rels', enc.encode(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
        ['word/document.xml', enc.encode(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p>${paragraphBody}</w:p><w:sectPr/></w:body></w:document>`)],
        ['word/_rels/document.xml.rels', enc.encode(`<Relationships xmlns="${REL}">${externalLink ? `<Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.org/" TargetMode="External"/>` : ''}</Relationships>`)]
    ]));
}
const textOf = doc => doc.inspect().paragraphs.map(p => p.exactText).join('\n');
const results = [];

async function reproduce({ name, source, expectedSource, find, replacement, expectedAccepted, expectedHyperlink }) {
    const doc = openDocx(source), inspected = doc.inspect().paragraphs[0];
    if (textOf(doc) !== expectedSource) throw new Error(`${name}: source fixture does not match independent expected text`);
    const applied = await doc.applyOperations([{
        type: 'replace', target: { paragraphId: inspected.paragraphId, fingerprint: inspected.fingerprint },
        replacements: [{ find, replace: replacement }], insertionAffinity: { hyperlink: 'preserve', formatting: 'left' }
    }], { author: 'Fidelity reproducer', atomic: true });
    if (applied.status !== 'ok') throw new Error(`${name}: apply failed: ${JSON.stringify(applied.error || applied.results)}`);
    const tracked = doc.toUint8Array();
    const states = {};
    for (const [state, resolution, expectedText] of [['accepted', 'accept', expectedAccepted], ['rejected', 'reject', expectedSource]]) {
        const resolvedDoc = openDocx(tracked);
        const resolutionResult = await resolvedDoc.resolveRevisions(resolution, { allAuthors: true });
        if (resolutionResult.status !== 'ok') throw new Error(`${name}: ${state} resolution failed: ${JSON.stringify(resolutionResult.error)}`);
        const documentXml = dec.decode(unzipDocx(resolvedDoc.toUint8Array()).get('word/document.xml'));
        // The fixture uses a single relationship and simple w:t nodes, so inspect
        // hyperlink display text without adding a dependency or XML-provider setup.
        const hyperlinks = [...documentXml.matchAll(/<w:hyperlink\b[^>]*>([\s\S]*?)<\/w:hyperlink>/g)].map(match =>
            [...match[1].matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map(text => text[1]).join(''));
        const actualText = textOf(resolvedDoc);
        states[state] = { expectedText, actualText, textMatches: actualText === expectedText,
            ...(expectedHyperlink ? { expectedHyperlinkText: expectedHyperlink[state], actualHyperlinkText: hyperlinks.join(''),
                hyperlinkMatches: hyperlinks.join('') === expectedHyperlink[state] } : {}) };
    }
    results.push({ name, ...states, passed: Object.values(states).every(state => state.textMatches && state.hyperlinkMatches !== false) });
}

await reproduce({
    name: 'localized-hyperlink-boundary',
    source: fixture('<w:hyperlink r:id="rIdLink"><w:r><w:t>example.org</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r>', true),
    expectedSource: 'example.org.', find: 'example.org', replacement: 'example.net', expectedAccepted: 'example.net.',
    expectedHyperlink: { accepted: 'example.net', rejected: 'example.org' }
});
await reproduce({
    name: 'localized-tab-line-break-rejection',
    source: fixture('<w:r><w:tab/><w:t>tabbed</w:t><w:br/><w:t xml:space="preserve">Line with </w:t></w:r>'),
    expectedSource: '\ttabbed\nLine with ', find: 'tabbed', replacement: 'aligned', expectedAccepted: '\taligned\nLine with '
});
console.log(JSON.stringify({ implementation: sourceRoot ? 'source checkout' : 'installed package', results }, null, 2));
if (results.some(result => !result.passed)) process.exitCode = 1;
