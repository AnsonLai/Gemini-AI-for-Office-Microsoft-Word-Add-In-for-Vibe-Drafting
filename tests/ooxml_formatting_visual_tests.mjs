/** Semantic OOXML fidelity guards; rendered Word evidence is a separate lane. */
import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { openDocx, validateDocxPackage } from '@ansonlai/docx-redline-js';
import { MemoryZip, zipDocx, unzipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { createParser, createSerializer } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';
import { executePureOoxmlBatch } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const enc = new TextEncoder(), dec = new TextDecoder();
const load = name => new Uint8Array(readFileSync(new URL(`./fixtures/wp6/${name}`, import.meta.url)));
const xml = (entries, name) => dec.decode(entries.get(name) || new Uint8Array());
const parse = value => createParser().parseFromString(value, 'application/xml');
const serialize = node => createSerializer().serializeToString(node);
const nodes = (root, local) => Array.from(root.getElementsByTagNameNS(W, local));
const stateText = bytes => openDocx(bytes).inspect().paragraphs.map(p => p.exactText).join('\n');
const evidence = [];
const options = { author: 'WP6 reviewer', atomic: true, date: '2026-09-30T12:00:00Z' };
function success(result, label) { assert.equal(result.status, 'ok', `${label}: ${JSON.stringify(result.error || result.results)}`); }
function keys(entries) {
    const comments = parse(xml(entries, 'word/comments.xml'));
    return nodes(comments, 'comment').map(comment => nodes(comment, 'p').at(-1).getAttributeNS(W14, 'paraId'));
}
function anchored(entries, expectedIds) {
    const tree = parse(xml(entries, 'word/document.xml'));
    for (const id of expectedIds) {
        for (const kind of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
            assert.equal(nodes(tree, kind).filter(n => n.getAttribute('w:id') === String(id)).length, 1, `${kind} ${id} remains anchored exactly once`);
        }
    }
}
function siblingEntries(entries) {
    const ids = parse(xml(entries, 'word/commentsIds.xml'));
    const extensible = parse(xml(entries, 'word/commentsExtensible.xml'));
    return {
        ids: Array.from(ids.getElementsByTagNameNS('*', 'commentId')).map(n => ({ paraId: n.getAttribute('w16cid:paraId'), durableId: n.getAttribute('w16cid:durableId') })),
        durable: Array.from(extensible.getElementsByTagNameNS('*', 'commentExtensible')).map(n => n.getAttribute('w16cex:durableId'))
    };
}
function synchronized(entries, count) {
    const siblings = siblingEntries(entries);
    assert.equal(siblings.ids.length, count, 'one commentsIds entry per comment, keyed on last paragraph');
    assert.equal(new Set(siblings.ids.map(n => n.durableId)).size, count, 'durable IDs unique');
    assert.deepEqual(new Set(siblings.ids.map(n => n.paraId)), new Set(keys(entries)));
    assert.deepEqual(new Set(siblings.durable), new Set(siblings.ids.map(n => n.durableId)));
}
async function record(name, source, tracked, checks = {}) {
    const acceptedDoc = openDocx(tracked), rejectedDoc = openDocx(tracked);
    success(await acceptedDoc.resolveRevisions('accept', { allAuthors: true }), `${name} accept`);
    success(await rejectedDoc.resolveRevisions('reject', { allAuthors: true }), `${name} reject`);
    const accepted = acceptedDoc.toUint8Array(), rejected = rejectedDoc.toUint8Array();
    assert.equal(stateText(rejected), stateText(source), `${name} Reject All exact source text`);
    assert.equal(stateText(accepted), stateText(tracked), `${name} Accept All exact current text`);
    evidence.push({ name, sourceBytes: source, trackedBytes: tracked, acceptedBytes: accepted, rejectedBytes: rejected,
        expectedAcceptedText: stateText(accepted), expectedRejectedText: stateText(rejected), ...checks });
    if (checks.expectedComments != null) {
        const comments = openDocx(tracked).inspect().comments;
        const identity = comment => ({ author: comment.author, text: comment.text });
        assert.equal(new Set(comments.map(comment => JSON.stringify(identity(comment)))).size, comments.length,
            'author and text uniquely identify every comment in the Word oracle fixture');
        evidence.at(-1).expectedThreadIdentities = comments.map(comment => ({
            ...identity(comment),
            parent: comment.parentCommentId == null ? null : identity(comments.find(parent => parent.id === comment.parentCommentId)),
            done: comment.done
        }));
    }
    return { accepted, rejected };
}

// Small synthetic source: target only literal spans; retained runs and section structures are oracles.
const entries = new Map([
    ['[Content_Types].xml', enc.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')],
    ['_rels/.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', load('formatting-layout.xml')],
    ['word/styles.xml', enc.encode(`<w:styles xmlns:w="${W}"><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`)],
    ['word/_rels/document.xml.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/" TargetMode="External"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')]
]);
const source = zipDocx(entries), doc = openDocx(source), inspected = doc.inspect().paragraphs[0];
assert.equal(inspected.exactText, 'Bold term and italic term');
const changed = await doc.applyOperations([{ type: 'replace', target: { paragraphId: inspected.paragraphId, fingerprint: inspected.fingerprint }, replacements: [
    { find: 'Bold term', replace: 'Bold clause' }, { find: 'italic term', replace: 'italic clause' }
] }], options);
success(changed, 'localized formatting');
assert.equal(doc.inspect().paragraphs[0].exactText, 'Bold clause and italic clause');
const tracked = doc.toUint8Array(), after = unzipDocx(tracked);
await validateDocxPackage(new MemoryZip(after));
for (const name of ['word/styles.xml', 'word/_rels/document.xml.rels']) assert.deepEqual(after.get(name), entries.get(name), `${name} unchanged`);
const sourceTree = parse(xml(entries, 'word/document.xml')), trackedTree = parse(xml(after, 'word/document.xml'));
assert.deepEqual(nodes(trackedTree, 'sectPr').map(serialize), nodes(sourceTree, 'sectPr').map(serialize), 'both section layouts remain exact');
assert.deepEqual(nodes(trackedTree, 'tabs').map(serialize), nodes(sourceTree, 'tabs').map(serialize), 'tab stops retained');
assert.equal(nodes(trackedTree, 'tab').filter(n => n.parentNode.localName === 'r').length, 1);
assert.equal(nodes(trackedTree, 'br').length, 1);
assert.ok(nodes(trackedTree, 'hyperlink').length >= 1, 'hyperlink containers retained across tracked replacement');
assert.ok(nodes(trackedTree, 'hyperlink').every(n => n.getAttribute('r:id') === 'rIdLink'));
for (const [needle, property] of [['clause', 'b'], ['clause', 'i']]) {
    const inserted = nodes(trackedTree, 'ins').flatMap(n => nodes(n, 'r')).filter(r => nodes(r, 't').map(n => n.textContent).join('').includes(needle));
    assert.ok(inserted.some(r => nodes(r, property).length), `${property} formatting inherited into replacement`);
}
const lifecycle = await record('formatting-layout-localized', source, tracked, {
    expectedMinimumBodyRevisions: 1,
    expectedFormatting: [
        { sourceText: 'Bold term', acceptedText: 'Bold clause', bold: true, italic: false, font: 'Calibri', size: 12 },
        { sourceText: 'italic term', acceptedText: 'italic clause', bold: false, italic: true },
        { sourceText: ' and ', acceptedText: ' and ', bold: false, italic: false }
    ],
    expectedSections: [{ orientation: 0, columns: 1 }, { orientation: 1, columns: 2 }],
    expectedTabStop: { paragraph: 1, positionPoints: 360, alignment: 2 }
});
const boldInserted = nodes(trackedTree, 'ins').flatMap(n => nodes(n, 'r')).find(r => nodes(r, 'b').length && nodes(r, 't').some(n => n.textContent.includes('clause')));
assert.equal(nodes(boldInserted, 'rFonts')[0].getAttribute('w:ascii'), 'Calibri', 'inserted bold run retains font');
assert.equal(nodes(boldInserted, 'sz')[0].getAttribute('w:val'), '24', 'inserted bold run retains size');
const acceptedTree = parse(xml(unzipDocx(lifecycle.accepted), 'word/document.xml'));
assert.equal(nodes(acceptedTree, 'hyperlink').map(n => n.textContent).join(''), 'example.org');
assert.equal(nodes(parse(xml(unzipDocx(lifecycle.rejected), 'word/document.xml')), 'hyperlink').map(n => n.textContent).join(''), 'example.org');
const plainRun = nodes(acceptedTree, 'r').find(r => nodes(r, 't').some(t => t.textContent.includes(' and ')));
assert.equal(nodes(plainRun, 'b').length + nodes(plainRun, 'i').length, 0, 'formatting does not bleed into adjacent plain run');

// Real Word last-paragraph identity and root/reply lifecycle.
const threadSource = load('multi-paragraph-thread.docx'), threadDoc = openDocx(threadSource);
await validateDocxPackage(new MemoryZip(unzipDocx(threadSource)));
const threadComments = threadDoc.inspect().comments;
assert.equal(threadComments.length, 2);
const root = threadComments.find(c => c.id === '0'), reply = threadComments.find(c => c.id === '1');
assert.equal(nodes(parse(xml(unzipDocx(threadSource), 'word/comments.xml')), 'comment')[0].getElementsByTagNameNS(W, 'p').length, 3);
assert.equal(root.paraId, keys(unzipDocx(threadSource))[0]);
assert.equal(reply.parentCommentId, root.id);
success(await threadDoc.applyOperations([{ type: 'comment_reply', parentCommentId: reply.id, commentContent: 'WP6 reply to reply.' }], options), 'Word multi-paragraph reply');
assert.equal(threadDoc.inspect().comments.at(-1).parentCommentId, root.id, 'Word threads flatten reply-to-reply');
await validateDocxPackage(new MemoryZip(unzipDocx(threadDoc.toUint8Array())));
anchored(unzipDocx(threadDoc.toUint8Array()), threadDoc.inspect().comments.map(c => c.id));
await record('word-multi-paragraph-reply', threadSource, threadDoc.toUint8Array(), { expectedComments: 3 });
success(await threadDoc.resolveComment(reply.id), 'resolve through reply');
assert.ok(threadDoc.inspect().comments.every(c => c.done), 'resolve affects whole thread');
success(await threadDoc.resolveComment(root.id, { resolved: false }), 'reopen');
assert.ok(threadDoc.inspect().comments.every(c => !c.done));
success(await threadDoc.deleteComments({ ids: [root.id] }), 'delete root cascade');
assert.equal(threadDoc.inspect().comments.length, 0);

const siblingSource = load('threaded-comments.docx'), siblingDoc = openDocx(siblingSource);
const beforeSiblings = unzipDocx(siblingSource);
synchronized(beforeSiblings, 3);
success(await siblingDoc.applyOperations([{ type: 'comment_reply', parentCommentId: '0', commentContent: 'WP6 synchronized reply.' }], options), 'optional sibling reply');
const siblingTracked = siblingDoc.toUint8Array(), siblingAfter = unzipDocx(siblingTracked);
synchronized(siblingAfter, 4);
anchored(siblingAfter, siblingDoc.inspect().comments.map(c => c.id));
for (const original of siblingEntries(beforeSiblings).ids) assert.ok(siblingEntries(siblingAfter).ids.some(n => n.paraId === original.paraId && n.durableId === original.durableId), 'existing durable IDs preserved');
await validateDocxPackage(new MemoryZip(siblingAfter));
await record('word-optional-sibling-reply', siblingSource, siblingTracked, { expectedComments: 4 });
success(await siblingDoc.deleteComments({ author: 'WP6 reviewer' }), 'delete only added reply');
synchronized(unzipDocx(siblingDoc.toUint8Array()), 3);
success(await siblingDoc.deleteComments({ ids: ['0'] }), 'delete root and its existing reply');
synchronized(unzipDocx(siblingDoc.toUint8Array()), 1);

// Recreate Word getOoxml Flat OPC from all XML source parts; verify actual add-in write result.
function flatOpc(sourceEntries) {
    const typeDoc = parse(xml(sourceEntries, '[Content_Types].xml'));
    const overrides = new Map(Array.from(typeDoc.getElementsByTagNameNS('*', 'Override')).map(n => [n.getAttribute('PartName'), n.getAttribute('ContentType')]));
    const defaults = new Map(Array.from(typeDoc.getElementsByTagNameNS('*', 'Default')).map(n => [n.getAttribute('Extension'), n.getAttribute('ContentType')]));
    return `<pkg:package xmlns:pkg="${PKG}">` + [...sourceEntries].filter(([name]) => name !== '[Content_Types].xml').map(([name, value]) => {
        const contentType = overrides.get(`/${name}`) || defaults.get(name.split('.').at(-1));
        return contentType?.includes('xml') ? `<pkg:part pkg:name="/${name}" pkg:contentType="${contentType}"><pkg:xmlData>${dec.decode(value).replace(/<\?xml[^?]*\?>/, '')}</pkg:xmlData></pkg:part>` : '';
    }).join('') + '</pkg:package>';
}
// A combined formatted paragraph containing tabs, breaks and a hyperlink cannot
// be reconciled by the installed engine; require atomic refusal with no Word write.
const combinedTree = parse(xml(entries, 'word/document.xml'));
const combinedParagraphs = nodes(combinedTree, 'p');
for (const extra of combinedParagraphs.slice(1, 3)) {
    while (extra.firstChild) combinedParagraphs[0].appendChild(extra.firstChild);
    extra.parentNode.removeChild(extra);
}
const combinedEntries = new Map(entries);
combinedEntries.set('word/document.xml', enc.encode(serialize(combinedTree)));
const combinedSource = zipDocx(combinedEntries);
const structuralTarget = openDocx(combinedSource).inspect().paragraphs[0];
assert.equal(structuralTarget.exactText, 'Bold term and italic term\ttabbed\nLine with example.org.');
const refusedWrites = [];
const refusedScope = { getOoxml: () => ({ value: flatOpc(combinedEntries) }), insertOoxml: value => refusedWrites.push(value) };
const refused = await executePureOoxmlBatch({ sync: async () => {} }, refusedScope,
    [{ type: 'replace', target: { paragraphId: structuralTarget.paragraphId, fingerprint: structuralTarget.fingerprint }, replacements: [{ find: 'tabbed', replace: 'aligned' }] }], { author: 'WP6 reviewer' });
assert.equal(refused.status, 'error');
assert.equal(refused.written, false);
assert.equal(refused.rolledBack, true);
assert.equal(refusedWrites.length, 0, 'unsupported combined target must not write');

const writes = [];
const scope = { getOoxml: () => ({ value: flatOpc(beforeSiblings) }), insertOoxml: value => writes.push(value) };
const siblingNativeOperations = [{ type: 'comment_reply', parentCommentId: '0', commentContent: 'Add-in sibling synchronization.' }];
const bridge = await executePureOoxmlBatch({ sync: async () => {} }, scope,
    siblingNativeOperations, { author: 'WP6 reviewer' });
success(bridge, 'add-in optional siblings');
assert.equal(writes.length, 1);
const writtenParts = new Map(Array.from(parse(writes[0]).getElementsByTagNameNS(PKG, 'part')).map(p => {
    const rootNode = Array.from(p.getElementsByTagNameNS(PKG, 'xmlData')[0].childNodes).find(n => n.nodeType === 1);
    return [p.getAttribute('pkg:name').slice(1), enc.encode(serialize(rootNode))];
}));
synchronized(writtenParts, 4);
const bridgeBytes = zipDocx(new Map([...beforeSiblings, ...writtenParts]));
const bridgeComments = openDocx(bridgeBytes).inspect().comments;
assert.equal(bridgeComments.length, 4);
anchored(writtenParts, bridgeComments.map(c => c.id));
assert.equal(bridgeComments.at(-1).parentCommentId, '0');
await record('word-addin-sibling-reply', siblingSource, bridgeBytes, { expectedComments: 4, insertionXmlBytes: writes[0], nativeOperations: siblingNativeOperations });

// A successful body write repairs the pre-0.8.1 Flat OPC commentsExtended type
// even when no comment operation touches the existing thread parts.
const repairedWrites = [];
const staleTypePackage = flatOpc(beforeSiblings).replace(
    'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
    'application/vnd.ms-word.commentsExtended+xml');
assert.ok(staleTypePackage.includes('application/vnd.ms-word.commentsExtended+xml'));
const repairScope = { getOoxml: () => ({ value: staleTypePackage }), insertOoxml: value => repairedWrites.push(value) };
const repaired = await executePureOoxmlBatch({ sync: async () => {} }, repairScope,
    [{ type: 'replace', target: { exactText: 'This agreement is governed by local law.' }, replacements: [{ find: 'local', replace: 'national' }] }],
    { author: 'WP6 reviewer' });
success(repaired, 'body edit repairs stale commentsExtended type');
assert.equal(repairedWrites.length, 1);
const repairedTree = parse(repairedWrites[0]);
const repairedExtended = Array.from(repairedTree.getElementsByTagNameNS(PKG, 'part')).find(p => p.getAttribute('pkg:name') === '/word/commentsExtended.xml');
assert.equal(repairedExtended.getAttribute('pkg:contentType'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml');
const extendedRoot = Array.from(repairedExtended.getElementsByTagNameNS(PKG, 'xmlData')[0].childNodes).find(n => n.nodeType === 1);
assert.equal(serialize(extendedRoot), serialize(parse(xml(beforeSiblings, 'word/commentsExtended.xml')).documentElement),
    'body type repair preserves existing thread state');

// Existing Word-authored footer: facade part targeting and exact lifecycle.
const footerSource = load('header-footer.docx'), footerDoc = openDocx(footerSource);
const footerParts = footerDoc.inspect().headersFooters;
assert.equal(footerParts.length, 6, 'Word fixture carries first/default/even headers and footers');
const firstFooter = footerParts.find(part => part.kind === 'footer' && part.type === 'first');
assert.equal(firstFooter.path, 'word/footer3.xml');
assert.equal(firstFooter.paragraphs[0].text, 'First page footer');
assert.equal(footerParts.find(part => part.path === 'word/footer2.xml').hasFields, true, 'default footer PAGE field present');
const footerBefore = unzipDocx(footerSource);
const footerResult = await footerDoc.applyOperations([{
    type: 'replace', part: { kind: 'footer', type: 'first', section: 0 },
    target: { exactText: firstFooter.paragraphs[0].text }, replacements: [{ find: 'First', replace: 'Initial' }]
}], options);
success(footerResult, 'Word-authored localized footer edit');
assert.deepEqual(footerResult.artifactsChanged, ['word/footer3.xml']);
const footerTracked = footerDoc.toUint8Array(), footerAfter = unzipDocx(footerTracked);
await validateDocxPackage(new MemoryZip(footerAfter));
for (const [name, bytes] of footerBefore) {
    if (name !== 'word/footer3.xml') assert.deepEqual(footerAfter.get(name), bytes, `${name} unchanged by footer operation`);
}
const footerState = await record('word-authored-footer-localized', footerSource, footerTracked, {
    expectedHeaderFooterText: [{ kind: 'footer', type: 'first', section: 1, accepted: 'Initial page footer', rejected: 'First page footer' }]
});
for (const [state, expected] of [['accepted', 'Initial page footer'], ['rejected', 'First page footer']]) {
    const bytes = footerState[state], inspectedPart = openDocx(bytes).inspect().headersFooters.find(part => part.path === 'word/footer3.xml');
    assert.equal(inspectedPart.paragraphs[0].text, expected, `footer ${state} exact text`);
    const finalEntries = unzipDocx(bytes);
    await validateDocxPackage(new MemoryZip(finalEntries));
    assert.equal(nodes(parse(xml(finalEntries, 'word/footer3.xml')), 'ins').length + nodes(parse(xml(finalEntries, 'word/footer3.xml')), 'del').length, 0, `footer ${state} revisions resolved`);
    for (const [name, original] of footerBefore) {
        if (name !== 'word/footer3.xml') assert.deepEqual(finalEntries.get(name), original, `${name} preserved by ${state} footer lifecycle`);
    }
}
assert.deepEqual(load('header-footer.docx'), footerSource, 'checked-in Word source fixture unchanged');

// Isolate native insertion with a small plain-text body and no comments, tables,
// fields or hyperlinks in the document. Expectations come from the requested edit.
const plainEntries = new Map(entries);
plainEntries.delete('word/styles.xml');
plainEntries.set('[Content_Types].xml', enc.encode(xml(entries, '[Content_Types].xml').replace(/<Override PartName="\/word\/styles.xml"[^>]*\/>/, '')));
plainEntries.set('word/_rels/document.xml.rels', enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'));
plainEntries.set('word/document.xml', enc.encode(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Alpha beta gamma.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph stays unchanged.</w:t></w:r></w:p><w:p><w:r><w:t>Final paragraph stays unchanged.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`));
const plainSource = zipDocx(plainEntries), plainWrites = [];
const plainNativeOperations = [{ type: 'replace', target: { exactText: 'Alpha beta gamma.' }, replacements: [{ find: 'beta', replace: 'BETA' }] }];
const plainBridge = await executePureOoxmlBatch({ sync: async () => {} },
    { getOoxml: () => ({ value: flatOpc(plainEntries) }), insertOoxml: value => plainWrites.push(value) },
    plainNativeOperations,
    { author: 'WP6 reviewer' });
success(plainBridge, 'plain localized bridge body edit');
assert.equal(plainWrites.length, 1);
const plainWrittenParts = new Map(Array.from(parse(plainWrites[0]).getElementsByTagNameNS(PKG, 'part')).map(p => {
    const rootNode = Array.from(p.getElementsByTagNameNS(PKG, 'xmlData')[0].childNodes).find(n => n.nodeType === 1);
    return [p.getAttribute('pkg:name').slice(1), enc.encode(serialize(rootNode))];
}));
const plainTracked = zipDocx(new Map([...plainEntries, ...plainWrittenParts]));
const expectedPlainSource = 'Alpha beta gamma.\nSecond paragraph stays unchanged.\nFinal paragraph stays unchanged.';
const expectedPlainAccepted = 'Alpha BETA gamma.\nSecond paragraph stays unchanged.\nFinal paragraph stays unchanged.';
assert.equal(stateText(plainSource), expectedPlainSource);
assert.equal(stateText(plainTracked), expectedPlainAccepted);
await validateDocxPackage(new MemoryZip(unzipDocx(plainTracked)));
await record('word-addin-plain-replacement', plainSource, plainTracked, {
    insertionXmlBytes: plainWrites[0], nativeOperations: plainNativeOperations, expectedAcceptedText: expectedPlainAccepted, expectedRejectedText: expectedPlainSource, expectedMinimumBodyRevisions: 1
});

// v0.8.2 fixes library issues #3/#4. Both localized and full-paragraph forms
// are mandatory offline regressions and exported into the default live Word lane.
for (const defect of [
    { name: 'hyperlink-boundary', targetText: 'example.org.', find: 'example.org', replace: 'example.net',
        expectedHyperlinks: [{ target: 'https://example.org/', sourceText: 'example.org', acceptedText: 'example.net', rejectedText: 'example.org' }] },
    { name: 'structural-lifecycle', targetText: '\ttabbed\nLine with ', find: 'tabbed', replace: 'aligned' }
]) {
    for (const shape of ['localized', 'full-paragraph']) {
        const name = `regression-${defect.name}-${shape}`;
        const probe = openDocx(source), inspectedTarget = probe.inspect().paragraphs.find(p => p.exactText === defect.targetText);
        assert.ok(inspectedTarget, `${defect.name} source target present`);
        success(await probe.applyOperations([{ type: 'replace', target: { paragraphId: inspectedTarget.paragraphId, fingerprint: inspectedTarget.fingerprint },
            ...(shape === 'localized' ? { replacements: [{ find: defect.find, replace: defect.replace }] }
                : { modified: defect.targetText.replace(defect.find, defect.replace) }),
            insertionAffinity: { hyperlink: 'preserve', formatting: 'left' } }], options), `${name} apply`);
        const trackedBytes = probe.toUint8Array(), acceptedDoc = openDocx(trackedBytes), rejectedDoc = openDocx(trackedBytes);
        success(await acceptedDoc.resolveRevisions('accept', { allAuthors: true }), `${defect.name} engine accept`);
        success(await rejectedDoc.resolveRevisions('reject', { allAuthors: true }), `${defect.name} engine reject`);
        const acceptedBytes = acceptedDoc.toUint8Array(), rejectedBytes = rejectedDoc.toUint8Array();
        const expectedRejectedText = 'Bold term and italic term\n\ttabbed\nLine with \nexample.org.\nSection boundary.\nLandscape section.';
        assert.equal(stateText(source), expectedRejectedText, `${defect.name} independent source expectation`);
        const expectedAcceptedText = expectedRejectedText.replace(defect.find, defect.replace);
        assert.equal(stateText(acceptedBytes), expectedAcceptedText, `${name} exact accepted text`);
        assert.equal(stateText(rejectedBytes), expectedRejectedText, `${name} exact rejected text`);
        await validateDocxPackage(new MemoryZip(unzipDocx(trackedBytes)));
        if (defect.expectedHyperlinks) {
            for (const [state, bytes] of [['accepted', acceptedBytes], ['rejected', rejectedBytes]]) {
                const links = nodes(parse(xml(unzipDocx(bytes), 'word/document.xml')), 'hyperlink');
                assert.equal(links.map(node => node.textContent).join(''), defect.expectedHyperlinks[0][`${state}Text`], `${name} ${state} hyperlink boundary`);
                assert.ok(links.every(node => node.getAttribute('r:id') === 'rIdLink'), `${name} hyperlink relationship preserved`);
                assert.equal(xml(unzipDocx(bytes), 'word/_rels/document.xml.rels'), xml(entries, 'word/_rels/document.xml.rels'));
            }
        }
        evidence.push({ name, sourceBytes: source, trackedBytes, acceptedBytes, rejectedBytes,
            expectedAcceptedText, expectedRejectedText, expectedMinimumBodyRevisions: 1,
            observedEngineAcceptedText: stateText(acceptedBytes), observedEngineRejectedText: stateText(rejectedBytes),
            ...(defect.expectedHyperlinks ? { expectedHyperlinks: defect.expectedHyperlinks,
                observedEngineHyperlinks: { acceptedTexts: nodes(parse(xml(unzipDocx(acceptedBytes), 'word/document.xml')), 'hyperlink').map(n => n.textContent),
                    rejectedTexts: nodes(parse(xml(unzipDocx(rejectedBytes), 'word/document.xml')), 'hyperlink').map(n => n.textContent) } } : {}) });
    }
}
const exportAt = process.argv.indexOf('--export-dir');
if (exportAt >= 0) {
    if (!process.argv[exportAt + 1] || process.argv[exportAt + 1].startsWith('--')) throw new Error('--export-dir requires a directory');
    const destination = resolve(process.argv[exportAt + 1]);
    mkdirSync(destination, { recursive: true });
    const cases = evidence.map(({ sourceBytes, trackedBytes, acceptedBytes, rejectedBytes, insertionXmlBytes, ...item }) => {
        if (insertionXmlBytes) { item.insertionXml = `${item.name}-insertion.xml`; writeFileSync(join(destination, item.insertionXml), insertionXmlBytes); }
        for (const [state, bytes] of Object.entries({ source: sourceBytes, tracked: trackedBytes, accepted: acceptedBytes, rejected: rejectedBytes })) {
            item[state] = `${item.name}-${state}.docx`;
            writeFileSync(join(destination, item[state]), bytes);
        }
        return item;
    });
    const corruptEntries = unzipDocx(siblingSource);
    const correctType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
    const sourceTypes = xml(corruptEntries, '[Content_Types].xml');
    assert.ok(sourceTypes.includes(correctType));
    corruptEntries.set('[Content_Types].xml', enc.encode(sourceTypes.replace(correctType, 'application/vnd.ms-word.commentsExtended+xml')));
    const control = 'control-bad-commentsExtended-type.docx';
    writeFileSync(join(destination, control), zipDocx(corruptEntries));
    writeFileSync(join(destination, 'manifest.json'), JSON.stringify({ cases, controls: [{ file: control, expectOpen: false }] }, null, 2));
}
console.log('PASS: WP6 OOXML formatting, exact lifecycle, Word-authored threads and sibling parts');
