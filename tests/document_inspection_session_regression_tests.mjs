import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';
import {
  computeWordSourceRevisionToken,
  detectExportTailIndex
} from '../src/taskpane/modules/docx-redline-js-integration/word-source-identity.js';

// Phase-2 WP2 regression coverage: handles rely on the revision token alone; the
// library re-derives descriptors at resolve time.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const escapeXml = text => String(text).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
const resolveOne = async (session, handle) => {
  const result = await session.resolveHandles([handle]);
  return { ...result, target: result.targets?.[0] };
};

function documentParts(texts, { rsid = '00000001', settingsRsid = 'AAAA0001' } = {}) {
  const paragraphs = texts.map((text, index) =>
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0').toUpperCase()}" w:rsidR="${rsid}" w:rsidRDefault="${rsid}"><w:r w:rsidR="${rsid}"><w:t>${escapeXml(text)}</w:t></w:r></w:p>`
  ).join('');
  return {
    documentXml: `<?xml version="1.0"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`,
    commentsXml: null,
    commentsExtendedXml: null,
    commentsIdsXml: null,
    commentsExtensibleXml: null,
    numberingXml: null,
    stylesXml: `<w:styles xmlns:w="${W}"/>`,
    additionalParts: {
      'word/settings.xml': `<w:settings xmlns:w="${W}"><w:rsids><w:rsidRoot w:val="${settingsRsid}"/><w:rsid w:val="${rsid}"/></w:rsids><w:doNotTrackMoves/></w:settings>`
    }
  };
}

function makeSession(getParts, options = {}) {
  let documentId = 'addin:regression';
  const session = createDocumentInspectionSession({
    readSource: async () => ({ documentId, revisionView: 'accepted', parts: getParts() }),
    ...options
  });
  return { session, setDocumentId: id => { documentId = id; } };
}

const texts = ['Alpha clause', 'Second paragraph', 'Beta clause', 'Fourth paragraph', 'Gamma clause', 'Sixth paragraph'];

// 1. An unrelated host re-read that only changes rsids keeps handles valid and
//    resolves to a fresh library descriptor equal to the inspected one.
{
  let parts = documentParts(texts);
  const { session } = makeSession(() => parts, { computeRevisionToken: computeWordSourceRevisionToken });
  const page = await session.inspect({ range: { start: 1, end: 3 }, limit: 3 });
  assert.equal(page.status, 'ok', JSON.stringify(page));
  const before = await session.resolveHandles(page.targets.map(target => target.handle));
  assert.equal(before.status, 'ok', JSON.stringify(before));

  parts = documentParts(texts, { rsid: '0A0B0C0D', settingsRsid: 'BBBB0002' });
  assert.notEqual(parts.documentXml, documentParts(texts).documentXml, 'fixture really differs in rsids');
  const after = await session.resolveHandles(page.targets.map(target => target.handle));
  assert.equal(after.status, 'ok', `rsid-only change must not expire handles: ${JSON.stringify(after)}`);
  assert.equal(after.targets.length, 3);
  for (let i = 0; i < 3; i++) {
    for (const key of ['index', 'exactText', 'paragraphId', 'fingerprint']) {
      assert.deepEqual(after.targets[i].descriptor[key], before.targets[i].descriptor[key], `descriptor ${key} of target ${i}`);
    }
  }
  assert.equal(after.targets[1].descriptor.exactText, 'Second paragraph');
  assert.equal(after.source.revisionToken.value, before.source.revisionToken.value);

  // A real text change is still stale under the Word-aware token.
  parts = documentParts(['Alpha clause', 'Second paragraph CHANGED', ...texts.slice(2)], { rsid: '0A0B0C0D', settingsRsid: 'BBBB0002' });
  const stale = await session.resolveHandles(page.targets.map(target => target.handle));
  assert.equal(stale.error.code, 'STALE_DOCUMENT_SOURCE');
}

// 2. Any content change refuses with STALE_DOCUMENT_SOURCE and expires every
//    handle and cursor, whichever part changed.
{
  const w15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
  const mutations = {
    documentText: parts => { parts.documentXml = parts.documentXml.replace('Second paragraph', 'Second paragraph edited'); },
    commentsXml: parts => { parts.commentsXml = `<w:comments xmlns:w="${W}"><w:comment w:id="9"><w:p><w:r><w:t>New</w:t></w:r></w:p></w:comment></w:comments>`; },
    commentsExtendedXml: parts => { parts.commentsExtendedXml = `<w15:commentsEx xmlns:w15="${w15}"><w15:commentEx w15:paraId="AAAA0001" w15:done="1"/></w15:commentsEx>`; },
    numberingXml: parts => { parts.numberingXml = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="3"/></w:numbering>`; },
    stylesXml: parts => { parts.stylesXml = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Other"/></w:styles>`; },
    additionalPartChanged: parts => { parts.additionalParts['word/settings.xml'] = `<w:settings xmlns:w="${W}"><w:zoom w:percent="150"/></w:settings>`; },
    additionalPartAdded: parts => { parts.additionalParts['word/footnotes.xml'] = `<w:footnotes xmlns:w="${W}"/>`; },
    additionalPartRemoved: parts => { delete parts.additionalParts['word/settings.xml']; }
  };
  for (const [label, mutate] of Object.entries(mutations)) {
    const parts = documentParts(texts);
    const { session } = makeSession(() => parts);
    const first = await session.inspect({ limit: 2 });
    const second = await session.inspect({ range: { start: 5, end: 6 }, limit: 2 });
    assert.equal(first.status, 'ok');
    const cursor = first.selection.nextCursor;
    assert.equal(typeof cursor, 'string', 'first page issued a cursor');
    const handles = [...first.targets, ...second.targets].map(target => target.handle);

    mutate(parts);
    const refused = await session.resolveHandles([handles[0]]);
    assert.equal(refused.status, 'error', label);
    assert.equal(refused.error.code, 'STALE_DOCUMENT_SOURCE', label);
    for (const handle of handles) {
      const expired = await session.resolveHandles([handle]);
      assert.equal(expired.error.code, 'STALE_OR_UNKNOWN_HANDLE', `${label}: every handle expires`);
    }
    const staleCursor = await session.inspect({ cursor });
    assert.equal(staleCursor.error.code, 'STALE_CURSOR', `${label}: cursors expire`);
  }
}

// 2b. Changes observed by inspect (not resolve) also expire cursors.
{
  const parts = documentParts(texts);
  const { session } = makeSession(() => parts);
  const first = await session.inspect({ limit: 2 });
  const cursor = first.selection.nextCursor;
  parts.documentXml = parts.documentXml.replace('Fourth paragraph', 'Fourth edited');
  const paged = await session.inspect({ cursor });
  assert.equal(paged.status, 'error');
  assert.equal(paged.error.code, 'STALE_CURSOR');
  assert.equal((await resolveOne(session, first.targets[0].handle)).error.code, 'STALE_OR_UNKNOWN_HANDLE');
}

// 3. Cursor continuation repeats the same query; cursor plus new options is refused.
{
  const parts = documentParts(texts);
  const { session } = makeSession(() => parts);

  const search = await session.inspect({ search: 'clause', limit: 2 });
  assert.equal(search.selection.totalMatches, 3);
  assert.deepEqual(search.targets.map(target => target.text), ['Alpha clause', 'Beta clause']);
  const searchNext = await session.inspect({ cursor: search.selection.nextCursor });
  assert.equal(searchNext.status, 'ok', JSON.stringify(searchNext));
  assert.deepEqual(searchNext.targets.map(target => target.text), ['Gamma clause'], 'search continuation returns only remaining matches');
  assert.equal(searchNext.selection.truncated, false);
  assert.equal(searchNext.selection.nextCursor, undefined);

  const ranged = await session.inspect({ range: { start: 2, end: 6 }, limit: 2 });
  assert.deepEqual(ranged.targets.map(target => target.text), ['Second paragraph', 'Beta clause']);
  const rangedNext = await session.inspect({ cursor: ranged.selection.nextCursor });
  assert.deepEqual(rangedNext.targets.map(target => target.text), ['Fourth paragraph', 'Gamma clause'], 'range and limit are preserved');
  const rangedLast = await session.inspect({ cursor: rangedNext.selection.nextCursor });
  assert.deepEqual(rangedLast.targets.map(target => target.text), ['Sixth paragraph']);

  const around = await session.inspect({ search: 'Alpha', around: 1, limit: 1 });
  assert.equal(around.selection.around, 1);
  assert.deepEqual(around.targets.map(target => target.text), ['Alpha clause', 'Second paragraph'], 'around adds neighbour context');

  const aroundMulti = await session.inspect({ search: 'clause', around: 1, limit: 2 });
  const aroundNext = await session.inspect({ cursor: aroundMulti.selection.nextCursor });
  assert.equal(aroundNext.status, 'ok', JSON.stringify(aroundNext));
  assert.equal(aroundNext.selection.around, 1, 'around is preserved across a cursor');
  assert.ok(aroundNext.targets.some(target => target.text === 'Gamma clause'));

  const cursor = search.selection.nextCursor;
  for (const extra of [{ search: 'x' }, { range: { start: 1, end: 2 } }, { around: 1 }, { limit: 1 }, { includeDocumentEnd: true }]) {
    const refused = await session.inspect({ cursor, ...extra });
    assert.equal(refused.status, 'error');
    assert.equal(refused.error.code, 'INVALID_INSPECTION_OPTIONS', JSON.stringify(extra));
  }
  // A refused combination does not consume the cursor.
  assert.equal((await session.inspect({ cursor })).status, 'ok');
  assert.equal((await session.inspect({ cursor: 'cursor_unknown' })).error.code, 'STALE_CURSOR');
}

// 4. Document-end handles and the Word export-tail exclusion.
{
  const parts = documentParts([...texts, '']);
  const { session } = makeSession(() => parts, { getExcludedParagraphIndex: () => 7 });
  const page = await session.inspect({ range: { start: 5, end: 7 }, limit: 3, includeDocumentEnd: true });
  assert.equal(page.status, 'ok', JSON.stringify(page));
  assert.deepEqual(page.targets.map(target => target.text), ['Gamma clause', 'Sixth paragraph'], 'export tail is not offered as a target');
  const end = await resolveOne(session, page.documentEnd.handle);
  assert.equal(end.status, 'ok', JSON.stringify(end));
  assert.equal(end.target.kind, 'document-end');
  assert.equal(end.target.position, 'after');
  assert.equal(end.target.fullTextInspected, false);
  assert.equal(end.target.descriptor.index, 6, 'the end handle anchors on the last real paragraph, not the export tail');
  assert.equal(end.target.descriptor.exactText, 'Sixth paragraph');

  const noEnd = await session.inspect({ limit: 1 });
  assert.equal(noEnd.documentEnd, undefined, 'document end is opt-in');

  // Without the exclusion the empty final paragraph would be the end anchor.
  const { session: plain } = makeSession(() => parts);
  const plainEnd = await plain.inspect({ includeDocumentEnd: true, limit: 1 });
  assert.equal((await resolveOne(plain, plainEnd.documentEnd.handle)).target.descriptor.index, 7);

  // End to end with the Word-aware detector and a certified host paragraph count.
  const wordParts = documentParts([...texts, '']);
  let hostParagraphCount = 6;
  const detecting = createDocumentInspectionSession({
    readSource: async () => ({ documentId: 'addin:tail', revisionView: 'accepted', parts: wordParts, hostParagraphCount }),
    computeRevisionToken: computeWordSourceRevisionToken,
    getExcludedParagraphIndex: (sourceParts, raw) => detectExportTailIndex(sourceParts, { hostParagraphCount: raw.hostParagraphCount })
  });
  const detected = await detecting.inspect({ range: { start: 6, end: 7 }, includeDocumentEnd: true, limit: 2 });
  assert.deepEqual(detected.targets.map(target => target.text), ['Sixth paragraph']);
  assert.equal((await resolveOne(detecting, detected.documentEnd.handle)).target.descriptor.index, 6);
}

// 5. Store limits trigger and invalidate.
{
  const many = Array.from({ length: 130 }, (_, index) => `Paragraph number ${index + 1}`);
  const parts = documentParts(many);

  const { session: handleSession } = makeSession(() => parts, { maxResults: 12 });
  const early = await handleSession.inspect({ range: { start: 1, end: 12 }, limit: 12 });
  let limitError = null;
  let calls = 1;
  while (!limitError && calls < 60) {
    const next = await handleSession.inspect({ range: { start: 13, end: 24 }, limit: 12 });
    calls++;
    if (next.status === 'error') limitError = next.error;
  }
  assert.equal(limitError?.code, 'HANDLE_STORE_LIMIT');
  assert.ok(calls >= 33 && calls <= 35, `400 handles / 12 per page, got ${calls} calls`);
  assert.equal((await resolveOne(handleSession, early.targets[0].handle)).error.code, 'STALE_OR_UNKNOWN_HANDLE', 'hitting the limit expires earlier handles');
  const fresh = await handleSession.inspect({ limit: 1 });
  assert.equal(fresh.status, 'ok', 'the session recovers after the limit invalidates the store');

  const { session: cursorSession } = makeSession(() => parts);
  const firstPage = await cursorSession.inspect({ limit: 1 });
  const firstCursor = firstPage.selection.nextCursor;
  let cursorError = null;
  let cursorCalls = 1;
  while (!cursorError && cursorCalls < 150) {
    const next = await cursorSession.inspect({ limit: 1 });
    cursorCalls++;
    if (next.status === 'error') cursorError = next.error;
  }
  assert.equal(cursorError?.code, 'CURSOR_STORE_LIMIT');
  assert.equal(cursorCalls, 101, 'the 101st cursor-bearing page trips the 100-cursor limit');
  assert.equal((await cursorSession.inspect({ cursor: firstCursor })).error.code, 'STALE_CURSOR', 'hitting the limit expires earlier cursors');
  assert.equal((await resolveOne(cursorSession, firstPage.targets[0].handle)).error.code, 'STALE_OR_UNKNOWN_HANDLE');
}

// 6. Handle, cursor and session IDs are unique and opaque.
{
  const parts = documentParts(texts);
  const sessions = [makeSession(() => parts).session, makeSession(() => parts).session];
  const seen = new Set();
  const ids = [];
  for (const session of sessions) {
    for (let round = 0; round < 3; round++) {
      const page = await session.inspect({ limit: 2 });
      ids.push(...page.targets.map(target => target.handle), page.selection.nextCursor);
    }
  }
  for (const id of ids) {
    assert.equal(typeof id, 'string');
    assert.equal(seen.has(id), false, `duplicate id ${id}`);
    seen.add(id);
    for (const leak of ['Alpha', 'clause', 'paragraph', 'Second']) assert.equal(id.includes(leak), false, `${id} leaks text`);
  }
  const handleIds = ids.filter(id => !id.startsWith('cursor_'));
  assert.ok(handleIds.every(id => /^session_[0-9a-f-]{36}_h[0-9a-z]+$/.test(id)), 'handles are session-scoped opaque tokens');
  assert.ok(ids.filter(id => id.startsWith('cursor_')).every(id => /^cursor_[0-9a-f-]{36}$/.test(id)), 'cursors are random UUID tokens');
  const prefixes = new Set(handleIds.map(id => id.slice(0, id.indexOf('_h'))));
  assert.equal(prefixes.size, 2, 'each session has its own session ID');
  // Handle counters are sequential per session, so a handle never encodes paragraph position.
  const alpha = (await sessions[0].inspect({ search: 'Gamma' })).targets[0].handle;
  const sixth = (await sessions[0].inspect({ search: 'Sixth' })).targets[0].handle;
  assert.doesNotMatch(alpha.replace(/^session_[0-9a-f-]{36}/, ''), /5|six|gamma/i);
  assert.notEqual(alpha, sixth);
}

// KB-16: Word gives footnote and endnote paragraphs new w14:paraId/textId values on every
// export. That alone must not change the token; a change to note text still must.
{
  const notes = (root, paraId, text) => `<w:${root}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">`
    + `<w:${root} w:id="1"><w:p w14:paraId="${paraId}" w14:textId="77777777" w:rsidR="00C816A6"><w:r><w:t>${text}</w:t></w:r></w:p></w:${root}></w:${root}s>`;
  const withNotes = (paraId, text = 'Service providers are bound.', rsid) => ({
    ...documentParts(texts, rsid ? { rsid } : {}),
    additionalParts: { '/word/footnotes.xml': notes('footnote', paraId, text), '/word/endnotes.xml': notes('endnote', paraId, 'End.') }
  });
  const first = await computeWordSourceRevisionToken(withNotes('141AECF0'));
  const reexported = await computeWordSourceRevisionToken(withNotes('33A537D1', undefined, '00BE3784'));
  assert.equal(reexported.value, first.value, 'a re-export that only renumbers note paragraph IDs and rsids keeps the token');
  const edited = await computeWordSourceRevisionToken(withNotes('141AECF0', 'Service providers are not bound.'));
  assert.notEqual(edited.value, first.value, 'changed footnote text changes the token');
}

console.log('PASS: inspection session rsid-stable handles, content-change expiry, cursor continuation, document end, store limits and opaque IDs');
