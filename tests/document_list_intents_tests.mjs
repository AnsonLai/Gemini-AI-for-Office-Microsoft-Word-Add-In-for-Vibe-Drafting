// End-to-end coverage for the typed list intents (convert_to_list, change_list_format,
// continue_list) through the production intent compiler, document tools and the real
// docx-redline-js library. Fixtures carry a real numbering.xml so formats and labels
// resolve. Reject All and Accept All are taken from the committed document.
import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { acceptTrackedChangesInOoxml, openDocx, rejectTrackedChangesInOoxml } from '@ansonlai/docx-redline-js';
import { inspectWordSource } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';
import { createDocumentTools } from '../src/taskpane/modules/chat/document-tools.js';
import { createFunctionResponse } from '../src/taskpane/modules/chat/document-agent-loop.js';
import { scenario as directScenario } from '../scripts/golden/nda-direct-scenario.mjs';
import { buildDocumentModel, evaluateCheck, evaluateDirectToolEvidence } from '../scripts/golden/golden-checks.mjs';
import { resolveReferences } from '../scripts/golden/golden-replay.mjs';

globalThis.Word = {
  ChangeTrackingMode: { off: 'Off', trackAll: 'TrackAll' },
  InsertLocation: { replace: 'Replace', Replace: 'Replace' }
};

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const escapeXml = value => String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');

// ---------------------------------------------------------------- fixtures
const levels = (format, text) => Array.from({ length: 9 }, (_, i) =>
  `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text(i)}"/><w:lvlJc w:val="left"/>`
  + `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`).join('');
const decimalText = i => Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.') + '.';
// num 1 and num 3 are separate decimal lists; num 2 is a bullet list; num 4 backs the "Numbered" paragraph style.
const NUMBERING = `<w:numbering xmlns:w="${W}">`
  + `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${levels('decimal', decimalText)}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${levels('bullet', () => '•')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/>${levels('decimal', decimalText)}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="multilevel"/>${levels('decimal', decimalText)}</w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
  + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
  + '<w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>'
  + '<w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num>'
  + '</w:numbering>';
const STYLES = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`
  + '<w:style w:type="paragraph" w:styleId="Numbered"><w:name w:val="Numbered"/><w:basedOn w:val="Normal"/>'
  + '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr></w:style></w:styles>';

let paraCounter = 0;
function para(text, { id, numId, level = 0, style } = {}) {
  const paraId = id || `1A00${String(++paraCounter).padStart(4, '0')}`;
  const pPr = style || numId
    ? `<w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ''}${numId ? `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>` : ''}</w:pPr>`
    : '';
  return `<w:p w14:paraId="${paraId}">${pPr}<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}
const part = (name, type, xml) => `<pkg:part pkg:name="/${name}" pkg:contentType="${type}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
function flatOpc(paragraphs) {
  return flatOpcFromParts(`<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraphs.join('')}</w:body></w:document>`, NUMBERING, STYLES);
}
function flatOpcFromParts(documentXml, numberingXml, stylesXml) {
  const strip = xml => xml.replace(/<\?xml[^>]*\?>/g, '');
  const rel = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`
    + part('_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>`)
    + part('word/_rels/document.xml.rels', 'application/vnd.openxmlformats-package.relationships+xml',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'numbering', 'numbering.xml')}${rel('rId2', 'styles', 'styles.xml')}</Relationships>`)
    + part('word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', strip(documentXml))
    + part('word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', strip(numberingXml))
    + part('word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml', strip(stylesXml))
    + '</pkg:package>';
}

// ---------------------------------------------------------------- harness
function createHarness(initialXml, { generateRedlines = true } = {}) {
  const counts = { contentWrites: 0 };
  let currentXml = initialXml;
  let tracking = Word.ChangeTrackingMode.trackAll;
  const body = {
    getOoxml() { return { value: currentXml }; },
    insertOoxml(payload) { counts.contentWrites++; currentXml = payload; }
  };
  const document = { body, load() {} };
  Object.defineProperty(document, 'changeTrackingMode', { get: () => tracking, set: value => { tracking = value; } });
  const context = { document, async sync() {} };
  const session = createDocumentInspectionSession({
    readSource: async () => {
      const source = inspectWordSource(currentXml);
      if (source.inspection?.status !== 'ok') throw new Error(`Inspection failed: ${JSON.stringify(source.inspection?.error)}`);
      return { documentId: 'list-intents-doc', revisionView: 'accepted', parts: source.parts };
    }
  });
  const receipts = [];
  const tools = createDocumentTools({
    session,
    getDocumentId: () => 'list-intents-doc',
    getAuthor: () => 'List Intents Test',
    getGenerateRedlines: () => generateRedlines,
    persistCheckpoint: async () => 1,
    run: async callback => callback(context),
    onReceipt: receipt => receipts.push(receipt)
  });
  return {
    counts, receipts, tools,
    get currentXml() { return currentXml; },
    /** Handle for the one paragraph whose accepted text starts with `prefix`. */
    async handle(prefix) {
      const result = await tools.inspectDocument({ search: prefix, limit: 8 });
      assert.equal(result.status, 'ok', JSON.stringify(result));
      const matches = result.targets.filter(item => item.text.startsWith(prefix));
      assert.equal(matches.length, 1, `exactly one paragraph starts with ${JSON.stringify(prefix)}`);
      return matches[0].handle;
    },
    apply: edits => tools.applyDocumentEdits({ edits })
  };
}

const view = (xml, action) => {
  const result = action === 'accept'
    ? acceptTrackedChangesInOoxml(xml, { allAuthors: true })
    : rejectTrackedChangesInOoxml(xml, { allAuthors: true });
  assert.equal(result.status, undefined, JSON.stringify(result.error));
  return inspectWordSource(result.oxml).inspection.paragraphs;
};
const shape = paragraphs => paragraphs.map(item => [item.exactText, item.list ? item.list.label : null]);
const formats = paragraphs => paragraphs.map(item => (item.list ? item.list.format : null));
const ids = paragraphs => paragraphs.map(item => item.paragraphId);

function assertCommitted(receipt, harness, label) {
  assert.equal(receipt.success, true, `${label}: ${JSON.stringify(receipt)}`);
  assert.equal(receipt.verified, true, label);
  assert.equal(receipt.written, true, label);
  assert.equal(harness.counts.contentWrites, 1, `${label}: one host write`);
}
function assertRefusedWithoutWrite(receipt, harness, before, code, label) {
  assert.equal(receipt.success, false, `${label}: ${JSON.stringify(receipt)}`);
  assert.equal(receipt.written, false, label);
  assert.ok(['refused', 'rolled_back'].includes(receipt.mutationOutcome), `${label}: ${receipt.mutationOutcome}`);
  assert.equal(harness.counts.contentWrites, 0, `${label}: no host write`);
  assert.equal(harness.currentXml, before, `${label}: document unchanged`);
  // The library's per-operation code and recovery reach the model; the atomic wrapper alone is not actionable.
  const failure = receipt.error?.failedOperations?.[0];
  assert.equal(failure?.code, code, `${label}: library code passes through (${JSON.stringify(receipt.error)})`);
  assert.equal(failure.editIndex, 0, label);
  assert.ok(failure.recovery?.hint, `${label}: library recovery hint passes through`);
  return failure;
}

// ---------------------------------------------------------------- (a) non-contiguous headers become one list
const HEADERS = [['1. DEFINITIONS', 'Body about definitions.'], ['2. TERM', 'Body about term.'], ['3. FEES', 'Closing paragraph.']];
function headerSource() {
  return flatOpc([para('Intro paragraph.'), ...HEADERS.flatMap(([header, body]) => [para(header), para(body)])]);
}
for (const mode of ['one convert per header', 'one contiguous range']) {
  const source = headerSource();
  const harness = createHarness(source);
  const sourceParagraphs = inspectWordSource(source).inspection.paragraphs;
  const edits = mode === 'one convert per header'
    ? await Promise.all(HEADERS.map(async ([header]) => ({ type: 'convert_to_list', target: await harness.handle(header) })))
    : [{ type: 'convert_to_list', target: await harness.handle('1. DEFINITIONS'), through: await harness.handle('3. FEES') }];
  const receipt = await harness.apply(edits);
  assertCommitted(receipt, harness, `(a) ${mode}`);

  const accepted = view(harness.currentXml, 'accept');
  assert.deepEqual(shape(accepted), [
    ['Intro paragraph.', null], ['DEFINITIONS', '1.'], ['Body about definitions.', null],
    ['TERM', '2.'], ['Body about term.', null], ['FEES', '3.'], ['Closing paragraph.', null]
  ], `(a) ${mode}: Accept All shows labels 1. 2. 3. with marker-free text`);
  assert.deepEqual(new Set(accepted.filter(item => item.list).map(item => item.list.numId)).size, 1, `(a) ${mode}: one shared list`);
  assert.deepEqual(new Set(formats(accepted).filter(Boolean)), new Set(['decimal']));
  assert.deepEqual(ids(accepted), ids(sourceParagraphs), `(a) ${mode}: w14:paraId preserved for every paragraph`);

  const rejected = view(harness.currentXml, 'reject');
  assert.deepEqual(shape(rejected), shape(sourceParagraphs.map(item => ({ ...item, list: null }))),
    `(a) ${mode}: Reject All restores the literal typed text exactly`);
  assert.deepEqual(rejected.map(item => item.exactText), sourceParagraphs.map(item => item.exactText));
  assert.deepEqual(ids(rejected), ids(sourceParagraphs));
}

// ---------------------------------------------------------------- (b) letter headers keep their typed start value
{
  const source = flatOpc([para('Preamble.'), para('C. Scope'), para('Scope body.'), para('D. Fees'), para('Fees body.')]);
  const harness = createHarness(source);
  const receipt = await harness.apply([
    { type: 'convert_to_list', target: await harness.handle('C. Scope'), format: 'upperLetter' },
    { type: 'convert_to_list', target: await harness.handle('D. Fees'), format: 'upperLetter' }
  ]);
  assertCommitted(receipt, harness, '(b)');
  const accepted = view(harness.currentXml, 'accept');
  assert.deepEqual(shape(accepted), [
    ['Preamble.', null], ['Scope', 'C.'], ['Scope body.', null], ['Fees', 'D.'], ['Fees body.', null]
  ], '(b) the typed start value C. is the visible label');
  assert.deepEqual(formats(accepted).filter(Boolean), ['upperLetter', 'upperLetter']);
  assert.deepEqual(shape(view(harness.currentXml, 'reject')).map(row => row[0]),
    ['Preamble.', 'C. Scope', 'Scope body.', 'D. Fees', 'Fees body.']);

  // markers:none keeps the text and uses format and start for the label.
  const plain = flatOpc([para('Preamble.'), para('Scope'), para('Fees')]);
  const plainHarness = createHarness(plain);
  const plainReceipt = await plainHarness.apply([
    { type: 'convert_to_list', target: await plainHarness.handle('Scope'), through: await plainHarness.handle('Fees'), markers: 'none', format: 'upperLetter', start: 3 }
  ]);
  assertCommitted(plainReceipt, plainHarness, '(b) markers none');
  assert.deepEqual(shape(view(plainHarness.currentXml, 'accept')), [['Preamble.', null], ['Scope', 'C.'], ['Fees', 'D.']]);
  assert.deepEqual(shape(view(plainHarness.currentXml, 'reject')), [['Preamble.', null], ['Scope', null], ['Fees', null]]);
}

// ---------------------------------------------------------------- (c) bullets to decimal
{
  const source = flatOpc([
    para('Lead-in.'), para('Passport', { numId: 2 }), para('Driver licence', { numId: 2 }), para('Health card', { numId: 2 }), para('Trailer.')
  ]);
  const harness = createHarness(source);
  const before = inspectWordSource(source).inspection.paragraphs;
  const receipt = await harness.apply([{ type: 'change_list_format', target: await harness.handle('Driver licence'), format: 'decimal' }]);
  assertCommitted(receipt, harness, '(c)');
  const accepted = view(harness.currentXml, 'accept');
  assert.deepEqual(shape(accepted), [
    ['Lead-in.', null], ['Passport', '1.'], ['Driver licence', '2.'], ['Health card', '3.'], ['Trailer.', null]
  ], '(c) the whole list becomes decimal, including items before the target');
  assert.deepEqual(formats(accepted), [null, 'decimal', 'decimal', 'decimal', null]);
  const rejected = view(harness.currentXml, 'reject');
  assert.deepEqual(shape(rejected), shape(before));
  assert.deepEqual(formats(rejected), [null, 'bullet', 'bullet', 'bullet', null], '(c) Reject All restores the bullets');
  assert.deepEqual(rejected.map(item => item.list?.numId), before.map(item => item.list?.numId));
  assert.deepEqual(ids(accepted), ids(before));
}

// ---------------------------------------------------------------- (d) a later list continues an earlier one
{
  const source = flatOpc([
    para('Alpha', { numId: 1 }), para('Beta', { numId: 1 }), para('Interlude.'), para('Gamma', { numId: 3 }), para('Delta', { numId: 3 })
  ]);
  const harness = createHarness(source);
  const before = inspectWordSource(source).inspection.paragraphs;
  assert.deepEqual(shape(before).map(row => row[1]), ['1.', '2.', null, '1.', '2.'], 'fixture: two separate decimal lists');
  const receipt = await harness.apply([{ type: 'continue_list', target: await harness.handle('Gamma'), from: await harness.handle('Beta') }]);
  assertCommitted(receipt, harness, '(d)');
  assert.deepEqual(shape(view(harness.currentXml, 'accept')), [
    ['Alpha', '1.'], ['Beta', '2.'], ['Interlude.', null], ['Gamma', '3.'], ['Delta', '4.']
  ], '(d) the second list continues the first');
  assert.deepEqual(shape(view(harness.currentXml, 'reject')), shape(before), '(d) Reject All restores both lists');
}

// ---------------------------------------------------------------- (e) changed-text header plus conversion in one batch
for (const order of ['rewrite first', 'convert first']) {
  const source = flatOpc([para('Lead-in.'), para('1. Old Terms'), para('Body.')]);
  const harness = createHarness(source);
  const target = await harness.handle('1. Old Terms');
  const rewrite = { type: 'rewrite_paragraph', target, text: '1. New Terms' };
  const convert = { type: 'convert_to_list', target };
  const receipt = await harness.apply(order === 'rewrite first' ? [rewrite, convert] : [convert, rewrite]);
  assertCommitted(receipt, harness, `(e) ${order}`);
  assert.deepEqual(shape(view(harness.currentXml, 'accept')), [['Lead-in.', null], ['New Terms', '1.'], ['Body.', null]],
    `(e) ${order}: label 1. with the new text`);
  assert.deepEqual(shape(view(harness.currentXml, 'reject')), [['Lead-in.', null], ['1. Old Terms', null], ['Body.', null]],
    `(e) ${order}: Reject All restores the literal old header`);
}

// ---------------------------------------------------------------- (f) library refusals pass through without a write
{
  const mixed = flatOpc([para('1. Alpha'), para('A. Beta')]);
  const harness = createHarness(mixed);
  const receipt = await harness.apply([{ type: 'convert_to_list', target: await harness.handle('1. Alpha'), through: await harness.handle('A. Beta') }]);
  assertRefusedWithoutWrite(receipt, harness, mixed, 'LIST_MARKER_UNSUPPORTED', '(f) mixed markers');
}
{
  const ambiguous = flatOpc([para('I. Introduction'), para('Body.')]);
  const harness = createHarness(ambiguous);
  const receipt = await harness.apply([{ type: 'convert_to_list', target: await harness.handle('I. Introduction') }]);
  const failure = assertRefusedWithoutWrite(receipt, harness, ambiguous, 'LIST_MARKER_AMBIGUOUS', '(f) ambiguous I.');
  assert.deepEqual(failure.recovery.allowedValues, ['upperLetter', 'upperRoman']);
  const resolved = createHarness(ambiguous);
  const ok = await resolved.apply([{ type: 'convert_to_list', target: await resolved.handle('I. Introduction'), format: 'upperRoman' }]);
  assertCommitted(ok, resolved, '(f) I. with format');
  assert.deepEqual(shape(view(resolved.currentXml, 'accept')), [['Introduction', 'I.'], ['Body.', null]]);
}
{
  const styled = flatOpc([para('Lead-in.'), para('Styled item', { style: 'Numbered' }), para('1. Typed header')]);
  const harness = createHarness(styled);
  // The inspector reports no direct list for a style-numbered paragraph, so only the library can refuse it.
  assert.equal(inspectWordSource(styled).inspection.paragraphs[1].list, null);
  const receipt = await harness.apply([{ type: 'convert_to_list', target: await harness.handle('Styled item'), markers: 'none', format: 'decimal' }]);
  assertRefusedWithoutWrite(receipt, harness, styled, 'LIST_FORMAT_STYLE_NUMBERED', '(f) style-numbered');
}

// ---------------------------------------------------------------- (g) golden step 12 on the real sample NDA
// The direct golden replay's own calls run through the production tools and are scored with the
// scenario's own checks. Headers are NormalWeb paragraphs between existing lists.
{
  const step = directScenario.steps.find(item => item.id === '12-headers-to-list');
  assert.ok(step.replay && !step.scopeDisposition, 'step 12 is a replay, not a scripted refusal');
  const ndaBytes = new Uint8Array(readFileSync(new URL('./fixtures/golden/sample-nda.docx', import.meta.url)));
  const nda = openDocx(ndaBytes);
  const entryText = name => new TextDecoder().decode(nda.entries.get(name));
  const source = flatOpcFromParts(entryText('word/document.xml'), entryText('word/numbering.xml'), entryText('word/styles.xml'));
  const before = inspectWordSource(source).inspection.paragraphs;
  const harness = createHarness(source);

  const inspectParts = step.replay[0].chat;
  assert.ok(Array.isArray(inspectParts) && inspectParts.length === 9, 'one parallel inspection turn covers the nine headers');
  const responses = [];
  for (const part of inspectParts) {
    assert.equal(part.functionCall.name, 'inspect_document');
    responses.push(createFunctionResponse('inspect_document', JSON.stringify(await harness.tools.inspectDocument(part.functionCall.args))));
  }
  const payload = { contents: [
    { role: 'user', parts: [{ text: `User Question:\n${step.prompt}` }] },
    { role: 'model', parts: inspectParts }, { role: 'user', parts: responses }
  ] };
  const applyCall = resolveReferences(step.replay[1].chat, [], payload, step.prompt).functionCall;
  assert.equal(applyCall.name, 'apply_document_edits');
  assert.equal(applyCall.args.edits.length, 9);
  assert.ok(applyCall.args.edits.every(edit => edit.type === 'convert_to_list' && typeof edit.target === 'string'));
  const receipt = await harness.tools.applyDocumentEdits(applyCall.args);
  assertCommitted(receipt, harness, '(g) NDA headers');
  assert.equal(receipt.accountedCount, 9);

  const parts = inspectWordSource(harness.currentXml).parts;
  const scored = openDocx(ndaBytes);
  scored.entries.set('word/document.xml', new TextEncoder().encode(parts.documentXml));
  scored.entries.set('word/numbering.xml', new TextEncoder().encode(parts.numberingXml));
  const model = await buildDocumentModel(scored.toUint8Array());
  for (const check of step.checks.filter(item => item.type !== 'directToolEvidence')) {
    const outcome = evaluateCheck(model, check);
    assert.equal(outcome.ok, true, `(g) ${check.label}: ${outcome.detail}`);
  }
  const evidence = evaluateDirectToolEvidence({
    served: [{ kind: 'chat', part: { functionCall: applyCall } }],
    functionResponseHistory: { apply_document_edits: [receipt] }, requestCounts: { chat: 3, aux: 0, other: 0 }
  });
  assert.equal(evidence.ok, true, evidence.detail);

  const accepted = view(harness.currentXml, 'accept');
  const headerRows = accepted.filter(item => item.list && /^[A-Z][A-Z ]+$/.test(item.exactText) && item.list.level === 0 && item.list.format === 'decimal'
    && item.paragraphId && before.some(old => old.paragraphId === item.paragraphId && /^[1-9]\. /.test(old.exactText)));
  assert.deepEqual(headerRows.map(item => item.list.label), ['1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.'], 'nine headers carry labels 1. to 9.');
  assert.equal(new Set(headerRows.map(item => item.list.numId)).size, 1, 'the nine headers share one list');
  const rejected = view(harness.currentXml, 'reject');
  assert.deepEqual(rejected.map(item => [item.exactText, item.list ? [item.list.numId, item.list.level, item.list.label] : null]),
    before.map(item => [item.exactText, item.list ? [item.list.numId, item.list.level, item.list.label] : null]),
    '(g) Reject All restores every paragraph and list binding of the sample NDA');
  assert.deepEqual(ids(rejected), ids(before));
}

// ---------------------------------------------------------------- (h) Word-authored list fixture
{
  const fixture = openDocx(new Uint8Array(readFileSync(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url))));
  const text = name => new TextDecoder().decode(fixture.entries.get(name));
  const source = flatOpcFromParts(text('word/document.xml'), text('word/numbering.xml'), text('word/styles.xml'));
  const before = inspectWordSource(source).inspection.paragraphs;
  const run = async (edits) => {
    const harness = createHarness(source);
    const resolved = [];
    for (const edit of edits) {
      const handles = {};
      for (const key of ['target', 'through', 'from']) if (edit[key]) handles[key] = await harness.handle(edit[key]);
      resolved.push({ ...edit, ...handles });
    }
    return { harness, receipt: await harness.apply(resolved) };
  };
  const labels = paragraphs => paragraphs.map(item => [item.exactText, item.list ? `${item.list.format}/L${item.list.level}${item.list.format === 'bullet' ? '' : `/${item.list.label}`}` : null]);

  const converted = await run([{ type: 'convert_to_list', target: '1. Supported Header First' }, { type: 'convert_to_list', target: '2. Supported Header Second' }]);
  assertCommitted(converted.receipt, converted.harness, '(h) headers');
  const acceptedHeaders = labels(view(converted.harness.currentXml, 'accept'));
  assert.deepEqual(acceptedHeaders.slice(-4), [
    ['Supported Header First', 'decimal/L0/1.'], ['Plain paragraph between supported headers.', null],
    ['Supported Header Second', 'decimal/L0/2.'], ['Plain paragraph after supported headers.', null]
  ], '(h) the two Word-authored typed headers become a decimal list; unmarked candidates stay plain');
  assert.deepEqual(acceptedHeaders.slice(0, 16), labels(before).slice(0, 16), '(h) every earlier list keeps its source labels');
  assert.deepEqual(labels(view(converted.harness.currentXml, 'reject')), labels(before), '(h) Reject All restores the source');

  const reformatted = await run([{ type: 'change_list_format', target: 'Bullet Root A', format: 'decimal' }]);
  assertCommitted(reformatted.receipt, reformatted.harness, '(h) bullets');
  const acceptedBullets = labels(view(reformatted.harness.currentXml, 'accept'));
  assert.deepEqual(acceptedBullets.slice(2, 6), [
    ['Bullet Root A', 'decimal/L0/1.'], ['Bullet Insertion Anchor', 'bullet/L1'],
    ['Bullet Untouched Tail', 'decimal/L0/2.'], ['Bullet Root B', 'decimal/L0/3.']
  ], '(h) the library redefines only the target level; the nested bullet level is untouched');
  assert.deepEqual(labels(view(reformatted.harness.currentXml, 'reject')), labels(before), '(h) Reject All restores every bullet');

  const joined = await run([{ type: 'continue_list', target: 'Number Restart Root', from: 'Number Root B' }]);
  assertCommitted(joined.receipt, joined.harness, '(h) join');
  const acceptedJoin = labels(view(joined.harness.currentXml, 'accept'));
  assert.deepEqual(acceptedJoin.filter(([name]) => name.startsWith('Number Restart')), [['Number Restart Root', 'decimal/L0/4.'], ['Number Restart Nested', 'decimal/L1/4.1.']],
    '(h) the restarted list continues after the last item of the earlier list');
  assert.deepEqual(labels(view(joined.harness.currentXml, 'reject')), labels(before), '(h) Reject All restores the restart');
}

console.log('PASS: list intents run end to end through the library');
