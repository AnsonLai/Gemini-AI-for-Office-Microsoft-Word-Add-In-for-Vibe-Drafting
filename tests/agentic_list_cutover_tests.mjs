import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  acceptTrackedChangesInOoxml,
  inspectDocumentParts,
  openDocx,
  rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';
import {
  initAgenticTools,
  executeInsertListItem
} from '../src/taskpane/modules/commands/agentic-tools.js';
import { planAgenticListOperations } from '../src/taskpane/modules/commands/list-operation-plan.js';
import { captureWordSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';

const FIXTURE_PATH = fileURLToPath(new URL('./fixtures/agentic-lists/nested-lists-source.docx', import.meta.url));
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const fixtureDoc = openDocx(readFileSync(FIXTURE_PATH));
const decodePart = name => {
  const part = fixtureDoc.entries.get(name);
  assert.ok(part, `Word-authored fixture is missing ${name}`);
  return new TextDecoder().decode(part).replace(/<\?xml[^>]*\?>/g, '');
};
const fixtureDocumentXml = decodePart('word/document.xml');
const fixtureNumberingXml = decodePart('word/numbering.xml');
const fixtureStylesXml = decodePart('word/styles.xml');

function flatOpcFromFixture(documentXml = fixtureDocumentXml) {
  const part = (name, contentType, xml) => `
    <pkg:part pkg:name="${name}" pkg:contentType="${contentType}">
      <pkg:xmlData>${xml}</pkg:xmlData>
    </pkg:part>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pkg:package xmlns:pkg="${NS_PKG}">
      ${part('/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', documentXml)}
      ${part('/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', fixtureNumberingXml)}
      ${part('/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml', fixtureStylesXml)}
    </pkg:package>`;
}

function bodyParagraphXml(documentXml) {
  const parsed = new DOMParser().parseFromString(documentXml, 'text/xml');
  const body = parsed.getElementsByTagNameNS(NS_W, 'body')[0];
  assert.ok(body, 'fixture document must have a Word body');
  const serializer = new XMLSerializer();
  return Array.from(body.childNodes)
    .filter(node => node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p')
    .map(node => serializer.serializeToString(node));
}

function addHistoricalListPropertiesToPlainParagraph(documentXml, paragraphText) {
  const parsed = new DOMParser().parseFromString(documentXml, 'text/xml');
  const body = parsed.getElementsByTagNameNS(NS_W, 'body')[0];
  const paragraphs = Array.from(body.childNodes).filter(node => node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p');
  const paragraph = paragraphs.find(node => Array.from(node.getElementsByTagNameNS(NS_W, 't')).map(text => text.textContent).join('') === paragraphText);
  assert.ok(paragraph, `Could not find plain fixture paragraph: ${paragraphText}`);

  let pPr = Array.from(paragraph.childNodes).find(node => node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'pPr');
  if (!pPr) {
    pPr = parsed.createElementNS(NS_W, 'w:pPr');
    paragraph.insertBefore(pPr, paragraph.firstChild);
  }
  const oldProperties = parsed.createElementNS(NS_W, 'w:pPr');
  const numPr = parsed.createElementNS(NS_W, 'w:numPr');
  const ilvl = parsed.createElementNS(NS_W, 'w:ilvl');
  ilvl.setAttribute('w:val', '0');
  const numId = parsed.createElementNS(NS_W, 'w:numId');
  numId.setAttribute('w:val', '1');
  numPr.appendChild(ilvl);
  numPr.appendChild(numId);
  oldProperties.appendChild(numPr);

  const pPrChange = parsed.createElementNS(NS_W, 'w:pPrChange');
  pPrChange.setAttribute('w:id', '91');
  pPrChange.setAttribute('w:author', 'Prior Editor');
  pPrChange.appendChild(oldProperties);
  pPr.appendChild(pPrChange);
  return new XMLSerializer().serializeToString(parsed.documentElement);
}

function sourceParagraphsFromFlatOpc(flatOpc) {
  return captureWordSourceBaseline(flatOpc);
}

function createWordHarness(flatOpc, { writeError = null } = {}) {
  const events = [];
  const sourceParagraphs = fixtureDoc.inspect().paragraphs;
  const directParagraphs = bodyParagraphXml(fixtureDocumentXml);
  const paragraphMocks = sourceParagraphs.map((source, index) => ({
    text: source.exactText,
    load(properties) { events.push({ type: 'paragraph.load', index: index + 1, properties }); },
    getOoxml() {
      events.push({ type: 'paragraph.getOoxml', index: index + 1 });
      return { value: directParagraphs[index] || '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>' };
    },
    insertParagraph(value, location) {
      events.push({ type: 'native.insertParagraph', index: index + 1, value, location });
      const listItem = { isNullObject: false };
      Object.defineProperty(listItem, 'level', {
        set(level) { events.push({ type: 'native.listLevel', level }); },
        get() { return 0; }
      });
      return {
        listItem,
        load(properties) { events.push({ type: 'inserted.load', properties }); },
        getRange(location) {
          events.push({ type: 'inserted.getRange', location });
          return { insertOoxml(xml, insertion) { events.push({ type: 'native.insertedRangeOoxml', xml, insertion }); } };
        }
      };
    }
  }));
  const paragraphs = {
    items: paragraphMocks,
    load(properties) { events.push({ type: 'paragraphs.load', properties }); }
  };
  const body = {
    paragraphs,
    getOoxml() {
      events.push({ type: 'body.getOoxml' });
      return { value: flatOpc };
    },
    insertOoxml(xml, location) {
      events.push({ type: 'body.insertOoxml', xml, location });
      if (writeError) throw writeError;
    }
  };
  let trackingMode = 'TrackAll';
  let runCount = 0;
  const document = {
    load(properties) { events.push({ type: 'document.load', properties }); },
    get changeTrackingMode() { return trackingMode; },
    set changeTrackingMode(value) { trackingMode = value; events.push({ type: 'trackingMode', value }); },
    body
  };
  globalThis.Word = {
    InsertLocation: { replace: 'Replace' },
    ChangeTrackingMode: { off: 'Off' },
    run: async callback => {
      runCount += 1;
      return callback({ document, async sync() { events.push({ type: 'sync' }); } });
    }
  };
  return { events, get runCount() { return runCount; } };
}

function initTools(redlineEnabled = true) {
  initAgenticTools({
    getRequestSignal: () => null,
    loadApiKey: () => 'list-cutover-test-key',
    loadModel: () => 'test-model',
    loadSystemMessage: () => '',
    loadRedlineSetting: () => redlineEnabled,
    loadRedlineAuthor: () => 'Cutover Test Editor',
    setChangeTrackingForAi: async () => ({}),
    restoreChangeTracking: async () => {},
    SAFETY_SETTINGS_BLOCK_NONE: [],
    API_LIMITS: {}
  });
}

function paragraphIndex(text, flatOpc = flatOpcFromFixture()) {
  const paragraph = fixtureDoc.inspect().paragraphs.find(item => item.exactText === text);
  assert.ok(paragraph, `Word-authored fixture is missing target ${text}`);
  const source = sourceParagraphsFromFlatOpc(flatOpc)[paragraph.index - 1];
  assert.ok(source, `Flat OPC is missing source target ${text}`);
  return paragraph.index;
}

function acceptedAndRejectedParagraphTexts(flatOpc) {
  const accepted = acceptTrackedChangesInOoxml(flatOpc, { allAuthors: true });
  const rejected = rejectTrackedChangesInOoxml(flatOpc, { allAuthors: true });
  assert.equal(accepted.status, undefined, JSON.stringify(accepted.error));
  assert.equal(rejected.status, undefined, JSON.stringify(rejected.error));
  return {
    accepted: sourceParagraphsFromFlatOpc(accepted.oxml).map(paragraph => paragraph.exactText),
    rejected: sourceParagraphsFromFlatOpc(rejected.oxml).map(paragraph => paragraph.exactText)
  };
}

async function testEligibleWordListsUseOneCanonicalBodyBatch() {
  initTools(true);
  const originalFlatOpc = flatOpcFromFixture();
  const originalTexts = sourceParagraphsFromFlatOpc(originalFlatOpc).map(paragraph => paragraph.exactText);
  for (const [targetText, indentLevel, insertedText, useBaseline] of [
    ['Bullet Root A', 0, 'Production bullet insertion', true],
    ['Number Nested Anchor', 0, 'Production decimal nested insertion', false]
  ]) {
    const harness = createWordHarness(originalFlatOpc);
    const index = paragraphIndex(targetText, originalFlatOpc);
    const result = await executeInsertListItem(
      index,
      insertedText,
      indentLevel,
      useBaseline ? sourceParagraphsFromFlatOpc(originalFlatOpc) : undefined
    );
    assert.equal(result.success, true, `${targetText}: ${result.message}`);
    assert.equal(result.mutationOutcome, 'applied');
    assert.equal(result.written, true);
    assert.equal(result.writeAttempted, true);
    assert.equal(result.receipts.length, 1, 'the engine receipt must flow through the production observer');
    assert.equal(harness.runCount, 1);
    assert.equal(harness.events.filter(event => event.type === 'body.getOoxml').length, 1,
      'the canonical route must read one immutable body Flat OPC snapshot');
    assert.equal(harness.events.filter(event => event.type === 'body.insertOoxml').length, 1,
      'the canonical route must commit the completed batch exactly once');
    assert.equal(harness.events.some(event => event.type === 'native.insertParagraph'), false,
      'an eligible canonical operation must not fall back to native insertion');
    assert.equal(harness.events.some(event => event.type === 'paragraph.getOoxml'), false,
      'the eligible route should not fetch per-paragraph OOXML');

    const insertedPackage = harness.events.find(event => event.type === 'body.insertOoxml').xml;
    const { accepted, rejected } = acceptedAndRejectedParagraphTexts(insertedPackage);
    const expectedAccepted = [...originalTexts];
    expectedAccepted.splice(index, 0, insertedText);
    assert.deepEqual(accepted, expectedAccepted, 'Accept All must contain exactly the added list item and preserve every source paragraph');
    assert.deepEqual(rejected, originalTexts, 'Reject All must reproduce the fixture source paragraph text exactly');
  }
}

async function testStaleAndUnavailableProvidedBaselinesRefuseBeforeNativeOrEngineWrites() {
  initTools(true);
  const sourceBaseline = sourceParagraphsFromFlatOpc(flatOpcFromFixture());
  const staleDocumentXml = fixtureDocumentXml.replace('Bullet Root A', 'Concurrent bullet root A');
  assert.notEqual(staleDocumentXml, fixtureDocumentXml, 'the fixture mutation must change the targeted source text');
  for (const baseline of [sourceBaseline, []]) {
    const flatOpc = baseline.length === 0
      ? flatOpcFromFixture()
      : flatOpcFromFixture(staleDocumentXml);
    const harness = createWordHarness(flatOpc);
    const result = await executeInsertListItem(
      paragraphIndex('Bullet Root A', flatOpcFromFixture()),
      'Must not overwrite stale context',
      0,
      baseline
    );
    assert.equal(result.success, false);
    assert.equal(result.error.code, 'STALE_DOCUMENT_CONTEXT');
    assert.equal(result.mutationOutcome, 'refused');
    assert.equal(result.writeAttempted, false);
    assert.equal(result.written, false);
    assert.equal(harness.events.filter(event => event.type === 'body.getOoxml').length, 1);
    assert.equal(harness.events.filter(event => event.type === 'body.insertOoxml').length, 0);
    assert.equal(harness.events.some(event => event.type === 'native.insertParagraph'), false,
      'a stale or unavailable provided baseline must not fall through to native mutation');
  }
}

async function testOrdinaryPlainParagraphRetainsNativeFallback() {
  initTools(true);
  const targetText = 'Plain paragraph before bullet list.';
  const flatOpc = flatOpcFromFixture();
  const harness = createWordHarness(flatOpc);
  const result = await executeInsertListItem(paragraphIndex(targetText, flatOpc), 'Keep this as a plain paragraph', 0);

  assert.equal(result.success, true, result.message);
  assert.equal(result.written, true);
  assert.equal(harness.events.filter(event => event.type === 'body.getOoxml').length, 1,
    'the source is inspected before choosing the native path');
  assert.equal(harness.events.filter(event => event.type === 'body.insertOoxml').length, 0,
    'a plain paragraph does not enter the canonical list engine');
  assert.equal(harness.events.filter(event => event.type === 'native.insertParagraph').length, 1,
    'an ordinary plain paragraph retains the established native insertion behavior');
}

async function testHistoricalPPrChangeListBindingRefusesWithoutWrites() {
  initTools(true);
  const targetText = 'Plain paragraph before bullet list.';
  const modifiedDocumentXml = addHistoricalListPropertiesToPlainParagraph(fixtureDocumentXml, targetText);
  const inspected = inspectDocumentParts({ documentXml: modifiedDocumentXml, numberingXml: fixtureNumberingXml, stylesXml: fixtureStylesXml });
  assert.deepEqual(inspected.paragraphs.find(item => item.exactText === targetText)?.list, {
    numId: '1', level: 0, label: '', format: 'bullet'
  }, 'the installed library currently treats historical pPrChange numPr as active list binding');
  const flatOpc = flatOpcFromFixture(modifiedDocumentXml);
  const sourceParagraphs = sourceParagraphsFromFlatOpc(flatOpc);
  const target = sourceParagraphs.find(item => item.exactText === targetText);
  const source = {
    documentXml: modifiedDocumentXml,
    paragraphs: sourceParagraphs,
    numberingXml: fixtureNumberingXml,
    stylesXml: fixtureStylesXml
  };
  const operations = planAgenticListOperations(source, {
    kind: 'insert_list_item', afterParagraphIndex: target.index,
    text: 'Keep this as a plain paragraph', indentLevel: 0
  });
  const engineResult = await applyOperationsToDocumentXml(
    modifiedDocumentXml,
    operations,
    'Historical list inspection test',
    { numberingXml: fixtureNumberingXml, stylesXml: fixtureStylesXml },
    { atomic: true, structuredContent: true, pairReplacements: true, generateRedlines: true }
  );
  assert.equal(engineResult.error.code, 'BATCH_OPERATION_FAILED');
  assert.equal(engineResult.results[0].error.code, 'EXISTING_REVISIONS');
  assert.equal(engineResult.results[0].receipt.finalDisposition, 'refused');

  const harness = createWordHarness(flatOpc);
  const result = await executeInsertListItem(target.index, 'Keep this as a plain paragraph', 0);

  assert.equal(result.success, false);
  assert.equal(result.error.code, 'BATCH_OPERATION_FAILED');
  assert.equal(result.operationResults[0].error.code, 'BATCH_OPERATION_FAILED');
  assert.equal(result.receipts[0].finalDisposition, 'refused');
  assert.equal(result.writeAttempted, false);
  assert.equal(result.written, false);
  assert.equal(harness.events.filter(event => event.type === 'body.getOoxml').length, 1,
    'the production cutover reads one immutable source snapshot');
  assert.equal(harness.events.filter(event => event.type === 'body.insertOoxml').length, 0,
    'the library refusal must not be committed');
  assert.equal(harness.events.some(event => event.type === 'native.insertParagraph'), false,
    'a canonical-path refusal must not replay a possibly mis-targeted native insertion');
}

async function testUnsupportedOutdentUsesEstablishedNativePath() {
  initTools(true);
  const flatOpc = flatOpcFromFixture();
  const harness = createWordHarness(flatOpc);
  const index = paragraphIndex('Number Restart Nested', flatOpc);
  const result = await executeInsertListItem(index, 'Native fallback for unsupported outdent', -1);

  assert.equal(result.success, true);
  assert.equal(result.written, true);
  assert.equal(harness.events.filter(event => event.type === 'body.getOoxml').length, 1,
    'the unsupported target is inspected before the fallback decision');
  assert.equal(harness.events.filter(event => event.type === 'body.insertOoxml').length, 0,
    'unsupported outdent mapping must return an empty batch, not write through the engine');
  assert.equal(harness.events.filter(event => event.type === 'native.insertParagraph').length, 1,
    'unsupported outdent mapping retains the previously supported native path');
}

async function testCanonicalHostFailureDoesNotFallBackToNativeWrites() {
  initTools(true);
  const flatOpc = flatOpcFromFixture();
  const hostError = Object.assign(new Error('Future Word host failure'), { code: 'FUTURE_WORD_ERROR' });
  const harness = createWordHarness(flatOpc, { writeError: hostError });
  const index = paragraphIndex('Bullet Root A', flatOpc);
  const result = await executeInsertListItem(index, 'Host failure must not replay natively', 0);

  assert.equal(result.success, false);
  assert.equal(result.status, 'error');
  assert.equal(result.mutationOutcome, 'indeterminate');
  assert.equal(result.writeAttempted, true);
  assert.equal(result.written, false);
  assert.equal(result.operationResults[0].hostError.code, 'FUTURE_WORD_ERROR');
  assert.equal(harness.events.filter(event => event.type === 'body.getOoxml').length, 1);
  assert.equal(harness.events.filter(event => event.type === 'body.insertOoxml').length, 1);
  assert.equal(harness.events.some(event => event.type === 'native.insertParagraph'), false,
    'a host-uncertain canonical attempt must never be replayed through the native fallback');
}

async function testLibrarySourceRefusalDoesNotFallBackToNativeWrites() {
  initTools(true);
  // A malformed Word package makes the shared source inspector refuse before
  // the operation factory; the executor must still stop without a native write.
  const harness = createWordHarness('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part');
  const index = paragraphIndex('Bullet Root A');
  const result = await executeInsertListItem(index, 'No fallback after source refusal', 0);

  assert.equal(result.success, false);
  assert.equal(result.writeAttempted, false);
  assert.equal(result.written, false);
  assert.equal(harness.events.filter(event => event.type === 'body.getOoxml').length, 1);
  assert.equal(harness.events.filter(event => event.type === 'body.insertOoxml').length, 0);
  assert.equal(harness.events.some(event => event.type === 'native.insertParagraph'), false,
    'a source inspection refusal must not enter the native path');
}

const previousWord = globalThis.Word;
try {
  await testEligibleWordListsUseOneCanonicalBodyBatch();
  await testStaleAndUnavailableProvidedBaselinesRefuseBeforeNativeOrEngineWrites();
  await testOrdinaryPlainParagraphRetainsNativeFallback();
  await testHistoricalPPrChangeListBindingRefusesWithoutWrites();
  await testUnsupportedOutdentUsesEstablishedNativePath();
  await testCanonicalHostFailureDoesNotFallBackToNativeWrites();
  await testLibrarySourceRefusalDoesNotFallBackToNativeWrites();
  console.log('PASS: agentic insert-list-item production cutover tests');
} catch (error) {
  console.error('FAIL:', error?.stack || error?.message || error);
  process.exitCode = 1;
} finally {
  globalThis.Word = previousWord;
}
