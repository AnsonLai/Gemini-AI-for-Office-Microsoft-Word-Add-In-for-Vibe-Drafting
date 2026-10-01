import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
  initAgenticTools,
  executeInsertListItem,
  executeEditList,
  executeConvertHeadersToList
} from '../src/taskpane/modules/commands/agentic-tools.js';

initAgenticTools({
  getRequestSignal: () => null,
  loadApiKey: () => 'test-key',
  loadModel: () => 'test-model',
  loadSystemMessage: () => '',
  loadRedlineSetting: () => false,
  loadRedlineAuthor: () => 'Test Editor',
  setChangeTrackingForAi: async () => ({}),
  restoreChangeTracking: async () => {},
  SAFETY_SETTINGS_BLOCK_NONE: [],
  API_LIMITS: {}
});

const originalWord = globalThis.Word;

function createWordHarness(paragraphTexts) {
  const events = [];
  const list = {
    id: 'list-1',
    load(properties) { events.push({ type: 'list.load', properties }); },
    setLevelNumbering(level, format) { events.push({ type: 'list.numbering', level, format }); }
  };
  const paragraphs = paragraphTexts.map((text, index) => ({
    text,
    font: { name: 'Calibri' },
    load(properties) { events.push({ type: 'paragraph.load', index: index + 1, properties }); },
    insertText(value, location) {
      events.push({ type: 'paragraph.insertText', index: index + 1, value, location });
      this.text = value;
    },
    startNewList() {
      events.push({ type: 'paragraph.startNewList', index: index + 1 });
      return list;
    },
    attachToList(listId, level) {
      events.push({ type: 'paragraph.attachToList', index: index + 1, listId, level });
    },
    getOoxml() {
      events.push({ type: 'paragraph.getOoxml', index: index + 1 });
      return { value: '<w:p><w:r><w:t>Paragraph</w:t></w:r></w:p>' };
    },
    insertParagraph(value, location) {
      events.push({ type: 'paragraph.insertParagraph', index: index + 1, value, location });
      return { load() {}, getRange() { return { insertOoxml() { events.push({ type: 'range.insertOoxml', index: index + 1 }); } }; } };
    },
    getRange() {
      events.push({ type: 'paragraph.getRange', index: index + 1 });
      return {
        expandTo() { return this; },
        getOoxml() { return { value: '' }; },
        insertOoxml() { events.push({ type: 'range.insertOoxml', index: index + 1 }); }
      };
    }
  }));
  const collection = {
    items: paragraphs,
    load(properties) { events.push({ type: 'paragraphs.load', properties }); }
  };
  let runCount = 0;
  globalThis.Word = {
    InsertLocation: { replace: 'Replace' },
    ListNumbering: {
      arabic: 'arabic', lowerLetter: 'lowerLetter', upperLetter: 'upperLetter',
      lowerRoman: 'lowerRoman', upperRoman: 'upperRoman'
    },
    run: async callback => {
      runCount += 1;
      return callback({
        async sync() { events.push({ type: 'sync' }); },
        document: { body: { paragraphs: collection } }
      });
    }
  };
  return { events, paragraphs, get runCount() { return runCount; } };
}

async function testEditListPreflightRejectsInvalidArgumentsBeforeWord() {
  const harness = createWordHarness(['One.', 'Two.']);
  const result = await executeEditList(1, 2, ['One'], 'Bullet', 'decimal');

  assert.equal(result.success, false);
  assert.equal(result.writeAttempted, false);
  assert.equal(result.written, false);
  assert.equal(result.error.code, 'INVALID_LIST_REQUEST');
  assert.equal(harness.runCount, 0, 'invalid request should fail before opening Word.run');
  assert.equal(harness.events.some(event => event.type.includes('insert')), false);
}

async function testInsertListItemPreflightAndLiveRangeGuard() {
  const invalidHarness = createWordHarness(['One.']);
  const invalid = await executeInsertListItem(1, 'First\nSecond', 0);
  assert.equal(invalid.success, false);
  assert.equal(invalid.writeAttempted, false);
  assert.equal(invalid.error.code, 'INVALID_LIST_REQUEST');
  assert.equal(invalidHarness.runCount, 0, 'invalid inserted text should fail before Word.run');

  const staleHarness = createWordHarness(['One.', 'Two.']);
  const stale = await executeInsertListItem(3, 'New item', 0);
  assert.equal(stale.success, false);
  assert.equal(stale.writeAttempted, false);
  assert.equal(stale.written, false);
  assert.equal(stale.error.code, 'INVALID_LIST_REQUEST');
  assert.equal(staleHarness.events.some(event => event.type === 'paragraph.getOoxml'), false);
  assert.equal(staleHarness.events.some(event => event.type === 'paragraph.insertParagraph'), false);
}

async function testEditListRechecksIndexesAgainstLiveParagraphCount() {
  const harness = createWordHarness(['One.', 'Two.']);
  const result = await executeEditList(1, 3, ['First', 'Second'], 'bullet', 'decimal');

  assert.equal(result.success, false);
  assert.equal(result.writeAttempted, false);
  assert.equal(result.written, false);
  assert.equal(result.error.code, 'INVALID_LIST_REQUEST');
  assert.equal(harness.runCount, 2, 'font inspection and live-source preflight may read Word');
  assert.equal(harness.events.some(event => event.type.includes('insert')), false, 'stale range must be refused before range replacement');
}

async function testConvertHeadersKeepsTextPairedWithOriginalIndex() {
  const harness = createWordHarness(['1. First source', 'Body paragraph', '3. Third source']);
  const result = await executeConvertHeadersToList(
    [3, 1],
    ['Third replacement', 'First replacement'],
    'upperLetter'
  );

  assert.equal(result.success, true);
  assert.equal(result.written, true);
  assert.deepEqual(
    harness.events.filter(event => event.type === 'paragraph.insertText').map(({ index, value }) => ({ index, value })),
    [
      { index: 1, value: 'First replacement' },
      { index: 3, value: 'Third replacement' }
    ],
    'sorting paragraph indexes must not detach newHeaderTexts from their original targets'
  );
  assert.equal(harness.events.some(event => event.type === 'paragraph.insertText' && event.index === 2), false);
}

async function testConvertHeadersRejectsDuplicateTargetsBeforeWord() {
  const harness = createWordHarness(['1. First', '2. Second']);
  const result = await executeConvertHeadersToList([2, 2], ['First replacement', 'Conflicting replacement']);

  assert.equal(result.success, false);
  assert.equal(result.writeAttempted, false);
  assert.equal(result.written, false);
  assert.equal(result.error.code, 'INVALID_LIST_REQUEST');
  assert.equal(harness.runCount, 0, 'ambiguous duplicate targets should fail before opening Word.run');
  assert.equal(harness.events.some(event => event.type.includes('insert')), false);
}

try {
  await testEditListPreflightRejectsInvalidArgumentsBeforeWord();
  await testEditListRechecksIndexesAgainstLiveParagraphCount();
  await testInsertListItemPreflightAndLiveRangeGuard();
  await testConvertHeadersKeepsTextPairedWithOriginalIndex();
  await testConvertHeadersRejectsDuplicateTargetsBeforeWord();
  console.log('PASS: agentic list tool validation tests');
} catch (error) {
  console.error('FAIL:', error?.stack || error?.message || error);
  process.exitCode = 1;
} finally {
  globalThis.Word = originalWord;
}
