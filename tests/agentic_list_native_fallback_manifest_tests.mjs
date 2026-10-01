import assert from 'node:assert/strict';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fixtureDirectory = path.join(root, 'tests/fixtures/agentic-lists');
const sourceObservations = JSON.parse(readFileSync(path.join(fixtureDirectory, 'source-observations.json'), 'utf8'));
const sourceTexts = sourceObservations.paragraphs.map(paragraph => paragraph.text);
assert.equal(sourceObservations.provenance.authoredBy, 'Microsoft Word desktop COM');
assert.equal(sourceObservations.provenance.verifiedAfterSaveReopen, true);
assert.equal(sourceObservations.paragraphCount, sourceTexts.length);

const sourceText = sourceTexts.join('\n');
const insertAfter = (anchorText, insertedText) => {
  const anchorIndex = sourceTexts.indexOf(anchorText);
  assert.notEqual(anchorIndex, -1, `Source fixture is missing paragraph: ${anchorText}`);
  return sourceTexts.toSpliced(anchorIndex + 1, 0, insertedText).join('\n');
};
const observation = text => {
  const item = sourceObservations.paragraphs.find(paragraph => paragraph.text === text);
  assert.ok(item, `Missing source list observation: ${text}`);
  return item;
};
const groupFor = text => text.startsWith('Bullet ')
  ? 'bullet-chain'
  : ['Number Root A', 'Number Nested Anchor', 'Number Root B', 'Number Continued Item'].includes(text)
    ? 'number-continuation'
    : text.startsWith('Number Restart ')
      ? 'number-restart'
      : null;
function numberingRow(text, listLevel = observation(text).listLevel, listString = observation(text).listString) {
  return { text, isList: true, listType: observation(text).listType, listLevel,
    ...(listString !== null && listString !== undefined ? { listString } : {}), group: groupFor(text) };
}
const bulletRows = sourceTexts.filter(text => text.startsWith('Bullet '));
const romanRows = ['Number Root A', 'Number Root B', 'Number Continued Item'];
const insertParagraph = (name, paragraphText, afterParagraphIndex, text, indentLevel,
  { sourcePreparation = [], usesPreparedSource = false, redlineEnabled = true, priorTrackingMode = 'off',
    expectedTrackingModeAfter = priorTrackingMode, expectedListLevel,
    romanStyle = null, directWhenTrackingOff = false } = {}) => {
  const expectedAcceptedTexts = insertAfter(paragraphText, text).split('\n');
  const expectedRejectedText = directWhenTrackingOff ? expectedAcceptedTexts.join('\n') : sourceText;

  let acceptedRows;
  let rejectedRows;
  let sourceRows;
  if (romanStyle) {
    const numbering = romanStyle === 'upperRoman' ? ['I', 'II', 'III', 'IV'] : ['i', 'ii', 'iii', 'iv'];
    const rootText = 'Number Root A';
    const rootIndex = romanRows.indexOf(rootText);
    const sourceRomanNumbers = romanStyle === 'upperRoman' ? ['I', 'II', 'III'] : ['i', 'ii', 'iii'];
    const sourceNumberRows = romanRows.map((item, index) => numberingRow(item,
      observation(item).listLevel,
      sourceRomanNumbers[index]));
    const updatedNumberRows = romanRows.map(item => numberingRow(item,
      observation(item).listLevel,
      item === 'Number Root A' ? numbering[0]
        : item === 'Number Root B' ? numbering[2]
          : item === 'Number Continued Item' ? numbering[3]
            : observation(item).listString));
    const insertedRow = { text, isList: true, listType: 4, listLevel: 1, listString: numbering[1], group: 'number-continuation' };
    sourceRows = sourceNumberRows;
    acceptedRows = [...updatedNumberRows.slice(0, rootIndex + 1), insertedRow, ...updatedNumberRows.slice(rootIndex + 1)];
    rejectedRows = sourceNumberRows;
  } else {
    const listNames = bulletRows;
    const sourceLevels = sourcePreparation.some(step => step.type === 'list-level')
      ? Object.fromEntries(listNames.map(item => [item, item === 'Bullet Insertion Anchor' ? 3 : observation(item).listLevel]))
      : Object.fromEntries(listNames.map(item => [item, observation(item).listLevel]));
    const anchorListRows = listNames.map(item => numberingRow(item, sourceLevels[item],
      item === 'Bullet Insertion Anchor' && sourceLevels[item] === 3 ? null : observation(item).listString));
    sourceRows = anchorListRows;
    const newLevel = expectedListLevel + 1;
    const insertedRow = { text, isList: true, listType: observation(paragraphText).listType,
      listLevel: newLevel, group: 'bullet-chain' };
    const anchorOffset = listNames.indexOf(paragraphText);
    acceptedRows = [...anchorListRows.slice(0, anchorOffset + 1), insertedRow, ...anchorListRows.slice(anchorOffset + 1)];
    rejectedRows = anchorListRows;
  }

  if (directWhenTrackingOff) {
    sourceRows = [numberingRow(paragraphText)];
    acceptedRows = [numberingRow(paragraphText), {
      text, isList: true, listType: observation(paragraphText).listType, listLevel: expectedListLevel + 1,
      listString: observation(paragraphText).listString, group: groupFor(paragraphText)
    }];
    rejectedRows = acceptedRows;
  }

  return {
    name,
    source: 'nested-lists-source.docx',
    ...(usesPreparedSource ? { usesPreparedSource: true } : {}),
    ...(sourcePreparation.length ? { sourcePreparation } : {}),
    agenticRequest: { tool: 'insert_list_item', afterParagraphIndex, text, indentLevel },
    productionInsert: {
      redlineEnabled,
      priorTrackingMode,
      expectedTrackingModeAfter,
      expectedRoute: 'native',
      expectedParagraphInsertCalls: 1,
      expectedBodyInsertCalls: 0,
      expectedListLevel
    },
    expectedSourceText: sourceText,
    expectedAcceptedText: expectedAcceptedTexts.join('\n'),
    expectedRejectedText,
    expectedComments: 0,
    ...(redlineEnabled ? { expectedMinimumBodyRevisions: 1 } : { expectedExactBodyRevisions: 0 }),
    expectedFormatting: [{ sourceText: 'Untouched bold sentinel', acceptedText: 'Untouched bold sentinel', bold: true, italic: false }],
    expectedNumbering: { source: sourceRows, accepted: acceptedRows, rejected: rejectedRows },
    expectedEngineReferencePackages: false
  };
};

const cases = [
  insertParagraph('native-insert-deeper-resolved-level', 'Bullet Insertion Anchor', 4,
    'Native insertion at requested deep bullet level', 1,
    { expectedListLevel: 2 }),
  insertParagraph('native-insert-outdent-from-deep-source', 'Bullet Insertion Anchor', 4,
    'Native insertion outdented from deep bullet source', -1,
    { usesPreparedSource: true, expectedListLevel: 1,
      sourcePreparation: [{ type: 'list-level', paragraphText: 'Bullet Insertion Anchor', level: 2 }] }),
  insertParagraph('native-insert-upper-roman-numbering', 'Number Root A', 8,
    'Native insertion into upper Roman list', 0,
    { usesPreparedSource: true, expectedListLevel: 0, romanStyle: 'upperRoman',
      sourcePreparation: [{ type: 'list-numbering', paragraphText: 'Number Root A', level: 0, numbering: 'upperRoman' }] }),
  insertParagraph('native-insert-lower-roman-numbering', 'Number Root A', 8,
    'Native insertion into lower Roman list', 0,
    { usesPreparedSource: true, expectedListLevel: 0, romanStyle: 'lowerRoman',
      sourcePreparation: [{ type: 'list-numbering', paragraphText: 'Number Root A', level: 0, numbering: 'lowerRoman' }] }),
  insertParagraph('native-insert-tracking-disabled-restores-mode', 'Bullet Root A', 3,
    'Native insertion while redlining is disabled', 0,
    { redlineEnabled: false, priorTrackingMode: 'trackAll', expectedTrackingModeAfter: 'trackAll',
      expectedListLevel: 0, directWhenTrackingOff: true })
];

for (const testCase of cases) {
  assert.equal(testCase.productionInsert.expectedRoute, 'native');
  assert.equal(testCase.agenticRequest.tool, 'insert_list_item');
  assert.ok(testCase.expectedSourceText && testCase.expectedAcceptedText && testCase.expectedRejectedText);
  assert.equal(testCase.expectedComments, 0);
}
assert.ok(cases.some(testCase => testCase.usesPreparedSource));
assert.ok(cases.find(testCase => testCase.name.includes('tracking-disabled')).expectedExactBodyRevisions === 0);

const outputArg = process.argv.indexOf('--export-dir');
const outputDirectory = outputArg >= 0 ? path.resolve(process.argv[outputArg + 1])
  : path.join(root, '.cache/reliability/agentic-list-native-fallback');
assert.ok(outputDirectory && !String(process.argv[outputArg + 1] || '').startsWith('--'), '--export-dir requires a directory');
mkdirSync(outputDirectory, { recursive: true });
const sourcePath = path.join(fixtureDirectory, 'nested-lists-source.docx');
const manifest = {
  schemaVersion: 1,
  provenance: {
    source: 'Word-authored nested list fixture; deep levels and Roman numbering are prepared with Office.js before the production tool call and frozen to separate prepared-source packages.',
    sourceObservations: 'tests/fixtures/agentic-lists/source-observations.json',
    nativeRouteAssertion: 'Production insert_list_item must call Paragraph.insertParagraph exactly once and body.insertOoxml zero times for every case.',
    resolutionOracle: 'Independent Word verification checks source, tracked, Accept All and Reject All states against these expected text and list-level observations.'
  },
  cases: cases.map(testCase => ({
    ...testCase,
    source: path.relative(outputDirectory, sourcePath).replaceAll('\\', '/')
  }))
};
const manifestPath = path.join(outputDirectory, 'manifest.json');
writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`PASS: exported ${cases.length} native insert fallback Word cases to ${manifestPath}`);
