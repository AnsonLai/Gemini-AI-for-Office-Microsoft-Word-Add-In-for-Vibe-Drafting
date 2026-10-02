import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { inspectDocumentParts } from '@ansonlai/docx-redline-js';
import { buildCanonicalContextText } from '../src/taskpane/modules/chat/canonical-context.js';
import { parseAnchoredParagraphs, sanitizeChangeSet } from '../src/taskpane/modules/commands/change-validation.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const REV = 'w:id="1" w:author="Editor" w:date="2026-10-01T00:00:00Z"';

// The state left behind by a tracked text-to-table edit: the source paragraph
// is a pending deletion and the new table text is a pending insertion.
const documentXml = `<w:document ${W}><w:body>`
  + '<w:p><w:r><w:t>RECITALS</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>A. First recital.</w:t></w:r><w:r><w:br/><w:t>B. Second recital.</w:t></w:r></w:p>'
  + `<w:p><w:pPr><w:rPr><w:del ${REV}/></w:rPr></w:pPr><w:del ${REV}><w:r><w:delText>Disclosing Party: Acme</w:delText></w:r></w:del></w:p>`
  + `<w:tbl><w:tr><w:tc><w:p><w:ins ${REV}><w:r><w:t>Acme Corp</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl>`
  + '<w:p><w:r><w:t>Kept </w:t></w:r><w:del ' + REV + '><w:r><w:delText>old </w:delText></w:r></w:del><w:r><w:t>words.</w:t></w:r></w:p>'
  + '</w:body></w:document>';
const baseline = inspectDocumentParts({ documentXml }).paragraphs
  .map(paragraph => ({ index: paragraph.index, exactText: paragraph.exactText }));

// What Word's Paragraph.text reports for the same paragraphs.
const wordParagraphs = [
  { index: 1, meta: 'Heading1', text: 'RECITALS' },
  { index: 2, meta: 'Normal', text: 'A. First recital.\u000bB. Second recital.' },
  { index: 3, meta: 'Normal', text: 'Disclosing Party: Acme' },
  { index: 4, meta: 'Normal|T:0,0', text: 'Acme Corp' },
  { index: 5, meta: 'Normal', text: 'Kept old words.' }
];

const text = buildCanonicalContextText(wordParagraphs, baseline);
assert.equal(text, [
  '[P1|Heading1] RECITALS',
  '[P2|Normal] A. First recital.',
  'B. Second recital.',
  '[P3|Normal|deleted] ',
  '[P4|Normal|T:0,0] Acme Corp',
  '[P5|Normal] Kept words.'
].join('\n'));

// The parsed view the validators use now matches the engine's targets.
const parsed = parseAnchoredParagraphs(text);
assert.deepEqual(parsed, baseline.map(paragraph => paragraph.exactText));

// A find copied from pending-deleted text is rejected locally with a
// correctable reason instead of failing later in the engine.
const stale = sanitizeChangeSet([{ paragraphIndex: 5, operation: 'edit_paragraph', anchorText: 'Kept',
  replacements: [{ find: 'old words', replace: 'new words' }] }], parsed.length, parsed);
assert.equal(stale.rejected[0]?.reason, 'replacement_find_not_found');

// Desktop Word's body OOXML adds one empty trailing paragraph; tolerate only that.
assert.equal(buildCanonicalContextText(wordParagraphs, [...baseline, { index: 6, exactText: '' }]), text);
assert.equal(buildCanonicalContextText(wordParagraphs, [...baseline, { index: 6, exactText: 'x' }]), null, 'non-empty extra is misaligned');
assert.equal(buildCanonicalContextText(wordParagraphs,
  [...baseline, { index: 6, exactText: '' }, { index: 7, exactText: '' }]), null, 'only one trailing artifact is tolerated');

// Misaligned views fall back to the Word text (caller keeps its own context).
assert.equal(buildCanonicalContextText(wordParagraphs.slice(1), baseline), null);
assert.equal(buildCanonicalContextText([{ index: 1, text: 'x' }], [{ index: 1, exactText: 'x' }]), null, 'meta is required');
assert.equal(buildCanonicalContextText([], []), null);

console.log('canonical_context_tests passed');
