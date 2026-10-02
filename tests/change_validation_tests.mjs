import assert from 'assert';
import {
  parseAnchoredParagraphs,
  verifyAnchor,
  normalizeForAnchor,
  sanitizeChangeSet,
  formatRejections,
  repairTruncatedJsonArray
} from '../src/taskpane/modules/commands/change-validation.js';

// Sample anchored document text in the [P#|meta] text format produced by
// extractEnhancedDocumentContext().
const DOC = [
  '[P1|Heading1] MUTUAL NON-DISCLOSURE AGREEMENT',
  '[P2|Normal] This Agreement is entered into by the parties below.',
  '[P3|Normal|§1] The Receiving Party shall protect Confidential Information.',
  '[P4|Normal|§1] ARTICLE 4 — TERM AND TERMINATION',
  '[P5|Normal] '
].join('\n');

function testParseAnchoredParagraphs() {
  const texts = parseAnchoredParagraphs(DOC);
  assert.strictEqual(texts.length, 5);
  assert.strictEqual(texts[0], 'MUTUAL NON-DISCLOSURE AGREEMENT');
  assert.strictEqual(texts[2], 'The Receiving Party shall protect Confidential Information.');
  assert.strictEqual(texts[4], '');
  // Empty / non-string input is safe.
  assert.deepStrictEqual(parseAnchoredParagraphs(''), []);
  assert.deepStrictEqual(parseAnchoredParagraphs(null), []);
}

function testExactAnchorMatch() {
  const texts = parseAnchoredParagraphs(DOC);
  const change = { paragraphIndex: 3, operation: 'edit_paragraph', anchorText: 'The Receiving Party shall protect' };
  assert.deepStrictEqual(verifyAnchor(change, texts), { ok: true });
}

function testOffByOneCorrection() {
  const texts = parseAnchoredParagraphs(DOC);
  // Model says P4 but the anchor really belongs to P3 (index+(-1)).
  const change = { paragraphIndex: 4, operation: 'edit_paragraph', anchorText: 'The Receiving Party shall protect' };
  const verdict = verifyAnchor(change, texts);
  assert.strictEqual(verdict.ok, true);
  assert.strictEqual(verdict.correctedIndex, 3);
}

function testNoMatchInWindowReportsActual() {
  const texts = parseAnchoredParagraphs(DOC);
  // Anchor matches nowhere near the claimed index.
  const change = { paragraphIndex: 1, operation: 'edit_paragraph', anchorText: 'Some text that does not exist anywhere' };
  const verdict = verifyAnchor(change, texts);
  assert.strictEqual(verdict.ok, false);
  assert.strictEqual(verdict.reason, 'anchor_mismatch');
  // actualTextSnippet reflects what is really at the claimed index (P1).
  assert.strictEqual(verdict.actualTextSnippet, 'MUTUAL NON-DISCLOSURE AGREEMENT');
}

function testNormalizationTolerance() {
  const texts = parseAnchoredParagraphs(DOC);
  // Extra whitespace and a stray [P#] prefix in the anchor should still match.
  const change = {
    paragraphIndex: 3,
    operation: 'edit_paragraph',
    anchorText: '[P3] The   Receiving    Party  shall'
  };
  assert.deepStrictEqual(verifyAnchor(change, texts), { ok: true });

  // Direct normalizer checks.
  assert.strictEqual(normalizeForAnchor('[P3|Normal|§1]  hello   world '), 'hello world');
  assert.strictEqual(normalizeForAnchor(null), '');
}

function testMissingAnchorIsCompatMode() {
  const texts = parseAnchoredParagraphs(DOC);
  assert.deepStrictEqual(verifyAnchor({ paragraphIndex: 2, operation: 'edit_paragraph' }, texts), { ok: true });
  assert.deepStrictEqual(verifyAnchor({ paragraphIndex: 2, anchorText: '' }, texts), { ok: true });
  assert.deepStrictEqual(verifyAnchor({ paragraphIndex: 2, anchorText: '   ' }, texts), { ok: true });
}

function testAmbiguousNeighborsRejected() {
  // Two adjacent paragraphs share the same prefix -> ambiguous -> reject.
  const doc = [
    '[P1|Normal] Payment terms are net 30 days.',
    '[P2|Normal] Payment terms are net 30 days.',
    '[P3|Normal] Payment terms are net 30 days.'
  ].join('\n');
  const texts = parseAnchoredParagraphs(doc);
  const change = { paragraphIndex: 1, operation: 'edit_paragraph', anchorText: 'Payment terms are net 30' };
  // Claimed index P1 matches exactly, so this is ok (exact match wins first).
  assert.strictEqual(verifyAnchor(change, texts).ok, true);

  // But a wrong claimed index with multiple matching neighbors is ambiguous.
  const change2 = { paragraphIndex: 5, operation: 'edit_paragraph', anchorText: 'Payment terms are net 30' };
  const verdict = verifyAnchor(change2, texts);
  // P5 is out of range and neighbors P3/P4 -> only P3 matches in window? window is +-2 of 5 => 3,4,6,7. Only P3 matches.
  assert.strictEqual(verdict.ok, true);
  assert.strictEqual(verdict.correctedIndex, 3);
}

// --------------------------------------------------------------------------
// WP2: sanitizeChangeSet
// --------------------------------------------------------------------------

const PARAGRAPH_COUNT = 10;

// Helper: returns the single rejection reason for a one-change input.
function rejectReason(change, count = PARAGRAPH_COUNT) {
  const { changes, rejected } = sanitizeChangeSet([change], count);
  assert.strictEqual(changes.length, 0, `expected change to be rejected: ${JSON.stringify(change)}`);
  assert.strictEqual(rejected.length, 1);
  return rejected[0].reason;
}

function testCleanChangeSetPassesThrough() {
  const input = [
    { paragraphIndex: 2, operation: 'edit_paragraph', anchorText: 'a', newContent: 'New text.' },
    { paragraphIndex: 4, operation: 'replace_range', endParagraphIndex: 6, anchorText: 'b', content: '| A | B |\n|---|---|\n| 1 | 2 |' },
    { paragraphIndex: 7, operation: 'modify_text', anchorText: 'c', originalText: 'foo', replacementText: 'bar' }
  ];
  const { changes, rejected } = sanitizeChangeSet(input, PARAGRAPH_COUNT);
  assert.strictEqual(rejected.length, 0);
  assert.strictEqual(changes.length, 3);
  // Valid changes are not altered (no [P#] markers present here).
  assert.deepStrictEqual(changes[0], input[0]);
  assert.deepStrictEqual(changes[2], input[2]);
}

function testInvalidOperation() {
  assert.strictEqual(rejectReason(null), 'invalid_operation');
  assert.strictEqual(rejectReason('not an object'), 'invalid_operation');
  assert.strictEqual(rejectReason({ paragraphIndex: 1, operation: 'delete_everything' }), 'invalid_operation');
  assert.strictEqual(rejectReason({ paragraphIndex: 1 }), 'invalid_operation');
}

function testIndexOutOfRange() {
  assert.strictEqual(rejectReason({ paragraphIndex: 0, operation: 'edit_paragraph', newContent: 'x' }), 'index_out_of_range');
  assert.strictEqual(rejectReason({ paragraphIndex: 99, operation: 'edit_paragraph', newContent: 'x' }), 'index_out_of_range');
  assert.strictEqual(rejectReason({ paragraphIndex: 'two', operation: 'edit_paragraph', newContent: 'x' }), 'index_out_of_range');
  // replace_range with bad end.
  assert.strictEqual(rejectReason({ paragraphIndex: 5, operation: 'replace_range', endParagraphIndex: 3, content: 'x' }), 'index_out_of_range');
  // Unknown paragraph count disables the upper bound but not the lower bound.
  assert.strictEqual(rejectReason({ paragraphIndex: 0, operation: 'edit_paragraph', newContent: 'x' }, 0), 'index_out_of_range');
  const { changes } = sanitizeChangeSet([{ paragraphIndex: 9999, operation: 'edit_paragraph', newContent: 'x' }], 0);
  assert.strictEqual(changes.length, 1, 'unknown count should not reject a high index');
}

function testAppendAtEndAllowed() {
  // paragraphCount = 5; index 6 (count+1) is allowed for content-bearing ops (append).
  for (const op of ['replace_paragraph', 'replace_range']) {
    const change = { paragraphIndex: 6, operation: op, anchorText: '', content: 'New tail content', endParagraphIndex: 6 };
    const { changes, rejected } = sanitizeChangeSet([change], 5);
    assert.strictEqual(changes.length, 1, `${op} append at count+1 should be allowed`);
    assert.strictEqual(rejected.length, 0);
  }
  // edit_paragraph at count+1 is also allowed (treated as append by the bridge).
  assert.strictEqual(
    sanitizeChangeSet([{ paragraphIndex: 6, operation: 'edit_paragraph', newContent: 'x' }], 5).changes.length,
    1
  );
  // But count+2 is still rejected (forces consolidation into one append change).
  assert.strictEqual(rejectReason({ paragraphIndex: 7, operation: 'replace_paragraph', content: 'x' }, 5), 'index_out_of_range');
  // And modify_text at count+1 is rejected (nothing to edit there).
  assert.strictEqual(rejectReason({ paragraphIndex: 6, operation: 'modify_text', originalText: 'a', replacementText: 'b' }, 5), 'index_out_of_range');
}

function testStripsParagraphMarkers() {
  const input = [{ paragraphIndex: 2, operation: 'edit_paragraph', newContent: '[P2] Hello [P3] world' }];
  const { changes } = sanitizeChangeSet(input, PARAGRAPH_COUNT);
  assert.strictEqual(changes.length, 1);
  assert.strictEqual(changes[0].newContent, 'Hello world');
}

function testEmptyContent() {
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'replace_paragraph', content: '' }), 'empty_content');
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'replace_paragraph' }), 'empty_content');
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'replace_range', endParagraphIndex: 3, content: '   ' }), 'empty_content');
  // edit_paragraph without newContent is also empty_content.
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'edit_paragraph' }), 'empty_content');
}

function testThoughtLeakageIntoUnusedFieldRejected() {
  // Regression: exact shape observed from gemini-3.5-flash — reasoning leaked
  // into replacementText, no content provided. Must be rejected (empty_content),
  // and the junk must NOT survive into an applied change (the engine falls back
  // to replacementText as content, which would write the junk into the document).
  const leaked = {
    paragraphIndex: 1,
    operation: 'replace_range',
    anchorText: 'Green grass, softest breeze,',
    endParagraphIndex: 3,
    originalText: '',
    replacementText: "This field is not used for replace_range, but schema requires it... Wait, the schema says 'content' is required..."
  };
  const { changes, rejected } = sanitizeChangeSet([leaked], 3);
  assert.strictEqual(changes.length, 0);
  assert.strictEqual(rejected[0].reason, 'empty_content');
}

function testWrongFieldContentRepair() {
  // replace_* with the payload in newContent instead of content -> repaired.
  const r1 = sanitizeChangeSet(
    [{ paragraphIndex: 2, operation: 'replace_paragraph', anchorText: 'a', newContent: 'Actual replacement text' }],
    5
  );
  assert.strictEqual(r1.changes.length, 1);
  assert.strictEqual(r1.changes[0].content, 'Actual replacement text');
  assert.strictEqual(r1.changes[0].newContent, undefined, 'inapplicable field stripped');

  // edit_paragraph with the payload in content instead of newContent -> repaired.
  const r2 = sanitizeChangeSet(
    [{ paragraphIndex: 2, operation: 'edit_paragraph', anchorText: 'a', content: 'Rewritten paragraph' }],
    5
  );
  assert.strictEqual(r2.changes.length, 1);
  assert.strictEqual(r2.changes[0].newContent, 'Rewritten paragraph');
  assert.strictEqual(r2.changes[0].content, undefined);
}

function testInapplicableFieldsStripped() {
  // Junk in unused fields is deleted so the engine's content fallback can never
  // reach it; the valid content is preserved.
  const { changes } = sanitizeChangeSet(
    [{
      paragraphIndex: 2,
      operation: 'replace_range',
      endParagraphIndex: 3,
      anchorText: 'a',
      content: '**Header**\n* bullet one\n* bullet two',
      originalText: 'stray junk',
      replacementText: 'model reasoning junk'
    }],
    5
  );
  assert.strictEqual(changes.length, 1);
  assert.strictEqual(changes[0].content, '**Header**\n* bullet one\n* bullet two');
  assert.strictEqual(changes[0].originalText, undefined);
  assert.strictEqual(changes[0].replacementText, undefined);

  // modify_text keeps its own fields but sheds content/newContent.
  const r2 = sanitizeChangeSet(
    [{ paragraphIndex: 2, operation: 'modify_text', anchorText: 'a', originalText: 'foo', replacementText: 'bar', content: 'junk', newContent: 'junk' }],
    5
  );
  assert.strictEqual(r2.changes.length, 1);
  assert.strictEqual(r2.changes[0].originalText, 'foo');
  assert.strictEqual(r2.changes[0].replacementText, 'bar');
  assert.strictEqual(r2.changes[0].content, undefined);
  assert.strictEqual(r2.changes[0].newContent, undefined);
}

function testOriginalTextTooLong() {
  const longText = 'x'.repeat(81);
  assert.strictEqual(
    rejectReason({ paragraphIndex: 2, operation: 'modify_text', originalText: longText, replacementText: 'ok' }),
    'original_text_too_long'
  );
}

function testModifyTextStructuralContent() {
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'modify_text', originalText: 'a', replacementText: 'line1\nline2' }), 'modify_text_structural_content');
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'modify_text', originalText: 'a', replacementText: '| A | B |' }), 'modify_text_structural_content');
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'modify_text', originalText: 'a', replacementText: '- bullet' }), 'modify_text_structural_content');
}

function testFormatText() {
  const texts = ['The Confidential Information and Confidential terms.', 'Second paragraph.'];
  const valid = { paragraphIndex: 1, operation: 'format_text', anchorText: 'The Confidential',
    find: 'Confidential Information', formatting: { bold: false } };
  const accepted = sanitizeChangeSet([valid, { ...valid, find: 'terms', formatting: { italic: true } }], texts.length, texts);
  assert.strictEqual(accepted.rejected.length, 0, 'distinct spans in one paragraph are not duplicates');
  assert.deepStrictEqual(verifyAnchor(valid, texts), { ok: true });

  const reason = change => sanitizeChangeSet([change], texts.length, texts).rejected[0]?.reason;
  assert.strictEqual(reason({ ...valid, find: 'confidential information' }), 'replacement_find_not_found');
  assert.strictEqual(verifyAnchor({ ...valid, find: 'missing' }, texts).reason, 'replacement_find_not_found');
  assert.strictEqual(reason({ ...valid, find: 'Confidential' }), 'format_find_ambiguous');
  assert.strictEqual(sanitizeChangeSet([{ ...valid, find: 'Confidential', occurrence: 2 }], texts.length, texts).changes.length, 1,
    'occurrence selects one appearance of a repeated find');
  assert.strictEqual(reason({ ...valid, find: 'Confidential', occurrence: 3 }), 'replacement_find_not_found');
  assert.strictEqual(reason({ ...valid, find: 'Confidential', occurrence: 0 }), 'invalid_format_text');
  assert.strictEqual(reason({ ...valid, formatting: {} }), 'invalid_format_text');
  assert.strictEqual(reason({ ...valid, formatting: { bold: 'no' } }), 'invalid_format_text');
  assert.strictEqual(reason({ ...valid, formatting: { color: false } }), 'invalid_format_text');
  assert.strictEqual(reason({ ...valid, find: 'Confidential\nInformation' }), 'invalid_format_text');
  assert.strictEqual(reason({ ...valid, find: undefined }), 'invalid_format_text');
  assert.match(formatRejections([{ paragraphIndex: 1, operation: 'format_text', reason: 'format_find_ambiguous' }]),
    /1-based "occurrence"/);
}

function testLocalizedReplacements() {
  const texts = ['The Seller shall notify the Buyer promptly.', 'A repeated term and term again.'];
  const valid = { paragraphIndex: 1, operation: 'edit_paragraph', anchorText: 'The Seller shall',
    replacements: [{ find: 'Seller', replace: 'Provider' }, { find: 'Buyer', replace: 'Customer' }] };
  const accepted = sanitizeChangeSet([valid], texts.length, texts);
  assert.strictEqual(accepted.rejected.length, 0);
  assert.deepStrictEqual(accepted.changes[0].replacements, valid.replacements);
  assert.deepStrictEqual(verifyAnchor(valid, texts), { ok: true });

  const missing = { ...valid, replacements: [{ find: 'seller', replace: 'Provider' }] };
  assert.strictEqual(sanitizeChangeSet([missing], texts.length, texts).rejected[0].reason, 'replacement_find_not_found');
  assert.strictEqual(verifyAnchor(missing, texts).reason, 'replacement_find_not_found');
  assert.strictEqual(rejectReason({ ...valid, replacements: [] }, 2), 'invalid_replacements');
  assert.strictEqual(rejectReason({ ...valid, newContent: 'Full rewrite' }, 2), 'invalid_replacements');
  assert.strictEqual(rejectReason({ ...valid, replacements: [{ find: '', replace: 'x' }] }, 2), 'invalid_replacements');

  const ambiguous = { paragraphIndex: 2, operation: 'edit_paragraph', anchorText: 'A repeated',
    replacements: [{ find: 'term', replace: 'phrase' }] };
  assert.strictEqual(sanitizeChangeSet([ambiguous], texts.length, texts).rejected[0].reason, 'replacement_ambiguous_find');
  assert.strictEqual(verifyAnchor(ambiguous, texts).reason, 'replacement_ambiguous_find');
  const selected = { ...ambiguous, replacements: [{ find: 'term', replace: 'phrase', occurrence: 2 }] };
  assert.strictEqual(sanitizeChangeSet([selected], texts.length, texts).changes.length, 1);
  assert.deepStrictEqual(verifyAnchor(selected, texts), { ok: true });
  const missingOccurrence = { ...ambiguous, replacements: [{ find: 'term', replace: 'phrase', occurrence: 3 }] };
  assert.strictEqual(sanitizeChangeSet([missingOccurrence], texts.length, texts).rejected[0].reason, 'replacement_find_not_found');

  // Anchors fold quote styles (models copy "Party’s" as "Party's"); finds stay exact.
  const curly = ['The Receiving Party’s legal counsel may retain one (1) copy.'];
  assert.deepStrictEqual(verifyAnchor({ paragraphIndex: 1, operation: 'edit_paragraph', anchorText: "The Receiving Party's legal counsel",
    replacements: [{ find: 'one (1) copy', replace: 'two (2) copies' }] }, curly), { ok: true });
  assert.strictEqual(verifyAnchor({ paragraphIndex: 1, operation: 'edit_paragraph', anchorText: 'The Receiving Party',
    replacements: [{ find: "Party's legal", replace: 'x' }] }, curly).reason, 'replacement_find_not_found');

  const corrected = { paragraphIndex: 2, operation: 'edit_paragraph', anchorText: 'The Seller shall',
    replacements: [{ find: 'Seller', replace: 'Provider' }] };
  assert.deepStrictEqual(verifyAnchor(corrected, texts), { ok: true, correctedIndex: 1 });
}

function testMalformedTable() {
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'replace_paragraph', content: 'A|B|C' }), 'malformed_table');
  // A real multiline markdown table is fine.
  const { changes } = sanitizeChangeSet([{ paragraphIndex: 2, operation: 'replace_paragraph', content: '| A | B |\n|---|---|\n| 1 | 2 |' }], PARAGRAPH_COUNT);
  assert.strictEqual(changes.length, 1);
}

function testSchemaTextLeak() {
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'edit_paragraph', newContent: 'Set paragraphIndex to 3' }), 'schema_text_in_content');
  assert.strictEqual(rejectReason({ paragraphIndex: 2, operation: 'replace_paragraph', content: '{"operation": "edit_paragraph"}' }), 'schema_text_in_content');
}

function testDeduplicate() {
  const input = [
    { paragraphIndex: 2, operation: 'edit_paragraph', newContent: 'first' },
    { paragraphIndex: 2, operation: 'edit_paragraph', newContent: 'second' },
    { paragraphIndex: 2, operation: 'modify_text', originalText: 'a', replacementText: 'b' }
  ];
  const { changes, rejected } = sanitizeChangeSet(input, PARAGRAPH_COUNT);
  assert.strictEqual(changes.length, 2, 'keeps first edit_paragraph + the different modify_text');
  assert.strictEqual(changes[0].newContent, 'first');
  assert.strictEqual(rejected.length, 1);
  assert.strictEqual(rejected[0].reason, 'duplicate_target');
}

function testFormatRejections() {
  assert.strictEqual(formatRejections([]), '');
  assert.strictEqual(formatRejections(null), '');
  const text = formatRejections([
    { paragraphIndex: 3, operation: 'modify_text', reason: 'original_text_too_long' },
    { paragraphIndex: 5, operation: 'edit_paragraph', reason: 'anchor_mismatch', anchorText: 'Foo', actualTextSnippet: 'Bar baz' }
  ]);
  const lines = text.split('\n');
  assert.strictEqual(lines.length, 2, 'one line per rejection');
  // Each line carries index, operation, and reason.
  assert.ok(lines[0].includes('P3') && lines[0].includes('modify_text') && lines[0].includes('original_text_too_long'));
  assert.ok(lines[1].includes('P5') && lines[1].includes('edit_paragraph') && lines[1].includes('anchor_mismatch'));
  // Anchor mismatch includes the claimed anchor and the actual text snippet.
  assert.ok(lines[1].includes('Foo') && lines[1].includes('Bar baz'));
}

function testRepairTruncatedJsonArray() {
  // Valid JSON parses without repair.
  const good = '[{"paragraphIndex":1,"operation":"edit_paragraph","newContent":"x"}]';
  const r1 = repairTruncatedJsonArray(good);
  assert.strictEqual(r1.repaired, false);
  assert.strictEqual(r1.changes.length, 1);

  // Simulates the observed gemini-3.5-flash repetition loop: complete objects
  // followed by one cut off mid-string when maxOutputTokens ran out.
  const obj = '{"paragraphIndex":1,"operation":"replace_range","anchorText":"Green grass, softest breeze,","endParagraphIndex":3,"replacementText":""}';
  const truncated = `[\n${obj},\n${obj},\n${obj},\n{"paragraphIndex":1,"operation":"replace_range","anchorText":"Green grass, s`;
  const r2 = repairTruncatedJsonArray(truncated);
  assert.strictEqual(r2.repaired, true);
  assert.strictEqual(r2.changes.length, 3, 'salvages the complete objects, drops the cut-off one');
  assert.strictEqual(r2.changes[0].operation, 'replace_range');

  // Braces and escaped quotes inside string values do not confuse the scanner.
  const tricky = '[{"paragraphIndex":2,"operation":"edit_paragraph","newContent":"a {brace} and \\"quote\\" and [bracket]"},{"paragraphIndex":3,"operation":"edit_paragraph","newContent":"cut off he';
  const r3 = repairTruncatedJsonArray(tricky);
  assert.strictEqual(r3.repaired, true);
  assert.strictEqual(r3.changes.length, 1);
  assert.ok(r3.changes[0].newContent.includes('{brace}'));

  // Unsalvageable inputs return null.
  assert.strictEqual(repairTruncatedJsonArray('not json at all').changes, null);
  assert.strictEqual(repairTruncatedJsonArray('').changes, null);
  assert.strictEqual(repairTruncatedJsonArray(null).changes, null);
  assert.strictEqual(repairTruncatedJsonArray('[{"a": "never closes').changes, null);
  // A top-level object (not array) is not a change set.
  assert.strictEqual(repairTruncatedJsonArray('{"paragraphIndex":1}').changes, null);
}

function testFormatRejectionsCap() {
  // A degenerate change set (same rejection dozens of times) must not produce a
  // giant feedback message: cap at 8 lines + a summary line.
  const rejected = [];
  for (let i = 0; i < 12; i++) {
    rejected.push({ paragraphIndex: 1, operation: 'replace_range', reason: 'empty_content' });
  }
  const text = formatRejections(rejected);
  const lines = text.split('\n');
  assert.strictEqual(lines.length, 9, '8 detail lines + 1 summary line');
  assert.ok(lines[8].includes('and 4 more'));
}

function testFormatRejectionsEngineSkip() {
  // Engine-skip entries (from the apply bridge) carry a free-text reason not in
  // the hint map; it must still render verbatim so the model learns the cause.
  const text = formatRejections([
    { paragraphIndex: 5, operation: 'replace_paragraph', reason: 'paragraph P5 is empty; replace_paragraph cannot target an empty paragraph' }
  ]);
  assert.ok(text.includes('P5'));
  assert.ok(text.includes('replace_paragraph'));
  assert.ok(text.includes('empty paragraph'));
}

testParseAnchoredParagraphs();
testExactAnchorMatch();
testOffByOneCorrection();
testNoMatchInWindowReportsActual();
testNormalizationTolerance();
testMissingAnchorIsCompatMode();
testAmbiguousNeighborsRejected();

testCleanChangeSetPassesThrough();
testInvalidOperation();
testIndexOutOfRange();
testAppendAtEndAllowed();
testStripsParagraphMarkers();
testEmptyContent();
testThoughtLeakageIntoUnusedFieldRejected();
testWrongFieldContentRepair();
testInapplicableFieldsStripped();
testOriginalTextTooLong();
testModifyTextStructuralContent();
testLocalizedReplacements();
testFormatText();
testMalformedTable();
testSchemaTextLeak();
testDeduplicate();
testFormatRejections();
testRepairTruncatedJsonArray();
testFormatRejectionsCap();
testFormatRejectionsEngineSkip();

console.log('change_validation_tests passed');
