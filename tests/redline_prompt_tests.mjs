import assert from 'assert';
import {
  buildRedlineDiffPrompt,
  buildCorrectiveRetryPrompt,
  REDLINE_DIFF_SCHEMA
} from '../src/taskpane/modules/commands/redline-prompt.js';

function testPromptContainsSentinels() {
  const prompt = buildRedlineDiffPrompt('MY_INSTRUCTION', 'MY_DOC');

  // Key structural lines that must survive the extraction unchanged.
  const sentinels = [
    'You are an expert legal editor.',
    'CRITICAL: Return ONLY valid JSON.',
    '"anchorText": REQUIRED for every change. Copy the first 30-60 characters',
    'PREFER LOCALIZED EDITS',
    '"replacements": [{"find":"exact current text","replace":"new text"}]',
    // The table example uses LITERAL backslash-n (not real newlines):
    '| Column A | Column B |\\n|---|---|\\n| Value A | Value B |',
    'Do NOT include the [P#] marker in any content fields.',
    'Return ONLY the JSON array, nothing else:'
  ];
  for (const s of sentinels) {
    assert.ok(prompt.includes(s), `prompt missing sentinel: ${JSON.stringify(s.slice(0, 50))}`);
  }
}

function testInterpolationSubstitution() {
  const prompt = buildRedlineDiffPrompt('FIX_THE_TYPOS', '[P1] hello world');
  assert.ok(prompt.includes('USER INSTRUCTION:\n"FIX_THE_TYPOS"'), 'instruction interpolated');
  assert.ok(prompt.includes('DOCUMENT CONTENT:\n"""[P1] hello world"""'), 'document interpolated');
  // The two interpolations are the only variable parts; the rest is stable.
  const a = buildRedlineDiffPrompt('A', 'D');
  const b = buildRedlineDiffPrompt('B', 'D');
  assert.strictEqual(a.length - b.length, 0, 'same-length instructions give same-length prompts');
}

function testSchemaShape() {
  assert.strictEqual(REDLINE_DIFF_SCHEMA.type, 'ARRAY');
  assert.deepStrictEqual(REDLINE_DIFF_SCHEMA.items.required, ['paragraphIndex', 'operation', 'anchorText']);
  const props = REDLINE_DIFF_SCHEMA.items.properties;
  for (const key of ['paragraphIndex', 'anchorText', 'endParagraphIndex', 'operation', 'newContent', 'replacements', 'content', 'originalText', 'replacementText']) {
    assert.ok(props[key], `schema missing property: ${key}`);
  }
  assert.deepStrictEqual(props.operation.enum, ['edit_paragraph', 'replace_paragraph', 'modify_text', 'replace_range', 'format_text']);
  assert.strictEqual(props.find.type, 'STRING');
  assert.deepStrictEqual(Object.keys(props.formatting.properties), ['bold', 'italic', 'underline', 'strikethrough']);
  assert.strictEqual(props.replacements.type, 'ARRAY');
  assert.deepStrictEqual(props.replacements.items.required, ['find', 'replace']);
}

function testCorrectiveRetryPrompt() {
  const base = buildRedlineDiffPrompt('FIX_IT', '[P1] hello');
  const previous = [{ paragraphIndex: 1, operation: 'replace_range', replacementText: 'leaked reasoning' }];
  const detail = '- P1 (replace_range): empty_content.';
  const retry = buildCorrectiveRetryPrompt(base, previous, detail);

  // Contains the full original prompt (same document + instruction)...
  assert.ok(retry.startsWith(base), 'retry prompt must start with the base prompt');
  // ...the model's own invalid output...
  assert.ok(retry.includes('"replacementText": "leaked reasoning"'));
  // ...the machine rejection reasons...
  assert.ok(retry.includes('- P1 (replace_range): empty_content.'));
  // ...and the corrective instructions.
  assert.ok(retry.includes('YOUR PREVIOUS RESPONSE WAS REJECTED'));
  assert.ok(retry.includes('CORRECTED JSON array'));
  assert.ok(retry.includes('NEVER write notes, schema commentary, or reasoning'));
  // Survives unserializable previous output.
  const circular = {}; circular.self = circular;
  assert.ok(buildCorrectiveRetryPrompt(base, [circular], 'x').includes('YOUR PREVIOUS RESPONSE WAS REJECTED'));

  // A degenerate previous response (many repeated objects) is truncated so the
  // retry prompt stays a sane size.
  const spam = [];
  for (let i = 0; i < 200; i++) {
    spam.push({ paragraphIndex: 1, operation: 'replace_range', anchorText: 'Green grass, softest breeze,', replacementText: '' });
  }
  const retrySpam = buildCorrectiveRetryPrompt(base, spam, detail);
  assert.ok(retrySpam.length < base.length + 6000, 'previous JSON echo must be capped');
  assert.ok(retrySpam.includes('truncated; the response continued repeating'));
}

testPromptContainsSentinels();
testInterpolationSubstitution();
testSchemaShape();
testCorrectiveRetryPrompt();

console.log('redline_prompt_tests passed');
