import assert from 'node:assert/strict';
import { glanceResponseText, parseGlanceResults } from '../src/taskpane/modules/glance/glance-results.js';

// Search grounding can split the answer across parts; thoughts are excluded.
assert.equal(glanceResponseText({ candidates: [{ content: { parts: [
  { text: 'hidden reasoning', thought: true }, { text: '{"q1": {"status": "gr' }, { text: 'een", "summary": "Fine"}}' },
] } }] }), '{"q1": {"status": "green", "summary": "Fine"}}');
assert.throws(() => glanceResponseText({ candidates: [{ finishReason: 'SAFETY' }] }), /no Glance answer/);
assert.throws(() => glanceResponseText({}), /no Glance answer/);
assert.throws(() => glanceResponseText({ candidates: [{ content: { parts: [{ text: '  ' }] } }] }), /empty Glance answer/);

// Fenced JSON is read; unknown statuses become gray; missing ids are left out.
const results = parseGlanceResults('```json\n{"q1": {"status": "red", "summary": "Typos"}, "q2": {"status": "glance-indicator pulse", "summary": 7}}\n```', ['q1', 'q2', 'q3']);
assert.deepEqual(results, {
  q1: { status: 'red', summary: 'Typos' },
  q2: { status: 'gray', summary: '' },
});
assert.equal(parseGlanceResults(`{"q1": {"status": "green", "summary": "${'x'.repeat(500)}"}}`, ['q1']).q1.summary.length, 200);
assert.throws(() => parseGlanceResults('No issues found.', ['q1']), /did not contain JSON/);

console.log('PASS: Glance answers are joined across parts and only known statuses are shown');
