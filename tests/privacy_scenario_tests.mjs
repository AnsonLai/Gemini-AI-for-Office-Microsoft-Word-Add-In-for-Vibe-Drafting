import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { openDocx } from '@ansonlai/docx-redline-js';
import { buildDocumentModel, scoreStep } from '../scripts/golden/golden-checks.mjs';
import { scenario } from '../scripts/golden/privacy-direct-scenario.mjs';

// The second golden document (a Word-authored privacy notice) is well-formed for the runner:
// every replay target exists in the fixture, and no step's checks pass before its edit.
const bytes = new Uint8Array(readFileSync(new URL('./fixtures/golden/privacy-notice.docx', import.meta.url)));
const paragraphs = openDocx(bytes).inspect().paragraphs;
const matches = (matcher, text) => typeof matcher === 'string' ? text.includes(matcher)
  : matcher.equals != null ? text === matcher.equals
    : matcher.startsWith != null ? text.startsWith(matcher.startsWith)
      : matcher.includes != null ? text.includes(matcher.includes)
        : new RegExp(matcher.regex, matcher.flags || '').test(text);

assert.equal(scenario.fixture, 'tests/fixtures/golden/privacy-notice.docx');
assert.equal(new Set(scenario.steps.map(step => step.id)).size, scenario.steps.length, 'step ids are unique');
assert.ok(paragraphs.some(paragraph => paragraph.list?.format === 'bullet') && paragraphs.some(paragraph => paragraph.list?.format === 'decimal'),
  'the fixture has bullet and numbered lists');
assert.ok(paragraphs.some(paragraph => paragraph.inTable), 'the fixture has a table');

for (const step of scenario.steps) {
  assert.ok(step.prompt && step.replay?.length, `${step.id} has a prompt and a replay plan`);
  const edits = step.replay.flatMap(entry => entry.chat?.functionCall?.name === 'apply_document_edits' ? entry.chat.functionCall.args.edits : []);
  for (const edit of edits) {
    if (edit.target?.$handle) {
      assert.ok(paragraphs.some(paragraph => matches(edit.target.$handle, paragraph.exactText)),
        `${step.id}: replay target ${JSON.stringify(edit.target.$handle)} exists in the fixture`);
    }
  }
}

const model = await buildDocumentModel(bytes);
for (const [index, step] of scenario.steps.entries()) {
  if (step.scopeDisposition) continue;
  const own = scoreStep(model, scenario.steps, index).filter(result => result.own && result.label !== 'edit batches have one verified commit and no nested model calls');
  assert.ok(own.some(result => !result.ok), `${step.id} is not already satisfied by the source document`);
}

console.log('privacy_scenario_tests passed');
