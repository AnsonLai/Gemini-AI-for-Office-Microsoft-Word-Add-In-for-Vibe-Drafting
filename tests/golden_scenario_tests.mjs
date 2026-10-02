// Offline sanity checks for the golden scenario harness (no Word, no model).
// The scenario itself runs in desktop Word: see scripts/README.md "Golden scenario".
import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { openDocx } from '@ansonlai/docx-redline-js';
import { scenario } from '../scripts/golden/nda-scenario.mjs';
import { buildDocumentModel, evaluateCheck, scoreStep } from '../scripts/golden/golden-checks.mjs';
import { classifyRequest, createReplayQueue, extractContext, resolveReferences } from '../scripts/golden/golden-replay.mjs';
import { buildRedlineDiffPrompt, REDLINE_DIFF_SCHEMA } from '../src/taskpane/modules/commands/redline-prompt.js';

const bytes = new Uint8Array(readFileSync(new URL(`../${scenario.fixture}`, import.meta.url)));
const inspected = openDocx(bytes).inspect().paragraphs;
// Same shape as the taskpane's canonical [P#|meta] view.
const contextText = inspected.map(paragraph => `[P${paragraph.index}|Normal${paragraph.inTable ? '|T:0,0' : ''}] ${paragraph.exactText}`).join('\n');
const chatPayload = {
  systemInstruction: { parts: [{ text: 'system' }] },
  contents: [{ role: 'user', parts: [{ text: `Context from the current document:\n"""${contextText}"""\n\nUser: hi` }] }],
  tools: [{ functionDeclarations: [{ name: 'apply_redlines' }] }]
};
const auxPayload = {
  contents: [{ parts: [{ text: buildRedlineDiffPrompt('do it', contextText) }] }],
  generationConfig: { responseMimeType: 'application/json', responseSchema: REDLINE_DIFF_SCHEMA }
};

// 1. Scenario shape: the owner's 16 prompts, each scripted and checked.
assert.equal(scenario.steps.length, 16);
assert.equal(new Set(scenario.steps.map(step => step.id)).size, 16);
for (const step of scenario.steps) {
  assert.ok(step.prompt && step.checks.length > 0 && step.replay.length >= 2, `${step.id} has prompt, checks and replay`);
  assert.ok(step.replay.at(-1).chat?.text, `${step.id} replay ends with a final chat reply`);
}

// 2. Request classification and context extraction (soft breaks continue a paragraph).
assert.equal(classifyRequest(chatPayload), 'chat');
assert.equal(classifyRequest(auxPayload), 'aux');
assert.equal(classifyRequest({ contents: [], tools: [{ google_search: {} }] }), 'other');
for (const payload of [chatPayload, auxPayload]) {
  const context = extractContext(payload);
  assert.equal(context.length, inspected.length);
  assert.deepEqual(context.map(paragraph => paragraph.text), inspected.map(paragraph => paragraph.exactText));
}

// 3. Every replay reference resolves against the original document (targets
//    named by the scenario exist before any edit).
const context = extractContext(chatPayload);
for (const step of scenario.steps) {
  for (const entry of step.replay) {
    assert.doesNotThrow(() => resolveReferences(entry.chat ?? entry.aux, context), `${step.id} references resolve`);
  }
}
const recitals = resolveReferences(scenario.steps[2].replay[0].chat, context).functionCall.args;
assert.equal(recitals.startParagraphIndex, 8, 'recitals are P8 in the original');
const headers = resolveReferences(scenario.steps[11].replay[0].chat, context).functionCall.args.paragraphIndices;
assert.deepEqual(headers, [10, 17, 24, 31, 33, 35, 45, 47, 49], 'all nine numbered headers are selected');
const signature = resolveReferences(scenario.steps[4].replay[1].aux, context).map(change => change.paragraphIndex);
assert.deepEqual(signature, [60, 61, 62, 63], 'both By: and Title: cells are targeted');

// 4. Replay queue serves in order and records desyncs instead of hanging.
const queue = createReplayQueue(scenario.steps[0]);
const first = queue.next(chatPayload);
assert.equal(first.response.candidates[0].content.parts[0].functionCall.name, 'apply_redlines');
const diff = JSON.parse(queue.next(auxPayload).response.candidates[0].content.parts[0].text);
assert.equal(diff[0].paragraphIndex, 1);
assert.equal(diff[0].anchorText, 'NON-DISCLOSURE AGREEMENT');
assert.equal(queue.next(chatPayload).response.candidates[0].content.parts[0].text, 'Done.');
const extra = queue.next(chatPayload);
assert.match(extra.error, /unscripted chat call/);
assert.deepEqual(queue.summary().desyncs, ['unscripted chat call (replay exhausted)']);
assert.equal(createReplayQueue(scenario.steps[0]).next(auxPayload).error, 'expected chat call, got aux');

// 5. Checks discriminate: on the untouched document every step's own checks
//    are not all satisfied, while style-derived formatting is read correctly.
const model = await buildDocumentModel(bytes);
assert.equal(model.paragraphs.length, inspected.length);
assert.equal(model.tables.length, 1);
assert.equal(model.comments.length, 0);
for (const [index, step] of scenario.steps.entries()) {
  const own = scoreStep(model, scenario.steps, index).filter(result => result.own);
  assert.ok(own.some(result => !result.ok), `${step.id} is not already satisfied by the source document`);
}
const titleBold = evaluateCheck(model, { type: 'format', find: 'NON-DISCLOSURE AGREEMENT', all: true, expect: { bold: true } });
assert.ok(titleBold.ok, `Strong character style counts as bold: ${titleBold.detail}`);
const bcBold = evaluateCheck(model, { type: 'format', find: 'British Columbia', paragraph: { includes: 'governed by' }, occurrence: 1, expect: { bold: true } });
assert.ok(bcBold.ok, `first British Columbia is bold via Strong: ${bcBold.detail}`);
const bcSecond = evaluateCheck(model, { type: 'format', find: 'British Columbia', paragraph: { includes: 'governed by' }, occurrence: 2, expect: { bold: false } });
assert.ok(bcSecond.ok, `second British Columbia is plain: ${bcSecond.detail}`);
const existingList = evaluateCheck(model, { type: 'list', items: [{ startsWith: 'Business plans' }, { startsWith: 'Technical data' }] });
assert.ok(existingList.ok, existingList.detail);
const archival = evaluateCheck(model, { type: 'listItem', text: { startsWith: 'This copy is to be used solely' }, level: 1 });
assert.ok(archival.ok, `2.2 is a level-1 item: ${archival.detail}`);

// Regression only re-checks earlier checks that once passed.
const passed = new Set();
const regressed = scoreStep(model, scenario.steps, 3, passed).filter(result => !result.own);
assert.equal(regressed.length, 0, 'never-passed earlier checks are not repeated');

console.log('golden_scenario_tests passed');
