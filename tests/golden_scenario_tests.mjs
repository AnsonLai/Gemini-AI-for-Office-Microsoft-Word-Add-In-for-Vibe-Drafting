// Offline sanity checks for the golden scenario harness (no Word, no model).
// The scenario itself runs in desktop Word: see scripts/README.md "Golden scenario".
import './setup-xml-provider.mjs';

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { openDocx } from '@ansonlai/docx-redline-js';
import { scenario } from '../scripts/golden/nda-scenario.mjs';
import { scenario as directScenario } from '../scripts/golden/nda-direct-scenario.mjs';
import { buildDocumentModel, evaluateCheck, evaluateDirectToolEvidence, evaluateScopeDisposition, scoreStep, semanticDocumentSnapshot } from '../scripts/golden/golden-checks.mjs';
import { classifyRequest, createReplayQueue, extractContext, extractLatestFunctionResponse, extractLatestInspection, resolveReferences } from '../scripts/golden/golden-replay.mjs';
import { createFunctionResponse } from '../src/taskpane/modules/chat/document-agent-loop.js';
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
assert.equal(directScenario.steps.length, scenario.steps.length, 'the direct scenario keeps all 16 original steps');
assert.deepEqual(directScenario.steps.map(step => step.prompt), scenario.steps.map(step => step.prompt), 'direct prompts are byte-for-byte preserved');
const unsupportedDirectStepIds = [
  '03-recitals-list', '04-parties-table', '06-signature-date-row',
  '09-archival-subbullet', '11-required-disclosure-rewrite', '12-headers-to-list'
];
for (const [index, directStep] of directScenario.steps.entries()) {
  const legacyStep = scenario.steps[index];
  if (unsupportedDirectStepIds.includes(directStep.id)) {
    assert.equal(directStep.scopeDisposition.status, 'unsupported');
    assert.equal(directStep.checks[0].type, 'scopeDisposition');
    assert.deepEqual(directStep.deferredChecks, legacyStep.checks, `${directStep.id} retains legacy assertions as deferred evidence`);
    assert.equal(directStep.replay.some(entry => entry.chat?.functionCall?.name === 'apply_document_edits'), false,
      `${directStep.id} refuses without a write call`);
  } else {
    const expectedChecks = directStep.id === '15-bold-purpose-phrase'
      ? legacyStep.checks.map(check => check.label === 'purpose phrase is bold' ? { ...check, persist: false } : check)
      : legacyStep.checks;
    assert.deepEqual(directStep.checks.slice(0, -1), expectedChecks, `${directStep.id} retains its independent semantic assertions`);
    assert.equal(directStep.checks.at(-1).type, 'directToolEvidence');
    assert.equal(directStep.checks.at(-1).persist, false, 'tool evidence is scoped to the request that produced it');
    assert.equal(Object.hasOwn(directStep, 'deferredChecks'), false);
  }
}
const confidentialBulletStep = directScenario.steps.find(step => step.id === '07-confidential-bullet');
const confidentialBulletCall = confidentialBulletStep.replay.find(entry => entry.chat?.functionCall?.name === 'apply_document_edits');
assert.deepEqual(confidentialBulletCall.chat.functionCall.args.edits[0], {
  type: 'insert_list_item',
  target: { $handle: { startsWith: 'Technical data, specifications' } },
  text: 'Photographs, videos, and other recordings of prototypes and physical hardware.'
}, 'the direct overlay retains the original request with a same-level typed list insertion');
assert.equal(confidentialBulletStep.checks[0].type, 'list', 'step 7 is checked against the resulting actual list sequence');
const purposeStep = directScenario.steps.find(step => step.id === '15-bold-purpose-phrase');
const titanStep = directScenario.steps.find(step => step.id === '16-titan-rewrite');
assert.equal(purposeStep.checks.find(check => check.label === 'purpose phrase is bold').persist, false,
  'the step 15 formatting assertion is checked at its own step but not after step 16 replaces the phrase');
assert.deepEqual(titanStep.checks.slice(0, -1), scenario.steps[15].checks,
  'the step 16 Titan and phrase-removal assertions remain unchanged');

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

// 5. Direct-tool references resolve opaque handles from the latest inspection
//    functionResponse in an actual Gemini history shape. Physical indexes in
//    the result are deliberately misleading and must never become edit targets.
const inspectionResponse = targets => createFunctionResponse('inspect_document', JSON.stringify({
  status: 'ok',
  targets,
  documentEnd: { handle: 'opaque-end-latest', location: 'End of document', index: 900 }
}));
const currentPrompt = 'Please change the latest clause.';
const capturedRequest = {
  contents: [
    { role: 'user', parts: [{ text: 'User Question:\nPlease change an older clause.' }] },
    { role: 'model', parts: [{ functionCall: { name: 'apply_document_edits', args: { edits: [] } } }] },
    { role: 'user', parts: [createFunctionResponse('apply_document_edits', JSON.stringify({ requestedCount: 99, accountedCount: 99 }))] },
    { role: 'user', parts: [{ text: `User Question:\n${currentPrompt}` }] },
    { role: 'model', parts: [{ functionCall: { name: 'inspect_document', args: { search: 'clause' } } }] },
    { role: 'user', parts: [inspectionResponse([
      { handle: 'opaque-old', text: 'Older clause', index: 1 }
    ])] },
    { role: 'model', parts: [{ functionCall: { name: 'inspect_document', args: { search: 'latest clause' } } }] },
    { role: 'user', parts: [inspectionResponse([
      { handle: 'opaque-latest', text: 'Latest target clause', index: 777 }
    ])] }
  ],
  tools: [{ functionDeclarations: [{ name: 'apply_document_edits' }] }]
};
const capturedInspection = extractLatestInspection(capturedRequest);
assert.equal(capturedInspection.targets[0].handle, 'opaque-latest');
const directQueue = createReplayQueue({ prompt: currentPrompt, replay: [{ chat: { functionCall: {
  name: 'apply_document_edits',
  args: { edits: [
    { type: 'replace_text', target: { $handle: { includes: 'Latest target clause' } }, find: 'Latest', replace: 'Current' },
    { type: 'append_paragraph', target: { $end: true }, text: 'A final plain paragraph.' }
  ] }
} } }] });
const directResult = directQueue.next(capturedRequest);
const directEdits = directResult.response.candidates[0].content.parts[0].functionCall.args.edits;
assert.equal(directEdits[0].target, 'opaque-latest', 'a semantic matcher resolves to the latest inspected opaque handle');
assert.equal(directEdits[1].target, 'opaque-end-latest', '$end resolves to the latest inspected document-end handle');
assert.equal(directEdits.some(edit => Number.isInteger(edit.target)), false, 'physical paragraph indexes are never edit targets');
assert.deepEqual(directQueue.summary().desyncs, []);
const capturedReceipt = { requestedCount: 2, accountedCount: 2, counters: { sourceReads: 2, contentWrites: 1, verificationReads: 1 } };
const receiptRequest = {
  ...capturedRequest,
  contents: [...capturedRequest.contents, {
    role: 'user', parts: [createFunctionResponse('apply_document_edits', JSON.stringify(capturedReceipt))]
  }]
};
assert.deepEqual(extractLatestFunctionResponse(receiptRequest, 'apply_document_edits', currentPrompt), capturedReceipt,
  'per-step evidence uses the latest actual function response and ignores older request history');
const failedInspectionRequest = {
  ...capturedRequest,
  contents: [...capturedRequest.contents, { role: 'user', parts: [createFunctionResponse('inspect_document', JSON.stringify({
    status: 'error', error: { code: 'STALE_CURSOR', message: 'Inspect again.' }, targets: []
  }))] }]
};
assert.throws(() => resolveReferences({ $handle: { includes: 'Older clause' } }, [], failedInspectionRequest, currentPrompt), /latest inspect_document response did not contain a successful JSON result/,
  'an older successful inspection cannot rescue the latest failed inspection');

// Scope refusals require an explicit user-facing refusal, no write call, and
// equality across every accepted-view semantic component (not just the target).
const beforeScope = {
  paragraphs: [{ text: 'Untouched clause', runs: [{ text: 'Untouched clause', format: { bold: false } }], list: null, table: null }],
  tables: [], comments: [], commentAnchors: [], revisionCount: 0, revisionAuthors: {}
};
const afterScope = structuredClone(beforeScope);
const refusedReplay = {
  served: [{ kind: 'chat', part: { text: 'I can’t convert list structure; it is unsupported and unchanged.' } }],
  requestCounts: { chat: 1, aux: 0, other: 0 }
};
const refusalCheck = directScenario.steps[2].scopeDisposition;
assert.equal(evaluateScopeDisposition(beforeScope, afterScope, [
  'User:\n\ncan you turn the recitals into a proper ordered list?',
  'Gemini:\n\nI can’t convert list structure; it is unsupported and unchanged.'
], refusedReplay, refusalCheck).ok, true);
const sameFormat = { bold: false, italic: false, underline: false, strikethrough: false, highlight: false };
const splitRunsModel = {
  ...beforeScope,
  paragraphs: [{
    text: 'Waiver. No failure or delay.',
    runs: [
      { text: 'Waiver.', format: { ...sameFormat, bold: true } },
      { text: ' No failure or delay', format: sameFormat },
      { text: '.', format: { ...sameFormat } }
    ],
    list: { numId: '4', level: 0, numFmt: 'decimal' }, table: null
  }]
};
const mergedRunsModel = structuredClone(splitRunsModel);
mergedRunsModel.paragraphs[0].runs = [
  { text: 'Waiver.', format: { ...sameFormat, bold: true } },
  { text: ' No failure or delay.', format: { ...sameFormat } }
];
assert.equal(semanticDocumentSnapshot(splitRunsModel), semanticDocumentSnapshot(mergedRunsModel),
  'adjacent accepted runs with identical effective formatting are semantically equal when Word merges them');
assert.equal(evaluateScopeDisposition(splitRunsModel, mergedRunsModel,
  ['Gemini:\n\nI can’t convert list structure; it is unsupported and unchanged.'], refusedReplay, refusalCheck).ok, true,
  'a Word run merge does not fail an unchanged-document refusal check');
const changedRunFormat = structuredClone(mergedRunsModel);
changedRunFormat.paragraphs[0].runs[1].format.bold = true;
assert.notEqual(semanticDocumentSnapshot(splitRunsModel), semanticDocumentSnapshot(changedRunFormat),
  'a real accepted formatting change remains visible in the semantic snapshot');
assert.equal(evaluateScopeDisposition(splitRunsModel, changedRunFormat,
  ['Gemini:\n\nI can’t convert list structure; it is unsupported and unchanged.'], refusedReplay, refusalCheck).ok, false,
  'a format change cannot be hidden by adjacent-run coalescing');
const changedRunText = structuredClone(mergedRunsModel);
changedRunText.paragraphs[0].text += ' amended';
changedRunText.paragraphs[0].runs[1].text += ' amended';
assert.notEqual(semanticDocumentSnapshot(splitRunsModel), semanticDocumentSnapshot(changedRunText),
  'changed accepted text remains visible in the semantic snapshot');
assert.equal(evaluateScopeDisposition(splitRunsModel, changedRunText,
  ['Gemini:\n\nI can’t convert list structure; it is unsupported and unchanged.'], refusedReplay, refusalCheck).ok, false,
  'changed text fails an unchanged-document refusal check');
const attemptedWrite = { served: [{ kind: 'chat', part: { functionCall: { name: 'apply_document_edits' } } }] };
attemptedWrite.requestCounts = { chat: 1, aux: 0, other: 0 };
assert.equal(evaluateScopeDisposition(beforeScope, afterScope, ['Gemini:\n\nI can’t convert list structure; it is unsupported.'],
  attemptedWrite, refusalCheck).ok, false, 'an explicit refusal cannot mask a write-tool call');
const changedScope = structuredClone(beforeScope);
changedScope.tables.push({ rows: [['new table']] });
assert.equal(evaluateScopeDisposition(beforeScope, changedScope, ['Gemini:\n\nI can’t convert list structure; it is unsupported.'],
  refusedReplay, refusalCheck).ok, false, 'a semantic change anywhere in the document fails scope verification');
const commitReceipt = {
  status: 'ok', success: true, completion: true, verified: true,
  written: true, writeAttempted: true, checkpointIndex: 1,
  requestedCount: 2, accountedCount: 2,
  counters: { sourceReads: 2, contentWrites: 1, verificationReads: 1 }
};
const directEvidence = {
  served: [{ kind: 'chat', part: { functionCall: { name: 'apply_document_edits' } } }],
  functionResponseHistory: { apply_document_edits: [commitReceipt] },
  requestCounts: { chat: 2, aux: 0, other: 0 }
};
assert.equal(evaluateDirectToolEvidence(directEvidence).ok, true, 'direct edit evidence requires an accounted single-write commit');
assert.equal(evaluateDirectToolEvidence({ ...directEvidence, requestCounts: { chat: 2, aux: 1, other: 0 } }).ok, false,
  'nested model calls fail the direct protocol gate');
assert.equal(evaluateDirectToolEvidence({
  ...directEvidence,
  functionResponseHistory: { apply_document_edits: [{ ...commitReceipt, counters: { ...commitReceipt.counters, contentWrites: 2 } }] }
}).ok, false, 'more than one content write fails the direct batch gate');

// 6. Checks discriminate: on the untouched document every step's own checks
//    are not all satisfied, while style-derived formatting is read correctly.
const model = await buildDocumentModel(bytes);
assert.equal(model.paragraphs.length, inspected.length);
assert.equal(model.tables.length, 1);
assert.equal(model.comments.length, 0);
const purposeStepIndex = directScenario.steps.findIndex(step => step.id === '15-bold-purpose-phrase');
assert.ok(scoreStep(model, directScenario.steps, purposeStepIndex).some(result => result.own
  && result.label === 'purpose phrase is bold'), 'step 15 still evaluates its formatting assertion at that step');
assert.equal(scoreStep(model, directScenario.steps, purposeStepIndex + 1).some(result => result.step === '15-bold-purpose-phrase'), false,
  'step 16 does not regress step 15 after intentionally replacing the phrase');
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
