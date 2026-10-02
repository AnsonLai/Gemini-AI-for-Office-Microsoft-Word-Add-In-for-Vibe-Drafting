// Offline fixtures for direct live model evidence and isolated model-profile selection.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { evaluateDirectToolEvidence, evaluateScopeDisposition } from '../scripts/golden/golden-checks.mjs';
import { createStepObservationCollector, resolveGoldenModelSettings } from '../scripts/golden/golden-replay.mjs';

const prompt = 'Add the requested sentence to the agreement.';
const inspectCall = { functionCall: { name: 'inspect_document', args: { search: 'agreement' } } };
const applyCall = { functionCall: { name: 'apply_document_edits', args: { edits: [{ type: 'replace_text' }] } } };
const completeReceipt = {
  status: 'ok',
  requestedCount: 1,
  accountedCount: 1,
  success: true,
  completion: true,
  verified: true,
  written: true,
  writeAttempted: true,
  checkpointIndex: 1,
  counters: { sourceReads: 2, contentWrites: 1, verificationReads: 1 }
};
const functionResponse = (name, value) => ({
  functionResponse: { name, response: { content: [{ text: JSON.stringify(value) }] } }
});
const chatPayload = contents => ({
  tools: [{ functionDeclarations: [{ name: 'inspect_document' }, { name: 'apply_document_edits' }] }],
  contents
});
const response = (...parts) => ({ candidates: [{ content: { role: 'model', parts }, finishReason: 'STOP' }] });
const previousHistory = [
  { role: 'user', parts: [{ text: 'Earlier unrelated request.' }] },
  { role: 'model', parts: [{ functionCall: { name: 'apply_document_edits', args: {} } }] },
  { role: 'user', parts: [functionResponse('apply_document_edits', completeReceipt)] }
];
const currentTurn = [
  ...previousHistory,
  { role: 'user', parts: [{ text: prompt }] }
];

// Successful direct edit: only real model function calls and the actual
// subsequent Gemini functionResponse history are counted and paired.
const observer = createStepObservationCollector(prompt);
observer.record({ payload: chatPayload(currentTurn), response: JSON.stringify(response(inspectCall)), status: 200 });
const afterInspect = [
  ...currentTurn,
  { role: 'model', parts: [inspectCall] },
  { role: 'user', parts: [functionResponse('inspect_document', { status: 'ok', targets: [] })] }
];
observer.record({ payload: chatPayload(afterInspect), response: response(applyCall), status: 200 });
const afterApply = [
  ...afterInspect,
  { role: 'model', parts: [applyCall] },
  { role: 'user', parts: [functionResponse('apply_document_edits', completeReceipt)] }
];
observer.record({ payload: chatPayload(afterApply), response: response({ text: 'Done.' }), status: 200 });
const observed = observer.summary();
assert.deepEqual(observed.requestCounts, { chat: 3, aux: 0, other: 0 });
assert.equal(observed.served.filter(item => item.part.functionCall?.name === 'apply_document_edits').length, 1);
assert.deepEqual(observed.functionResponseHistory.apply_document_edits, [completeReceipt],
  'the old step receipt in model history is excluded and the current receipt is parsed from the latest request');
assert.equal(evaluateDirectToolEvidence(observed).ok, true);

// A function call in a model response is not itself a receipt. If the final
// captured history was clipped before the tool result, the gate must fail.
const truncated = createStepObservationCollector(prompt);
truncated.record({ payload: chatPayload(currentTurn), response: response(applyCall), status: 200 });
truncated.record({ payload: chatPayload([
  ...currentTurn,
  { role: 'model', parts: [applyCall] }
]), response: response({ text: 'Done.' }), status: 200 });
assert.equal(truncated.summary().functionResponseHistory.apply_document_edits, undefined);
assert.equal(evaluateDirectToolEvidence(truncated.summary()).ok, false,
  'a missing request-history receipt is not synthesized from tool arguments or model output');

// Candidate output carrying a response-shaped part is not provider tool evidence,
// and unsuccessful/invalid responses cannot contribute served calls.
const fabricated = createStepObservationCollector(prompt);
fabricated.record({ payload: chatPayload(currentTurn), response: response(functionResponse('apply_document_edits', completeReceipt)), status: 200 });
fabricated.record({ payload: chatPayload(currentTurn), response: response(applyCall), status: 503 });
fabricated.record({ payload: chatPayload(currentTurn), response: 'not json', status: 200 });
assert.deepEqual(fabricated.summary().functionResponseHistory, {});
assert.equal(fabricated.summary().served.some(item => item.part.functionCall?.name === 'apply_document_edits'), false);
assert.equal(fabricated.summary().invalidResponses, 2);

// Live refusal checks use captured provider calls/counts too. An observed nested
// JSON or research request invalidates an otherwise plausible refusal.
const unchanged = {
  paragraphs: [{ text: 'Original wording.', runs: [] }],
  tables: [], comments: [], commentAnchors: [], revisionCount: 0, revisionAuthors: {}
};
const refusal = createStepObservationCollector(prompt);
refusal.record({ payload: chatPayload(currentTurn), response: response({ text: 'I cannot convert the existing list; that is unsupported.' }), status: 200 });
const refusalEvidence = refusal.summary();
const disposition = { status: 'unsupported', refusalPattern: '(?:cannot|unsupported).*(?:list|convert)' };
assert.equal(evaluateScopeDisposition(unchanged, unchanged, ['Assistant: I cannot convert the existing list; that is unsupported.'], refusalEvidence, disposition).ok, true);
refusal.record({ payload: { generationConfig: { responseMimeType: 'application/json' }, contents: currentTurn },
  response: response({ text: '{}' }), status: 200 });
assert.equal(evaluateScopeDisposition(unchanged, unchanged, ['Assistant: I cannot convert the existing list; that is unsupported.'], refusal.summary(), disposition).ok, false,
  'nested auxiliary model requests are visible to the live refusal gate');

assert.deepEqual(resolveGoldenModelSettings(), {
  profile: 'configured', fast: 'gemini-flash-latest', slow: 'gemini-2.5-pro'
});
assert.deepEqual(resolveGoldenModelSettings({ profile: 'configured', fastModel: 'gemini-flash-latest', slowModel: 'gemini-2.5-pro' }), {
  profile: 'configured', fast: 'gemini-flash-latest', slow: 'gemini-2.5-pro'
});
assert.deepEqual(resolveGoldenModelSettings({ profile: 'fast' }), {
  profile: 'fast', fast: 'gemini-flash-latest', slow: 'gemini-flash-latest'
});
assert.deepEqual(resolveGoldenModelSettings({ profile: 'slow' }), {
  profile: 'slow', fast: 'gemini-2.5-pro', slow: 'gemini-2.5-pro'
});
assert.throws(() => resolveGoldenModelSettings({ profile: 'fast', slowModel: 'gemini-2.5-pro' }), /cannot be combined/);
assert.throws(() => resolveGoldenModelSettings({ profile: 'unknown' }), /must be configured/);

// Keep the user-facing runner defaults and lifecycle/logging guarantees visible
// without importing its intentional desktop-Word server side effects in tests.
const runnerSource = readFileSync(new URL('../scripts/run-golden-scenario.mjs', import.meta.url), 'utf8');
const driverSource = readFileSync(new URL('../scripts/golden/golden-driver.js', import.meta.url), 'utf8');
assert.match(runnerSource, /argument\('--protocol', 'direct'\)/);
assert.match(runnerSource, /closeAllConnections\?\.\(\)/);
assert.match(driverSource, /await pendingModelLogs/);
assert.match(driverSource, /await recordLiveExchange/);

console.log('live_golden_observation_tests passed');
