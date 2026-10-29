import assert from 'node:assert/strict';
import { runDocumentAgentLoop } from '../src/taskpane/modules/chat/document-agent-loop.js';
import { createToolRegistry } from '../src/taskpane/modules/chat/tool-registry.js';
import { sanitizeHistory } from '../src/taskpane/modules/chat/chat-history.js';

const defaultEditArgs = { edits: [{ type: 'replace_text', target: 'h1', find: 'old', replace: 'new' }] };
const call = (name, args) => ({ functionCall: { name, args: args ?? (name === 'apply_document_edits' ? defaultEditArgs : {}) } });
const model = (...parts) => ({ candidates: [{ content: { parts } }] });
const final = (text = 'The requested work is complete.') => model({ thought: true, text: 'Private reasoning' }, { text });
const applied = (overrides = {}) => ({
  status: 'ok', success: true, message: 'The requested edits were applied.',
  completion: true, verified: true, written: true, writeAttempted: true, mutationOutcome: 'applied',
  receipts: [{ operationIndex: 1 }], operationResults: [{ status: 'ok' }], ...overrides,
});
const noOp = () => ({
  status: 'ok', success: true, message: 'No changes were needed; the document already matches.',
  completion: true, verified: true, written: false, writeAttempted: false, mutationOutcome: 'no_op', receipts: [], operationResults: [],
});
const refused = (code = 'INVALID_REQUEST', overrides = {}) => ({
  status: 'error', success: false, message: 'The edit was refused.',
  error: { code, message: 'The edit was refused.' }, written: false, writeAttempted: false,
  mutationOutcome: 'refused', receipts: [], operationResults: [], ...overrides,
});

const makeRegistry = executors => createToolRegistry(executors);

function harness(queue, executors = {}, overrides = {}) {
  const events = [], requests = [], checkpoints = [];
  let loads = 0, refreshes = 0;
  const registry = makeRegistry(executors, overrides.registryOptions);
  const state = {
    history: [{ role: 'user', parts: [{ text: 'Original request' }] }],
    records: [],
  };
  const options = {
    state,
    registry,
    requestStartTime: 0,
    now: () => 1,
    loadRegistry: async () => { loads++; return registry; },
    requestModel: async history => {
      requests.push(structuredClone(history));
      const next = queue.shift();
      if (next instanceof Error) throw next;
      assert.ok(next, 'the loop must not issue an unexpected model call');
      return typeof next === 'function' ? next() : next;
    },
    refreshContext: async () => {
      refreshes++;
      return { formattedText: 'Refreshed document context', sourceBaseline: [{ exactText: 'Updated', fingerprint: 'second' }] };
    },
    createCheckpoint: async (automatic, name) => { checkpoints.push({ automatic, name }); return 17; },
    onEvent: event => events.push(event),
    ...overrides,
  };
  return {
    state, events, requests, checkpoints,
    run: () => runDocumentAgentLoop(options),
    get loads() { return loads; },
    get refreshes() { return refreshes; },
  };
}

function assertPairedHistory(state) {
  assert.equal(sanitizeHistory(state.history).length, state.history.length);
}

function responseTurn(state) {
  return [...state.history].reverse().find(message => message.role === 'user'
    && message.parts.some(part => part?.functionResponse));
}

// A plain answer never loads editing tools or touches Word.
{
  const test = harness([final()]);
  await test.run();
  assert.equal(test.loads, 0);
  assert.equal(test.refreshes, 0);
  assert.equal(test.state.stopReason, 'response');
  assert.deepEqual(test.events.map(event => event.type), ['run_start', 'turn_start', 'response', 'run_end']);
  assert.equal(test.events.find(event => event.type === 'response').text, 'The requested work is complete.');
}

// Multiple inspections are allowed; one later apply receives opaque handles and owns its source checkpoint.
{
  const seen = [];
  const test = harness([
    model(call('inspect_document', { search: 'term' }), call('inspect_document', { around: 1, limit: 2 })),
    model(call('apply_document_edits', { edits: [{ type: 'replace_text', target: 'h1', find: 'old', replace: 'Updated' }] })),
    final(),
  ], {
    inspectDocument: async args => { seen.push(['inspect', args]); return { status: 'ok', success: true, message: 'Inspected current document.' }; },
    applyDocumentEdits: async args => { seen.push(['apply', args]); return applied({ checkpointIndex: 23 }); },
  });
  await test.run();
  assert.deepEqual(seen.map(item => item[0]), ['inspect', 'inspect', 'apply']);
  assert.equal(seen[2][1].edits[0].replace, 'Updated');
  assert.equal(test.checkpoints.length, 0, 'the source-aligned apply executor owns its checkpoint');
  assert.equal(test.events.find(event => event.type === 'tool_end' && event.name === 'apply_document_edits').checkpointIndex, 23);
  assert.equal(test.refreshes, 1);
  assert.equal(test.state.records.at(-1).written, true);
  assertPairedHistory(test.state);
}

// Two mutation calls are rejected as one exchange before any checkpoint or executor.
{
  let executions = 0;
  const test = harness([model(call('apply_document_edits'), call('apply_document_edits')), final('Please combine the edits.')], {
    applyDocumentEdits: async () => { executions++; return applied(); },
  });
  await test.run();
  assert.equal(executions, 0);
  assert.equal(test.checkpoints.length, 0);
  assert.equal(test.state.records.length, 2);
  assert.match(responseTurn(test.state).parts[0].functionResponse.response.content[0].text, /Only one document edit call/);
  assert.equal(responseTurn(test.state).parts.length, 2);
  assertPairedHistory(test.state);
}

// Inspection and mutation in the same exchange are rejected wholly, so handles must be inspected first.
{
  let inspections = 0, edits = 0;
  const test = harness([model(call('inspect_document'), call('apply_document_edits')), final('I will inspect first.')], {
    inspectDocument: async () => { inspections++; return { status: 'ok', success: true, message: 'Inspected.' }; },
    applyDocumentEdits: async () => { edits++; return applied(); },
  });
  await test.run();
  assert.equal(inspections, 0);
  assert.equal(edits, 0);
  assert.equal(test.checkpoints.length, 0);
  assert.match(responseTurn(test.state).parts[1].functionResponse.response.content[0].text, /inspect first/);
  assertPairedHistory(test.state);
}

// Two consecutive rejected mutation proposals consume the no-progress budget without reaching the host.
{
  let edits = 0, inspections = 0;
  const test = harness([
    model(call('apply_document_edits'), call('apply_document_edits')),
    model(call('inspect_document'), call('apply_document_edits')),
    final('This third model response must not be requested.'),
  ], {
    inspectDocument: async () => { inspections++; return { status: 'ok', success: true, message: 'Inspected.' }; },
    applyDocumentEdits: async () => { edits++; return applied(); },
  });
  await test.run();
  assert.equal(test.state.stopReason, 'no_progress');
  assert.equal(test.requests.length, 2, 'the guard stops before requesting a third model turn');
  assert.equal(edits, 0);
  assert.equal(inspections, 0);
  assert.equal(test.checkpoints.length, 0);
  assert.equal(test.state.records.filter(record => record.mutating).length, 3);
  assert.equal(test.state.records.some(record => record.written || record.writeAttempted), false, 'preflight refusal does not claim a host attempt');
  const functionCalls = test.state.history.flatMap(message => message.parts || []).filter(part => part.functionCall).length;
  const functionResponses = test.state.history.flatMap(message => message.parts || []).filter(part => part.functionResponse).length;
  assert.equal(functionResponses, 4);
  assert.equal(functionCalls, functionResponses);
  assertPairedHistory(test.state);
}

// Unknown tools are structured failures by themselves; mixed unknown proposals reject all calls.
{
  const alone = harness([model(call('legacy_edit')), final('I could not use that tool.')]);
  await alone.run();
  assert.match(alone.requests[1].at(-1).parts[0].functionResponse.response.content[0].text, /outside the supported document workflow/);
  assert.match(alone.events.find(event => event.type === 'response').text, /outside the supported document workflow/);
  assert.notEqual(alone.events.find(event => event.type === 'response').text, 'I could not use that tool.');
  assert.equal(alone.checkpoints.length, 0);
}
{
  let edits = 0;
  const mixed = harness([model(call('legacy_edit'), call('apply_document_edits')), final('Please use an available tool.')], {
    applyDocumentEdits: async () => { edits++; return applied(); },
  });
  await mixed.run();
  assert.equal(edits, 0);
  assert.equal(mixed.checkpoints.length, 0);
  assert.equal(responseTurn(mixed.state).parts.length, 2);
  assert.match(responseTurn(mixed.state).parts[0].functionResponse.response.content[0].text, /unsupported tool call/);
  assertPairedHistory(mixed.state);
}

// Source-owned checkpoint failure is covered by document_tools_tests; the loop does not capture a second snapshot.



// A failed mutation cannot be followed by a fabricated model claim of completion.
{
  const test = harness([model(call('apply_document_edits')), final('Done.')], {
    applyDocumentEdits: async () => refused('UNSUPPORTED_EDIT'),
  });
  await test.run();
  assert.equal(test.requests.length, 2);
  assert.equal(test.events.find(event => event.type === 'response').text, 'The edit was refused.');
  assert.equal(test.state.history.at(-1).parts.find(part => part.text && !part.thought)?.text, 'The edit was refused.');
  assert.notEqual(test.state.history.at(-1).parts.find(part => part.text && !part.thought)?.text, 'Done.');
}

// STOP without a response is reported honestly, including after an unsuccessful mutation.
{
  const test = harness([{ candidates: [{ finishReason: 'STOP' }] }]);
  await test.run();
  assert.equal(test.events.find(event => event.type === 'response').text, 'The model stopped without a final response.');
  assert.notEqual(test.events.find(event => event.type === 'response').text, 'Task completed successfully.');
}
{
  const test = harness([model(call('apply_document_edits')), { candidates: [{ finishReason: 'STOP' }] }], {
    applyDocumentEdits: async () => refused('UNSUPPORTED_EDIT'),
  });
  await test.run();
  assert.equal(test.events.find(event => event.type === 'response').text, 'The edit was refused.');
  assert.notEqual(test.events.find(event => event.type === 'response').text, 'Task completed successfully.');
}

// A no-op is a healthy success and resets the two-cycle no-progress guard.
{
  let count = 0;
  const test = harness([
    model(call('apply_document_edits')),
    model(call('apply_document_edits')),
    model(call('apply_document_edits')),
    model(call('apply_document_edits')),
  ], {
    applyDocumentEdits: async () => {
      count++;
      return count === 2 ? noOp() : refused('INVALID_REQUEST');
    },
  });
  await test.run();
  assert.equal(count, 4);
  assert.equal(test.events.filter(event => event.type === 'mutation_failure').at(-1).failures, 2);
  assert.equal(test.state.stopReason, 'no_progress');
}

// A stale refusal earns one refreshed source; read-only inspection does not reset failure cycles.
{
  let edits = 0, inspections = 0;
  const test = harness([
    model(call('apply_document_edits')),
    model(call('inspect_document')),
    model(call('apply_document_edits')),
  ], {
    applyDocumentEdits: async () => {
      edits++;
      return refused('STALE_DOCUMENT_CONTEXT');
    },
    inspectDocument: async () => { inspections++; return { status: 'ok', success: true, message: 'Inspected again.' }; },
  });
  await test.run();
  assert.equal(edits, 2);
  assert.equal(inspections, 1);
  assert.equal(test.refreshes, 1);
  assert.equal(Object.hasOwn(test.state, "docSourceBaseline"), false, "strong descriptors remain private to the source-bound tool session");
  assert.equal(test.state.stopReason, 'no_progress');
  assert.equal(test.requests.length, 3);
}

// Uncertain writes stop the request and retain host evidence instead of replaying.
for (const outcome of ['indeterminate', 'partial', 'applied_with_host_error']) {
  let calls = 0;
  const test = harness([model(call('apply_document_edits')), final('Retrying.')], {
    applyDocumentEdits: async () => {
      calls++;
      return refused('HOST_RESULT_UNCERTAIN', { requiresInspection: true, mutationOutcome: outcome, written: outcome !== 'indeterminate', writeAttempted: true });
    },
  });
  await test.run();
  assert.equal(test.state.stopReason, 'requires_inspection');
  assert.equal(calls, 1);
  assert.equal(test.requests.length, 1);
  assert.equal(test.refreshes, 0);
  assert.equal(test.state.records[0].mutationOutcome, outcome);
  assert.match(test.events.find(event => event.reason === 'requires_inspection').message, /Inspect the document/);
}

// Refresh failure keeps confirmed host evidence and clears stale provider history.
{
  const test = harness([model(call('apply_document_edits'))], {
    applyDocumentEdits: async () => applied(),
  }, { refreshContext: async () => { throw new Error('Host read failed'); } });
  await test.run();
  assert.equal(test.state.stopReason, 'refresh_failed');
  assert.equal(test.state.records[0].written, true);
  assert.equal(test.requests.length, 1);
  assert.deepEqual(test.state.history, []);
}

// Cancellation before a request and while a commit settles.
{
  const controller = new AbortController(); controller.abort();
  const test = harness([], {}, { signal: controller.signal });
  await test.run();
  assert.equal(test.requests.length, 0);
  assert.equal(test.state.stopReason, 'cancelled');
}

{
  const controller = new AbortController();
  const test = harness([model(call('apply_document_edits'))], {
    applyDocumentEdits: async () => { controller.abort(); return applied(); },
  }, { signal: controller.signal });
  await test.run();
  assert.equal(test.state.stopReason, 'cancelled');
  assert.equal(test.state.records[0].written, true, 'cancellation cannot erase a settled host write');
  assert.match(test.events.find(event => event.reason === 'cancelled').message, /Confirmed edits remain/);
  assert.equal(test.refreshes, 1, 'refresh a settled write even if cancellation arrived during the commit');
  assertPairedHistory(test.state);
}

// Provider errors after a tool never replay the operation; first-call errors still throw to the adapter.
{
  const test = harness([model(call('apply_document_edits')), new Error('Provider offline')], {
    applyDocumentEdits: async () => applied(),
  });
  await test.run();
  assert.equal(test.state.stopReason, 'provider_error');
  assert.equal(test.requests.length, 2);
  assert.equal(test.state.records.length, 1);
  assert.match(test.events.find(event => event.reason === 'provider_error').message, /Confirmed edits remain/);
}
{
  const test = harness([new Error('Provider offline')]);
  await assert.rejects(test.run(), /Provider offline/);
  assert.equal(test.events.at(-1).type, 'run_end');
}

// Malformed recovery only accepts complete JSON calls for the new tools.
{
  const test = harness([
    { candidates: [{ finishReason: 'MALFORMED_FUNCTION_CALL', finishMessage: 'apply_document_edits({"edits":[{"type":"replace","target":{"handle":"h1"},"content":{"text":"fixed { wording }"}}]})' }] },
    final(),
  ], { applyDocumentEdits: async () => applied() });
  await test.run();
  assert.equal(test.events.filter(event => event.type === 'call_recovered').length, 1);
  assert.equal(test.state.records.filter(record => record.mutating).length, 1);
}
{
  const test = harness(Array.from({ length: 6 }, () => ({ candidates: [{ finishReason: 'MALFORMED_FUNCTION_CALL' }] })));
  await test.run();
  assert.equal(test.state.stopReason, 'max_turns');
  assert.equal(test.requests.length, 6);
}

// A checkpointing operation is blocked before execution when its capture fails; a settled edit refreshes before timeout returns.
{
  let clock = 1, edits = 0;
  const test = harness([model(call('apply_document_edits'))], {
    applyDocumentEdits: async () => { edits++; clock = 180001; return applied(); },
  }, { now: () => clock });
  await test.run();
  assert.equal(test.state.stopReason, 'timeout');
  assert.equal(edits, 1);
  assert.equal(test.refreshes, 1);
}

// Bad provider payloads remain diagnostic failures.
{
  const test = harness([{ candidates: [{ finishReason: 'SAFETY', safetyRatings: {} }] }]);
  await assert.rejects(test.run(), /missing content.parts.*Content blocked by safety filters/);
  assert.equal(test.events.at(-1).type, 'run_end');
}

console.log('PASS: inspect/apply agent loop, single-mutation policy, truthful outcomes, history, freshness and cancellation');

// KB-10: read-only inspection turns draw on their own allowance, so a model that searches
// one term per turn still reaches its edit; the total stays bounded.
{
  const inspectTurn = term => model(call('inspect_document', { search: term, limit: 8 }));
  const test = harness([
    ...['1.', '2.', '3.', '4.', '5.', '6.', '7.'].map(inspectTurn),
    model(call('apply_document_edits')),
    final(),
  ], { inspectDocument: async () => ({ status: 'ok', targets: [] }), applyDocumentEdits: async () => applied() });
  await test.run();
  assert.equal(test.state.stopReason, 'response', 'seven inspection turns no longer exhaust the six-turn budget');
  assert.equal(test.requests.length, 9);
  assert.equal(test.state.inspectionTurns, 7);
  assert.equal(test.state.records.filter(record => record.mutating && record.written).length, 1);
}
{
  const test = harness(Array.from({ length: 14 }, (_, index) => model(call('inspect_document', { search: `term ${index}` }))),
    { inspectDocument: async () => ({ status: 'ok', targets: [] }) });
  await test.run();
  assert.equal(test.state.stopReason, 'max_turns', 'eight inspection credits plus six turns is the hard bound');
  assert.equal(test.requests.length, 14);
}
{
  // A turn that edits still uses the ordinary budget.
  const test = harness(Array.from({ length: 6 }, () => model(call('apply_document_edits'))),
    { applyDocumentEdits: async () => applied() });
  await test.run();
  assert.equal(test.state.inspectionTurns, 0);
  assert.equal(test.requests.length, 6);
}
console.log('PASS: inspection turns use a separate allowance (KB-10)');
