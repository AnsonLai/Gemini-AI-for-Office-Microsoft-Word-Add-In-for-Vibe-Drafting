import assert from 'node:assert/strict';
import { runDocumentAgentLoop } from '../src/taskpane/modules/chat/document-agent-loop.js';
import { createToolRegistry } from '../src/taskpane/modules/chat/tool-registry.js';
import { sanitizeHistory } from '../src/taskpane/modules/chat/chat-history.js';

const call = (name, args = { instruction: 'change the clause' }) => ({ functionCall: { name, args } });
const model = (...parts) => ({ candidates: [{ content: { parts } }] });
const final = () => model({ thought: true, text: 'Private reasoning' }, { text: 'Done.' });
const applied = () => ({ status: 'ok', success: true, showToUser: true, message: 'Edit applied.', written: true, writeAttempted: true, mutationOutcome: 'applied', receipts: [{ operationIndex: 1 }] });
const refused = (code = 'INVALID_REQUEST') => ({ status: 'error', success: false, showToUser: false, message: 'Edit refused.', error: { code }, written: false, writeAttempted: false, mutationOutcome: 'refused' });

function harness(queue, executors = {}, overrides = {}) {
  const events = [], requests = [], checkpoints = [];
  let loads = 0, refreshes = 0;
  const registry = createToolRegistry(executors);
  const initialBaseline = [{ index: 1, exactText: 'Original', fingerprint: 'first' }];
  const state = { history: [{ role: 'user', parts: [{ text: 'Original request' }] }], docText: '[P1] Original', docSourceBaseline: initialBaseline, records: [] };
  const options = {
    state, registry, requestStartTime: 0, now: () => 1,
    loadRegistry: async () => { loads++; return registry; },
    requestModel: async history => {
      requests.push(structuredClone(history));
      const next = queue.shift();
      if (next instanceof Error) throw next;
      assert.ok(next, 'the loop must not issue an unexpected model call');
      return typeof next === 'function' ? next() : next;
    },
    refreshContext: async () => { refreshes++; return { formattedText: '[P1] Updated', sourceBaseline: [{ index: 1, exactText: 'Updated', fingerprint: 'second' }] }; },
    createCheckpoint: async (automatic, name) => { checkpoints.push({ automatic, name }); return 'checkpoint'; },
    onEvent: event => events.push(event),
    ...overrides,
  };
  return { state, events, requests, checkpoints, initialBaseline, run: () => runDocumentAgentLoop(options),
    get loads() { return loads; }, get refreshes() { return refreshes; } };
}

// A normal answer stays on the lazy/provider boundary and never touches editing support.
{
  const test = harness([final()]);
  await test.run();
  assert.equal(test.loads, 0);
  assert.equal(test.refreshes, 0);
  assert.equal(test.state.stopReason, 'response');
  assert.deepEqual(test.events.map(event => event.type), ['run_start', 'turn_start', 'response', 'run_end']);
  assert.equal(test.events.find(event => event.type === 'response').text, 'Done.');
}

// One exchange shares its original snapshot; the next model/operation gets the refreshed one.
{
  const sources = [];
  const test = harness([model(call('apply_redlines'), call('insert_list_item', { afterParagraphIndex: 1, text: 'New item' })), model(call('apply_redlines')), final()], {
    executeRedline: async (_instruction, text, baseline) => { sources.push([text, baseline]); return applied(); },
    executeInsertListItem: async (_index, _text, _level, baseline) => { sources.push(['queued', baseline]); return applied(); },
  });
  await test.run();
  assert.equal(test.checkpoints.length, 3);
  assert.equal(test.events.filter(event => event.type === 'tool_end')[0].checkpointIndex, 'checkpoint');
  assert.equal(sources[0][1], test.initialBaseline);
  assert.equal(sources[1][1], test.initialBaseline);
  assert.equal(sources[2][1][0].exactText, 'Updated');
  assert.equal(test.refreshes, 2);
  assert.match(test.requests[1].at(-1).parts.at(-1).text, /Updated/);
  assert.equal(sanitizeHistory(test.state.history).length, test.state.history.length);
  assert.equal(test.state.records[0].written, true);
}

// A stale refusal earns one reinspection only; read-only tools do not reset failure budgets.
{
  const test = harness([model(call('apply_redlines')), model(call('perform_research')), model(call('apply_redlines'))], {
    executeRedline: async () => refused('STALE_DOCUMENT_CONTEXT'), executeResearch: async () => 'Search result',
  });
  await test.run();
  assert.equal(test.refreshes, 1);
  assert.equal(test.requests.length, 3);
  assert.equal(test.state.stopReason, 'no_progress');
  assert.equal(test.events.filter(event => event.type === 'mutation_failure').at(-1).failures, 2);
}

// Every uncertain host outcome blocks remaining queued tools and every provider continuation.
for (const outcome of ['indeterminate', 'partial', 'applied_with_host_error']) {
  let queued = 0;
  const test = harness([model(call('insert_comment'), call('apply_redlines'))], {
    executeComment: async () => ({ ...applied(), status: 'error', message: 'Host result uncertain.', written: outcome !== 'indeterminate', mutationOutcome: outcome }),
    executeRedline: async () => { queued++; return applied(); },
  });
  await test.run();
  assert.equal(test.state.stopReason, 'requires_inspection');
  assert.equal(queued, 0);
  assert.equal(test.requests.length, 1);
  assert.equal(test.refreshes, 0);
  assert.equal(test.state.records[0].mutationOutcome, outcome);
  const responses = test.state.history.at(-1).parts.filter(part => part.functionResponse);
  assert.equal(responses.length, 2, 'blocked queued calls still need paired responses');
  assert.match(responses[1].functionResponse.response.content[0].text, /not executed/);
  assert.equal(sanitizeHistory(test.state.history).length, test.state.history.length);
}

// Refresh failure preserves host evidence and clears stale provider context.
{
  const test = harness([model(call('apply_redlines'))], { executeRedline: async () => applied() }, {
    refreshContext: async () => { throw new Error('Host read failed'); },
  });
  await test.run();
  assert.equal(test.state.stopReason, 'refresh_failed');
  assert.equal(test.state.records[0].written, true);
  assert.equal(test.requests.length, 1);
  assert.deepEqual(test.state.history, []);
}

// Cancel before provider, during checkpoint capture, and while a commit is settling.
{
  const controller = new AbortController(); controller.abort();
  const test = harness([], {}, { signal: controller.signal });
  await test.run();
  assert.equal(test.requests.length, 0);
  assert.equal(test.state.stopReason, 'cancelled');
}
{
  const controller = new AbortController(); let mutations = 0;
  const test = harness([model(call('apply_redlines'), call('insert_comment'))], {
    executeRedline: async () => { mutations++; return applied(); },
    executeComment: async () => { mutations++; return applied(); },
  }, { signal: controller.signal, createCheckpoint: async () => { controller.abort(); } });
  await test.run();
  assert.equal(mutations, 0);
  assert.equal(test.state.stopReason, 'cancelled');
  assert.equal(test.state.history.at(-1).parts.filter(part => part.functionResponse).length, 2);
}
{
  const controller = new AbortController();
  const test = harness([model(call('apply_redlines'), call('insert_comment'))], {
    executeRedline: async () => { controller.abort(); return applied(); },
    executeComment: async () => { assert.fail('queued edits must stop after cancellation'); },
  }, { signal: controller.signal });
  await test.run();
  assert.equal(test.state.stopReason, 'cancelled');
  assert.equal(test.state.records[0].written, true, 'cancellation cannot erase a settled host write');
  assert.match(test.events.find(event => event.reason === 'cancelled').message, /Confirmed edits remain/);
  assert.equal(test.refreshes, 1, 'refresh settled writes even if cancellation arrived during the commit');
  assert.equal(sanitizeHistory(test.state.history).length, test.state.history.length);
}

// Provider errors after tools stop without replay or losing confirmed write observations.
{
  const test = harness([model(call('apply_redlines')), new Error('Provider offline')], { executeRedline: async () => applied() });
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

// Unknown tools return honest responses; malformed known calls use the registry name list.
{
  const test = harness([model(call('invented_tool')), final()]);
  await test.run();
  assert.match(test.requests[1].at(-1).parts[0].functionResponse.response.content[0].text, /Unknown tool/);
  assert.equal(test.checkpoints.length, 0);
}
{
  const test = harness([{ candidates: [{ finishReason: 'MALFORMED_FUNCTION_CALL', finishMessage: "apply_redlines({instruction: 'fix wording'})" }] }, final()], { executeRedline: async () => applied() });
  await test.run();
  assert.equal(test.events.filter(event => event.type === 'call_recovered').length, 1);
  assert.equal(test.state.records.length, 1);
}
{
  const test = harness(Array.from({ length: 6 }, () => ({ candidates: [{ finishReason: 'MALFORMED_FUNCTION_CALL' }] })));
  await test.run();
  assert.equal(test.state.stopReason, 'max_turns');
  assert.equal(test.requests.length, 6);
  assert.equal(test.state.history[0].parts[0].text, 'Your previous tool call was malformed. Please try again — either rephrase the tool call with valid arguments, or respond with text only.');
}

// Timeout bounds queued writes and removes a timed-out request that did no work.
{
  const test = harness([], {}, { now: () => 180001 });
  await test.run();
  assert.equal(test.state.stopReason, 'timeout');
  assert.deepEqual(test.state.history, []);
}
{
  let clock = 1, second = 0;
  const test = harness([model(call('apply_redlines'), call('insert_comment'))], {
    executeRedline: async () => { clock = 180001; return applied(); },
    executeComment: async () => { second++; return applied(); },
  }, { now: () => clock });
  await test.run();
  assert.equal(test.state.stopReason, 'timeout');
  assert.equal(second, 0);
  assert.equal(test.refreshes, 1, 'retain confirmed-write refresh before returning from timeout');
}

// Bad provider payloads remain diagnostic failures, and invalid calls never reach executors.
{
  const test = harness([{ candidates: [{ finishReason: 'SAFETY', safetyRatings: {} }] }]);
  await assert.rejects(test.run(), /missing content.parts.*Content blocked by safety filters/);
  assert.equal(test.events.at(-1).type, 'run_end');
}
{
  const test = harness([model(call('apply_redlines', {})), final()], {
    executeRedline: async () => assert.fail('invalid arguments must be refused before executor'),
  });
  await test.run();
  assert.match(test.requests[1].at(-1).parts[0].functionResponse.response.content[0].text, /instruction is required/);
  assert.equal(test.refreshes, 0);
}

console.log('PASS: portable document-agent loop, history, freshness, budgets, cancellation, and host-outcome barriers');
