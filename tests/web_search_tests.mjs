import assert from 'node:assert/strict';
import { createWebSearch, WEB_RESULTS_NOTE } from '../src/taskpane/modules/chat/web-search.js';
import { createToolRegistry } from '../src/taskpane/modules/chat/tool-registry.js';
import { runDocumentAgentLoop } from '../src/taskpane/modules/chat/document-agent-loop.js';
import { GeminiRequestError } from '../src/taskpane/modules/chat/gemini-client.js';

const grounded = (text, chunks = []) => ({ candidates: [{
  content: { parts: [{ text: 'thinking…', thought: true }, ...[].concat(text).map(value => ({ text: value }))] },
  groundingMetadata: { groundingChunks: chunks },
}] });

// One grounded request with the built-in search tool; answer joined across parts, sources kept.
{
  const payloads = [];
  const search = createWebSearch({ request: async payload => { payloads.push(payload); return grounded(['Acme Corp was ', 'incorporated in 2019.'], [
    { web: { uri: 'https://example.com/acme', title: 'Acme history' } },
    { web: { uri: 'https://example.com/acme', title: 'Duplicate' } },
    { web: { uri: 'javascript:alert(1)', title: 'Not a web link' } },
    { retrievedContext: { uri: 'https://example.com/other' } },
  ]); }, safetySettings: [{ category: 'HARM_CATEGORY_HARASSMENT', threshold: 'BLOCK_NONE' }] });
  const result = await search({ query: '  When was Acme Corp incorporated?  ' });
  assert.deepEqual(result, {
    status: 'ok', query: 'When was Acme Corp incorporated?', answer: 'Acme Corp was incorporated in 2019.',
    sources: [{ title: 'Acme history', uri: 'https://example.com/acme' }], note: WEB_RESULTS_NOTE,
  });
  assert.equal(payloads.length, 1);
  assert.deepEqual(payloads[0].tools, [{ google_search: {} }], 'the grounded call carries no function declarations');
  assert.match(payloads[0].contents[0].parts[0].text, /Question: When was Acme Corp incorporated\?$/);
  assert.equal(payloads[0].safetySettings.length, 1);
}

// Long answers are truncated; invalid queries never reach the provider.
{
  let sent = 0;
  const search = createWebSearch({ request: async () => { sent++; return grounded('x'.repeat(5000)); } });
  const long = await search({ query: 'q' });
  assert.equal(long.answer.length, 4000);
  assert.equal(long.answerTruncated, true);
  for (const query of [undefined, '', '  ', 42, 'x'.repeat(301)]) {
    assert.equal((await search({ query })).error.code, 'INVALID_SEARCH_QUERY');
  }
  assert.equal(sent, 1);
}

// Provider failures and empty answers become tool errors; cancellation propagates.
{
  const failing = createWebSearch({ request: async () => { throw new GeminiRequestError('REQUEST_HTTP', 'Gemini rate limit reached. Wait before trying again.', 429); } });
  assert.deepEqual(await failing({ query: 'q' }), { status: 'error', error: { code: 'WEB_SEARCH_FAILED', message: 'Gemini rate limit reached. Wait before trying again.' } });
  const empty = createWebSearch({ request: async () => ({ candidates: [{ finishReason: 'SAFETY' }] }) });
  assert.equal((await empty({ query: 'q' })).error.code, 'WEB_SEARCH_EMPTY');
  const cancelled = createWebSearch({ request: async () => { throw new GeminiRequestError('REQUEST_CANCELLED', 'Request cancelled by user'); } });
  await assert.rejects(cancelled({ query: 'q' }), { code: 'REQUEST_CANCELLED' });
}

// In the loop: a search between two identical refused edits does not let the edit be resent.
{
  const editCall = { name: 'apply_document_edits', args: { edits: [{ type: 'replace_text', target: 'h:1', find: 'Old', replace: 'New' }] } };
  const turns = [
    [{ functionCall: editCall }],
    [{ functionCall: { name: 'search_web', args: { query: 'Is New the right term?' } } }],
    [{ functionCall: editCall }],
    [{ text: 'I could not make that change.' }],
  ];
  let applyCalls = 0, searchCalls = 0;
  const registry = createToolRegistry({
    searchWeb: async () => { searchCalls++; return { status: 'ok', answer: 'Yes.', sources: [], note: WEB_RESULTS_NOTE }; },
    applyDocumentEdits: async () => { applyCalls++; return { status: 'error', success: false, written: false, writeAttempted: false, mutationOutcome: 'refused', error: { code: 'TARGET_NOT_FOUND', message: 'No match.' }, message: 'No changes made. No match.' }; },
  });
  const state = { history: [{ role: 'user', parts: [{ text: 'Fix the term.' }] }], records: [] };
  await runDocumentAgentLoop({
    state, registry, loadRegistry: async () => registry,
    requestModel: async () => ({ candidates: [{ content: { parts: turns.shift() } }] }),
    refreshContext: async () => ({ formattedText: '' }), requestStartTime: 0, now: () => 1, maxNoProgressToolLoops: 5,
  });
  assert.equal(searchCalls, 1);
  assert.equal(applyCalls, 1, 'the repeated refused edit is blocked even after a web search');
  assert.equal(state.records.at(-1).error?.code ?? JSON.parse(state.records.at(-1).resultText).error.code, 'REPEATED_REFUSED_EDITS');
}

// A failed search is read-only: it never replaces the model's own reply.
{
  const turns = [
    [{ functionCall: { name: 'search_web', args: { query: 'Current filing fee?' } } }],
    [{ text: 'I could not confirm the fee online, so I left the clause unchanged.' }],
  ];
  const registry = createToolRegistry({ searchWeb: async () => ({ status: 'error', error: { code: 'WEB_SEARCH_FAILED', message: 'Gemini rate limit reached.' } }) });
  const events = [];
  const state = { history: [{ role: 'user', parts: [{ text: 'Check the filing fee.' }] }], records: [] };
  await runDocumentAgentLoop({
    state, registry, loadRegistry: async () => registry,
    requestModel: async () => ({ candidates: [{ content: { parts: turns.shift() } }] }),
    refreshContext: async () => ({ formattedText: '' }), requestStartTime: 0, now: () => 1,
    onEvent: event => events.push(event),
  });
  assert.equal(state.stopReason, 'response');
  assert.equal(events.find(event => event.type === 'response').text, 'I could not confirm the fee online, so I left the clause unchanged.');
  assert.equal(state.records[0].readOnly, true);
}

console.log('PASS: search_web grounds one separate request, labels results, and cannot unlock refused edits');
