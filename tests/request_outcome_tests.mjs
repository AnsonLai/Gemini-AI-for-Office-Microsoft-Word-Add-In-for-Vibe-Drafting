import assert from 'node:assert/strict';
import { landedEditMessage, summarizeRequestOutcome, waitForSettle, wroteOrAttempted } from '../src/taskpane/modules/chat/request-outcome.js';

// What a finished (or cancelled) request did to the document, from its tool records.
{
  assert.deepEqual(summarizeRequestOutcome([]), { wrote: false, checkpointIndex: -1 });
  assert.deepEqual(summarizeRequestOutcome([{ name: 'inspect_document', success: true }]), { wrote: false, checkpointIndex: -1 });
  assert.deepEqual(summarizeRequestOutcome([{ name: 'apply_document_edits', success: false, written: false, writeAttempted: true }]),
    { wrote: false, checkpointIndex: -1 }, 'an attempt that was not confirmed is not a landed edit');
  assert.deepEqual(summarizeRequestOutcome([
    { name: 'apply_document_edits', success: true, written: true, checkpointIndex: 4 },
    { name: 'inspect_document', success: true },
    { name: 'apply_document_edits', success: true, written: true, checkpointIndex: 9 },
    { name: 'apply_document_edits', success: false, written: false },
  ]), { wrote: true, checkpointIndex: 9 }, 'the Revert button belongs to the last verified edit');
  assert.deepEqual(summarizeRequestOutcome([{ name: 'apply_document_edits', success: true, written: true }]),
    { wrote: true, checkpointIndex: -1 }, 'a write without a checkpoint has nothing to revert');
}

// A write was attempted or confirmed; reads and refusals are safe to retry.
{
  assert.equal(wroteOrAttempted([{ success: true }, { success: false, written: false, writeAttempted: false }]), false);
  assert.equal(wroteOrAttempted([{ written: true }]), true);
  assert.equal(wroteOrAttempted([{ writeAttempted: true }]), true);
}

// The note in the fresh chat after a refresh cancelled a request whose edit landed.
{
  assert.match(landedEditMessage(3), /edit from the cancelled request was applied to the document/);
  assert.match(landedEditMessage(3), /Revert/);
  assert.match(landedEditMessage(-1), /edit from the cancelled request was applied to the document/);
  assert.doesNotMatch(landedEditMessage(-1), /Revert/, 'no checkpoint, no Revert button to point at');
}

// The wait for a request to settle is limited.
{
  const value = { wrote: true, checkpointIndex: 2 };
  assert.deepEqual(await waitForSettle(Promise.resolve(value), 1000), { timedOut: false, value });
  assert.deepEqual(await waitForSettle(new Promise(() => {}), 20), { timedOut: true, value: undefined });
  const late = new Promise(resolve => setTimeout(resolve, 200, value));
  assert.equal((await waitForSettle(late, 20)).timedOut, true);
  assert.deepEqual(await late, value, 'the request can still settle after the wait gave up');
  await assert.rejects(waitForSettle(Promise.reject(new Error('settle failed')), 1000), /settle failed/);
}

console.log('PASS request outcome: landed edit, write evidence, bounded settle wait');
