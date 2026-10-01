import assert from 'node:assert/strict';
import { advanceMutationLoopGuard, getMaxLoopStopMessage } from '../src/taskpane/modules/chat/mutation-loop-guard.js';

let state = { consecutiveNoProgressToolLoops: 0, lastNoProgressSignature: '' };
let outcome = advanceMutationLoopGuard(state, {
  attemptedMutatingTools: 1,
  successfulMutatingTools: 0,
  signature: 'apply_redlines|failed'
});
assert.equal(outcome.countedFailure, true);
assert.equal(outcome.state.consecutiveNoProgressToolLoops, 1);
state = outcome.state;

outcome = advanceMutationLoopGuard(state, {
  attemptedMutatingTools: 0,
  successfulMutatingTools: 0
});
assert.equal(outcome.reset, false, 'a read-only turn must not reset the failed-mutation budget');
assert.equal(outcome.state.consecutiveNoProgressToolLoops, 1);
state = outcome.state;

outcome = advanceMutationLoopGuard(state, {
  attemptedMutatingTools: 1,
  successfulMutatingTools: 0,
  signature: 'apply_redlines|failed'
});
assert.equal(outcome.shouldStop, true, 'a second failure after a read-only turn must stop the retry loop');
assert.equal(outcome.state.consecutiveNoProgressToolLoops, 2);
state = outcome.state;

outcome = advanceMutationLoopGuard(state, {
  attemptedMutatingTools: 1,
  successfulMutatingTools: 1,
  signature: 'apply_redlines|written'
});
assert.equal(outcome.reset, true, 'a confirmed mutation starts a fresh no-progress budget');
assert.deepEqual(outcome.state, { consecutiveNoProgressToolLoops: 0, lastNoProgressSignature: '' });

assert.equal(getMaxLoopStopMessage({ keepLooping: true, loopCount: 6, maxLoops: 6, confirmedWrites: true }),
  'Stopped after the maximum number of model turns. Confirmed edits remain in the document; inspect them before asking to continue.');
assert.equal(getMaxLoopStopMessage({ keepLooping: false, loopCount: 6, maxLoops: 6, confirmedWrites: true }), null,
  'a normal completion at the limit must not be reported as a forced stop');

console.log('PASS: mutation loop guard tests');
