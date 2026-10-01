/** Update the no-progress counter only when a mutation was attempted or confirmed. */
export function advanceMutationLoopGuard(state = {}, outcome = {}) {
  const current = {
    consecutiveNoProgressToolLoops: Number.isInteger(state.consecutiveNoProgressToolLoops)
      ? state.consecutiveNoProgressToolLoops : 0,
    lastNoProgressSignature: typeof state.lastNoProgressSignature === 'string'
      ? state.lastNoProgressSignature : ''
  };
  const attempted = Number.isInteger(outcome.attemptedMutatingTools)
    ? outcome.attemptedMutatingTools : 0;
  const successful = Number.isInteger(outcome.successfulMutatingTools)
    ? outcome.successfulMutatingTools : 0;

  if (successful > 0) {
    return {
      state: { consecutiveNoProgressToolLoops: 0, lastNoProgressSignature: '' },
      reset: true,
      countedFailure: false,
      signatureChanged: false,
      shouldStop: false
    };
  }

  // Navigation, research, and other read-only exchanges do not earn a reset.
  if (attempted === 0) {
    return {
      state: current,
      reset: false,
      countedFailure: false,
      signatureChanged: false,
      shouldStop: false
    };
  }

  const signature = typeof outcome.signature === 'string' ? outcome.signature : '';
  const signatureChanged = !!(current.lastNoProgressSignature && signature
    && signature !== current.lastNoProgressSignature);
  const consecutiveNoProgressToolLoops = current.consecutiveNoProgressToolLoops + 1;
  const maximum = Number.isInteger(outcome.maxNoProgressToolLoops) && outcome.maxNoProgressToolLoops > 0
    ? outcome.maxNoProgressToolLoops : 2;

  return {
    state: { consecutiveNoProgressToolLoops, lastNoProgressSignature: signature },
    reset: false,
    countedFailure: true,
    signatureChanged,
    shouldStop: consecutiveNoProgressToolLoops >= maximum
  };
}

/** Return a terminal message when the model loop exhausts its request budget. */
export function getMaxLoopStopMessage({ keepLooping, loopCount, maxLoops, confirmedWrites }) {
  if (keepLooping !== true || !Number.isInteger(loopCount) || !Number.isInteger(maxLoops)
    || loopCount < maxLoops) {
    return null;
  }
  return confirmedWrites === true
    ? 'Stopped after the maximum number of model turns. Confirmed edits remain in the document; inspect them before asking to continue.'
    : 'Stopped after the maximum number of model turns without a final response. Please review the document before retrying.';
}
