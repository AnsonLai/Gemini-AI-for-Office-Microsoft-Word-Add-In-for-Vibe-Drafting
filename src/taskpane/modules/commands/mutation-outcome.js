/** Consumer host observations; these are deliberately separate from engine receipts. */
export function createMutationObserver(getSignal) {
  let attempted = false;
  let pending = false;
  let confirmed = 0;
  const receipts = [];
  const operationResults = [];
  return {
    attempt() {
      if (getSignal?.()?.aborted) throw Object.assign(new Error('The request was cancelled before document editing.'), { name: 'AbortError' });
      attempted = true; pending = true;
    },
    confirm() { if (pending) { confirmed++; pending = false; } },
    async sync(context) {
      await context.sync();
      if (pending) { confirmed++; pending = false; }
    },
    observe(result) {
      operationResults.push({ status: result?.status, error: result?.error, hostError: result?.hostError,
        rolledBack: result?.rolledBack, written: result?.written,
        writeAttempted: result?.writeAttempted, mutationOutcome: result?.mutationOutcome });
      if (Array.isArray(result?.receipts)) receipts.push(...result.receipts);
      else if (result?.receipt) receipts.push(result.receipt);
      if (result?.writeAttempted) attempted = true;
      if (result?.written) confirmed++;
      if (result?.writeAttempted && !result?.written) pending = true;
    },
    result(error = null) {
      const lastOperation = operationResults[operationResults.length - 1];
      const mutationOutcome = error
        ? pending ? (confirmed ? 'partial' : 'indeterminate') : confirmed ? lastOperation?.status === 'error' ? 'partial' : 'applied_with_host_error' : lastOperation?.rolledBack ? 'rolled_back' : lastOperation?.mutationOutcome === 'prepared' ? 'prepared' : 'refused'
        : pending ? (confirmed ? 'partial' : 'indeterminate') : confirmed ? 'applied' : 'noop';
      return { status: error || pending ? 'error' : 'ok', success: !error && !pending, written: confirmed > 0,
        writeAttempted: attempted, confirmedHostWrites: confirmed, mutationOutcome, receipts, operationResults,
        ...(error ? { error: { ...(error?.code ? { code: error.code } : {}), message: error?.message || String(error) } } : {}) };
    }
  };
}

export function requiresMutationInspection(result) {
  return ['indeterminate', 'partial', 'applied_with_host_error'].includes(result?.mutationOutcome);
}
