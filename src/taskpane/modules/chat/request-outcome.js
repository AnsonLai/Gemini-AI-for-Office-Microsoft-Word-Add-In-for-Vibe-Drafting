/** True when a write was confirmed or attempted; reads and refused edits leave the document alone. */
export function wroteOrAttempted(records) {
  return records.some(record => record.written === true || record.writeAttempted === true);
}

/**
 * What a request did to the document: whether an edit landed and which checkpoint the
 * Revert button belongs to (the last verified edit, as in the request's own bubble).
 */
export function summarizeRequestOutcome(records) {
  const landed = records.filter(record => record.written === true);
  const withCheckpoint = landed.filter(record => record.success === true && Number.isInteger(record.checkpointIndex));
  return {
    wrote: landed.length > 0,
    checkpointIndex: withCheckpoint.length ? withCheckpoint.at(-1).checkpointIndex : -1,
  };
}

/** The note left in a fresh chat when refresh cancelled a request whose edit still landed. */
export function landedEditMessage(checkpointIndex) {
  return checkpointIndex === -1
    ? 'An edit from the cancelled request was applied to the document. Inspect the document before continuing.'
    : 'An edit from the cancelled request was applied to the document. Use Revert to undo it.';
}

/** Wait for `promise`, but give up after `ms`. The promise itself is not cancelled. */
export async function waitForSettle(promise, ms) {
  let timer;
  const timeout = new Promise(resolve => { timer = setTimeout(resolve, ms, { timedOut: true, value: undefined }); });
  try {
    return await Promise.race([promise.then(value => ({ timedOut: false, value })), timeout]);
  } finally {
    clearTimeout(timer);
  }
}
