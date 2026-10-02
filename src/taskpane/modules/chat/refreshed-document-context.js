/** Add the refreshed document view to the same valid function-response turn. */
export function appendRefreshedDocumentContext(functionResponses, formattedText, reason = 'confirmed-write') {
  if (!Array.isArray(functionResponses)) {
    throw new TypeError('Function responses must be an array.');
  }
  if (typeof formattedText !== 'string') {
    throw new TypeError('Refreshed document context must be text.');
  }

  let text;
  if (reason === 'mixed') {
    text = `Some edits in this step were applied, and at least one other edit was refused as stale before any write. Replan the remaining work from this refreshed document context; do not reuse the refused batch and do not redo the edits that were already applied:
"""${formattedText}"""`;
  } else if (reason === 'stale-refusal') {
    text = `The previous edit was refused before any write because its source context was stale. Replan from this refreshed document context; do not reuse the refused batch:
"""${formattedText}"""`;
  } else {
    text = `Current document context after the confirmed edit:
"""${formattedText}"""`;
  }
  functionResponses.push({ text });
  return functionResponses;
}

/** Only a proven no-write stale refusal can trigger a new planning snapshot. */
export function isNoWriteStaleContextRefusal(result) {
  return result?.error?.code === 'STALE_DOCUMENT_CONTEXT'
    && result.written === false
    && result.writeAttempted === false
    && result.mutationOutcome === 'refused';
}
