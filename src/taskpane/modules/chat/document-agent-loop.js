import { appendFunctionExchange, maintainHistoryWindow } from './chat-history.js';
import { advanceMutationLoopGuard, getMaxLoopStopMessage } from './mutation-loop-guard.js';
import { appendRefreshedDocumentContext, isNoWriteStaleContextRefusal } from './refreshed-document-context.js';
import { recoverFunctionCall } from './function-call-recovery.js';

/** Gemini's function-response shape; keep one response for every requested call. */
export function createFunctionResponse(name, text) {
  return { functionResponse: { name, response: { name, content: [{ text: String(text || '') }] } } };
}

function candidateParts(candidate, names, emit) {
  let parts = Array.isArray(candidate?.content?.parts) ? candidate.content.parts : [];
  const malformed = ['MALFORMED_FUNCTION_CALL', 'UNEXPECTED_TOOL_CALL'].includes(candidate?.finishReason);
  if (!Array.isArray(candidate?.content?.parts) && malformed) {
    const recovered = recoverFunctionCall(candidate.finishMessage, names);
    if (recovered) {
      parts = [{ functionCall: recovered }];
      emit({ type: 'call_recovered', name: recovered.name });
    }
  }
  if (parts.length || malformed) return { parts, malformed };
  if (candidate?.finishReason === 'STOP') return { parts: [], malformed: false };
  let diagnostic = `Finish Reason: ${candidate?.finishReason || 'NOT_FOUND'}`;
  const ratings = (Array.isArray(candidate?.safetyRatings) ? candidate.safetyRatings : []).filter(rating => rating?.probability !== 'NEGLIGIBLE');
  if (ratings.length) diagnostic += ` | Safety: ${ratings.map(rating => `${rating.category}:${rating.probability}`).join(', ')}`;
  if (candidate?.finishReason === 'SAFETY') diagnostic += ' | Content blocked by safety filters.';
  if (candidate?.finishReason === 'RECITATION') diagnostic += ' | Content blocked due to copyright/recitation filters.';
  throw new Error(`Gemini response was missing content.parts. ${diagnostic}`);
}

function recordText(record) {
  if (!record) return '';
  if (typeof record?.completionText === 'string' && record.completionText.trim()) return record.completionText;
  if (typeof record?.result?.message === 'string' && record.result.message.trim()) return record.result.message;
  if (typeof record?.resultText === 'string' && record.resultText.trim()) return record.resultText;
  if (typeof record?.result === 'string' && record.result.trim()) return record.result;
  return record?.success === true
    ? 'The last document operation was confirmed.'
    : 'The last document operation was not confirmed. Inspect the document before retrying.';
}

function latestMutationRecord(records) {
  return [...records].reverse().find(record => record.mutating === true);
}

function unresolvedFailureRecord(records) {
  let latestSuccessfulMutation = -1;
  let latestFailure = -1;
  for (let index = 0; index < records.length; index++) {
    const record = records[index];
    if (record.mutating === true && record.success === true) latestSuccessfulMutation = index;
    if (record.success !== true) latestFailure = index;
  }
  return latestFailure > latestSuccessfulMutation ? records[latestFailure] : null;
}

function withResponseText(parts, text) {
  let replaced = false;
  const next = parts.map(part => {
    if (!part?.thought && typeof part?.text === 'string') {
      replaced = true;
      return { ...part, text };
    }
    return part;
  });
  if (!replaced) next.push({ text });
  return next;
}

function exchangeRefusal(calls, descriptions) {
  if (descriptions.some(description => description.known === false) && calls.length > 1) {
    return {
      code: 'UNKNOWN_TOOL_IN_EXCHANGE',
      message: 'This exchange includes an unsupported tool call, so none of its calls were run. Use only the available tools and retry.'
    };
  }
  const mutationCount = descriptions.filter(description => description.mutating === true).length;
  if (mutationCount > 1) {
    return {
      code: 'MULTIPLE_MUTATION_CALLS',
      message: 'Only one document edit call is allowed in a model exchange. None of these calls were run; combine the requested edits into one apply_document_edits call.'
    };
  }
  if (mutationCount === 1 && calls.length > 1) {
    return {
      code: 'MIXED_MUTATION_EXCHANGE',
      message: 'An edit cannot share an exchange with inspection or another tool. None of these calls were run; inspect first, then send one apply_document_edits call.'
    };
  }
  return null;
}

function recordRefusal(state, name, description, message, code, mutating = false) {
  state.records.push({
    name,
    instruction: description?.instruction,
    result: message,
    resultText: message,
    success: false,
    mutating,
    refused: true,
    written: false,
    writeAttempted: false,
    status: 'error',
    error: { code, message },
    mutationOutcome: 'refused'
  });
}

function emitRefusalEvent(name, description, message, emit) {
  emit({ type: 'tool_start', name, ...description });
  emit({
    type: 'tool_end',
    name,
    display: { show: true },
    completionText: message,
    completionWithCheckpoint: false
  });
}

/** Only describe observed work; a failed tool record is not a successful edit. */
function completedWork(records) {
  const confirmed = records.filter(record => record.written === true);
  if (confirmed.length) return 'Confirmed edits remain in the document.';
  return records.some(record => record.success === true)
    ? 'The completed tool results remain available in this conversation.'
    : 'No document edit was confirmed.';
}

/**
 * Small serial loop over the inspect/apply protocol. All provider, host and UI
 * work is injected. State stays observable even when a dependency throws, so
 * the taskpane can preserve history/host evidence and unlock its UI in finally.
 */
export async function runDocumentAgentLoop({
  state, registry, loadRegistry, requestModel, refreshContext,
  signal, requestStartTime, now = Date.now, maxLoops = 6, maxInspectionTurns = 8,
  maxNoProgressToolLoops = 2, totalTimeoutMs = 300000, historyLimit = 10,
  onEvent = () => {},
}) {
  let guard = {};
  let staleContextRefreshUsed = false;
  state.loopCount = 0;
  // Turns whose calls are all read-only inspection draw on their own allowance first,
  // so a model that searches one term per turn still reaches its edit (KB-10).
  state.inspectionTurns = 0;
  state.stopReason = null;
  state.records ||= [];
  const emit = event => onEvent(event);
  const stop = (reason, message) => {
    state.stopReason = reason;
    if (reason === 'timeout' && state.records.length === 0 && state.history.at(-1)?.role === 'user') state.history.pop();
    emit({ type: 'stop', reason, message });
  };
  const cancelled = () => signal?.aborted === true;
  const timedOut = () => now() - requestStartTime > totalTimeoutMs;
  const interrupt = () => {
    if (cancelled()) {
      stop('cancelled', state.records.some(record => record.written || record.writeAttempted)
        ? `Request cancelled. ${completedWork(state.records)} Inspect the document before retrying.`
        : 'Request cancelled.');
      return true;
    }
    if (timedOut()) {
      stop('timeout', `The request timed out. ${completedWork(state.records)} Inspect the document before retrying.`);
      return true;
    }
    return false;
  };
  let lastRefusal = null;
  let lastTurnCompletion = null;
  emit({ type: 'run_start' });
  try {
    while (state.loopCount - state.inspectionTurns < maxLoops && !state.stopReason) {
      if (interrupt()) break;
      state.loopCount++;
      emit({ type: 'turn_start', turn: state.loopCount });
      let result;
      try {
        result = await requestModel(state.history);
      } catch (apiError) {
        if (cancelled()) { interrupt(); break; }
        // Never restart a provider conversation by replaying document tools.
        if (state.records.length > 0) {
          state.history = [];
          stop('provider_error', `${completedWork(state.records)} The AI response failed after tool execution; inspect the document before continuing.`);
          break;
        }
        throw apiError;
      }
      if (!Array.isArray(result?.candidates) || result.candidates.length === 0) {
        throw new Error('Gemini response contained no candidates.');
      }
      const candidate = result.candidates[0];
      const { parts, malformed } = candidateParts(candidate, [...registry.recoveryToolNames], emit);
      if (!parts.length && malformed) {
        state.history.push({ role: 'model', parts: [{ text: `I encountered an issue trying to use a tool (${candidate.finishReason}). Let me try a different approach.` }] });
        state.history.push({ role: 'user', parts: [{ text: 'Your previous tool call was malformed. Please try again — either rephrase the tool call with valid arguments, or respond with text only.' }] });
        continue;
      }
      const calls = parts.filter(part => part?.functionCall);
      if (!calls.length) {
        if (interrupt()) break;
        const modelText = parts.find(part => typeof part?.text === 'string' && part.text.trim() && !part.thought)?.text;
        const lastMutation = latestMutationRecord(state.records);
        const unresolvedFailure = unresolvedFailureRecord(state.records);
        const responseText = unresolvedFailure
          ? recordText(unresolvedFailure)
          : modelText || recordText(lastMutation) || 'The model stopped without a final response.';
        const responseParts = unresolvedFailure || !modelText ? withResponseText(parts, responseText) : parts;
        state.history.push({ role: 'model', parts: responseParts });
        emit({ type: 'response', text: responseText });
        state.stopReason = 'response';
        break;
      }

      registry = await loadRegistry();
      const responses = [];
      const plannedCalls = calls.map(({ functionCall }) => ({
        functionCall,
        description: registry.describe(functionCall.name, functionCall.args || {})
      }));
      const refusal = exchangeRefusal(calls, plannedCalls.map(item => item.description));
      if (refusal) {
        for (const { functionCall, description } of plannedCalls) {
          emitRefusalEvent(functionCall.name, description, refusal.message, emit);
          recordRefusal(state, functionCall.name, description, refusal.message, refusal.code, description.mutating === true);
          responses.push(createFunctionResponse(functionCall.name, refusal.message));
        }
        appendFunctionExchange(state.history, { role: 'model', parts }, { role: 'user', parts: responses });
        if (plannedCalls.some(({ description }) => description.mutating === true)) {
          const signature = `preflight-refusal:${refusal.code}:${plannedCalls.map(({ functionCall }) => functionCall.name).join(',')}`;
          const progress = advanceMutationLoopGuard(guard, {
            attemptedMutatingTools: 1,
            successfulMutatingTools: 0,
            signature,
            maxNoProgressToolLoops,
          });
          guard = progress.state;
          if (progress.countedFailure) emit({ type: 'mutation_failure', failures: guard.consecutiveNoProgressToolLoops, signatureChanged: progress.signatureChanged });
          if (progress.shouldStop) {
            stop('no_progress', 'Stopped after repeated document edit proposals were refused before execution. No host edit was attempted; send one supported apply_document_edits call after inspection.');
          }
        }
        if (state.stopReason) break;
        continue;
      }

      if (plannedCalls.every(({ description }) => description.known === true && description.mutating !== true)
        && state.inspectionTurns < maxInspectionTurns) state.inspectionTurns++;
      let attempted = 0;
      lastTurnCompletion = null;
      let successful = 0;
      let confirmed = false;
      let stale = false;
      const signatures = [];
      for (const { functionCall, description } of plannedCalls) {
        const { name } = functionCall;
        const args = functionCall.args || {};
        if (state.stopReason || interrupt()) {
          responses.push(createFunctionResponse(name, 'Tool not executed because this request stopped. No replay was attempted.'));
          continue;
        }
        emit({ type: 'tool_start', name, ...description });
        if (interrupt()) {
          responses.push(createFunctionResponse(name, 'Tool not executed because this request stopped before editing.'));
          continue;
        }
        let execution;
        try {
          execution = await registry.execute(name, args);
        } catch (error) {
          if (!cancelled() && error?.name !== 'AbortError') throw error;
          stop('cancelled', 'Request cancelled. Inspect the document if an edit was already in progress.');
          responses.push(createFunctionResponse(name, 'The tool was interrupted. Inspect the document before editing again.'));
          continue;
        }
        const mutation = execution.mutation;
        state.records.push({
          name,
          instruction: execution.instruction,
          result: execution.recordResult,
          completionText: execution.completionText,
          resultText: execution.resultText,
          checkpointIndex: execution.checkpointIndex,
          success: execution.success === true,
          recordSuccess: execution.recordSuccess,
          mutating: execution.mutating === true,
          requiresInspection: execution.requiresInspection === true,
          ...execution.hostWriteEvidence,
        });
        responses.push(createFunctionResponse(name, execution.resultText));
        if (execution.mutating) {
          attempted++;
          confirmed ||= mutation?.written === true;
          stale ||= isNoWriteStaleContextRefusal(mutation);
          if (execution.success) {
            successful++;
            if (mutation?.written === true) lastTurnCompletion = execution.completionText || mutation?.message || lastTurnCompletion;
          }
          else {
            let serialized;
            try { serialized = JSON.stringify(args); } catch { serialized = '[unserializable-args]'; }
            signatures.push(`${name}|${serialized}|${execution.resultText || ''}`);
            lastRefusal = mutation?.message || execution.rawResult?.message || lastRefusal;
          }
        }
        emit({ type: 'tool_end', ...execution, checkpointIndex: execution.checkpointIndex });
        if (execution.mutating && execution.requiresInspection) {
          const outcomeText = execution.completionText || execution.rawResult?.message || execution.resultText || 'The host could not confirm the edit outcome.';
          stop('requires_inspection', `${outcomeText} Inspect the document before editing again; do not replay this operation.`);
        }
        // A settled executor's write evidence is retained even if Cancel arrived
        // while Word was committing. Never interpret cancellation as rollback.
        if (!state.stopReason) interrupt();
      }

      const refreshStale = stale && !staleContextRefreshUsed;
      if ((!state.stopReason || ['timeout', 'cancelled'].includes(state.stopReason)) && (confirmed || (refreshStale && !cancelled()))) {
        try {
          const context = await refreshContext();
          if (refreshStale) staleContextRefreshUsed = true;
          const reason = refreshStale ? (confirmed ? 'mixed' : 'stale-refusal') : 'confirmed-write';
          appendRefreshedDocumentContext(responses, context.formattedText, reason);
          emit({ type: 'context_refresh', reason });
        } catch {
          state.history = [];
          stop('refresh_failed', confirmed
            ? 'A document edit was confirmed, but I could not refresh the document context. Inspect the current document before retrying; I stopped before another model turn.'
            : 'The edit was refused because its source context was stale, and I could not refresh the document context. I stopped before another model turn; restart the request after inspecting the document.');
        }
      }
      if (state.stopReason === 'refresh_failed') break;
      appendFunctionExchange(state.history, { role: 'model', parts }, { role: 'user', parts: responses });
      if (state.stopReason) break;
      const progress = advanceMutationLoopGuard(guard, {
        attemptedMutatingTools: attempted, successfulMutatingTools: successful,
        signature: signatures.join('||').slice(0, 2000), maxNoProgressToolLoops,
      });
      guard = progress.state;
      if (progress.countedFailure) emit({ type: 'mutation_failure', failures: guard.consecutiveNoProgressToolLoops, signatureChanged: progress.signatureChanged });
      if (progress.shouldStop) {
        // Say why it stopped, so the user is not left with only a generic stop line.
        const reason = lastRefusal ? ` Last reason: ${String(lastRefusal).replace(/^No changes made\.\s*/i, '')}` : '';
        stop('no_progress', state.records.some(record => record.written === true)
          ? `Stopped to prevent a retry loop: repeated document edit attempts are failing.${reason} Inspect the document before retrying.`
          : `Stopped to prevent a retry loop: repeated document edit attempts were refused, and no changes were made.${reason}`);
      }
    }
    if (!state.stopReason) {
      // The turn limit can land right after a verified edit, before the model's closing reply:
      // report that edit rather than a generic stop.
      stop('max_turns', lastTurnCompletion
        ? `${lastTurnCompletion} (The model reached its turn limit before writing a summary.)`
        : getMaxLoopStopMessage({ keepLooping: true, loopCount: state.loopCount - state.inspectionTurns, maxLoops,
          confirmedWrites: state.records.some(record => record.written === true) }));
    }
    state.history = maintainHistoryWindow(state.history, historyLimit);
    return state;
  } finally {
    emit({ type: 'run_end', reason: state.stopReason, turns: state.loopCount });
  }
}
