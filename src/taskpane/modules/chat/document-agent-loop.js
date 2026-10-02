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
  if (candidate?.finishReason === 'STOP') return { parts: [{ text: 'Task completed successfully.' }], malformed: false };
  let diagnostic = `Finish Reason: ${candidate?.finishReason || 'NOT_FOUND'}`;
  const ratings = (Array.isArray(candidate?.safetyRatings) ? candidate.safetyRatings : []).filter(rating => rating?.probability !== 'NEGLIGIBLE');
  if (ratings.length) diagnostic += ` | Safety: ${ratings.map(rating => `${rating.category}:${rating.probability}`).join(', ')}`;
  if (candidate?.finishReason === 'SAFETY') diagnostic += ' | Content blocked by safety filters.';
  if (candidate?.finishReason === 'RECITATION') diagnostic += ' | Content blocked due to copyright/recitation filters.';
  throw new Error(`Gemini response was missing content.parts. ${diagnostic}`);
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
 * Small serial loop over the existing tool protocol. All provider, host and UI
 * work is injected. State stays observable even when a dependency throws, so
 * the taskpane can preserve history/host evidence and unlock its UI in finally.
 * WP1 deliberately retains the 11-tool protocol and existing edit executors.
 */
export async function runDocumentAgentLoop({
  state, registry, loadRegistry, requestModel, refreshContext, createCheckpoint,
  signal, requestStartTime, now = Date.now, maxLoops = 6,
  maxNoProgressToolLoops = 2, totalTimeoutMs = 180000, historyLimit = 10,
  onEvent = () => {},
}) {
  let guard = {};
  let staleContextRefreshUsed = false;
  state.loopCount = 0;
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
  emit({ type: 'run_start' });
  try {
    while (state.loopCount < maxLoops && !state.stopReason) {
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
        state.history.push({ role: 'model', parts });
        emit({ type: 'response', text: parts.find(part => part?.text && !part.thought)?.text || 'Response generated (see document for changes).' });
        state.stopReason = 'response';
        break;
      }

      registry = await loadRegistry();
      const responses = [];
      let attempted = 0;
      let successful = 0;
      let confirmed = false;
      let stale = false;
      const signatures = [];
      for (const { functionCall } of calls) {
        const { name } = functionCall;
        const args = functionCall.args || {};
        if (state.stopReason || interrupt()) {
          responses.push(createFunctionResponse(name, 'Tool not executed because this request stopped. No replay was attempted.'));
          continue;
        }
        const description = registry.describe(name, args);
        emit({ type: 'tool_start', name, ...description });
        const checkpointIndex = description.withCheckpoint ? await createCheckpoint(true, name) : undefined;
        // Cancellation during checkpoint capture must not reach an executor.
        if (interrupt()) {
          responses.push(createFunctionResponse(name, 'Tool not executed because this request stopped before editing.'));
          continue;
        }
        let execution;
        try {
          execution = await registry.execute(name, args, { docText: state.docText, docSourceBaseline: state.docSourceBaseline });
        } catch (error) {
          if (!cancelled() && error?.name !== 'AbortError') throw error;
          stop('cancelled', 'Request cancelled. Inspect the document if an edit was already in progress.');
          responses.push(createFunctionResponse(name, 'The tool was interrupted. Inspect the document before editing again.'));
          continue;
        }
        const mutation = execution.mutation;
        state.records.push({
          name, instruction: execution.instruction, result: execution.recordResult,
          success: execution.recordSuccess, ...execution.hostWriteEvidence,
        });
        responses.push(createFunctionResponse(name, execution.resultText));
        if (execution.mutating) {
          attempted++;
          confirmed ||= mutation?.written === true;
          stale ||= isNoWriteStaleContextRefusal(mutation);
          if (execution.success) successful++;
          else {
            let serialized;
            try { serialized = JSON.stringify(args); } catch { serialized = '[unserializable-args]'; }
            signatures.push(`${name}|${serialized}|${execution.resultText || ''}`);
          }
        }
        emit({ type: 'tool_end', ...execution, checkpointIndex });
        if (execution.mutating && execution.requiresInspection) {
          stop('requires_inspection', `${execution.resultText || ''} Inspect the document before editing again; do not replay this operation.`);
        }
        // A settled executor's write evidence is retained even if Cancel arrived
        // while Word was committing. Never interpret cancellation as rollback.
        if (!state.stopReason) interrupt();
      }

      const refreshStale = stale && !staleContextRefreshUsed;
      if ((!state.stopReason || ['timeout', 'cancelled'].includes(state.stopReason)) && (confirmed || (refreshStale && !cancelled()))) {
        try {
          const context = await refreshContext();
          state.docText = context.formattedText;
          state.docSourceBaseline = context.sourceBaseline;
          if (refreshStale) staleContextRefreshUsed = true;
          const reason = refreshStale ? (confirmed ? 'mixed' : 'stale-refusal') : 'confirmed-write';
          appendRefreshedDocumentContext(responses, state.docText, reason);
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
        state.history = [];
        stop('no_progress', 'Stopped to prevent a retry loop: repeated document edit attempts are failing with no applied changes.');
      }
    }
    if (!state.stopReason) {
      stop('max_turns', getMaxLoopStopMessage({ keepLooping: true, loopCount: state.loopCount, maxLoops,
        confirmedWrites: state.records.some(record => record.written === true) }));
    }
    state.history = maintainHistoryWindow(state.history, historyLimit);
    return state;
  } finally {
    emit({ type: 'run_end', reason: state.stopReason, turns: state.loopCount });
  }
}
