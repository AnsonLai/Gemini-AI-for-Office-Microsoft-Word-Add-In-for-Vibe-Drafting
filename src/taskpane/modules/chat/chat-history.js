/**
 * Maintains a rolling window of chat history while preserving function call/response pairs
 */
function maintainHistoryWindow(history, maxMessages) {
  if (history.length <= maxMessages) {
    return history;
  }

  const limit = Math.max(0, Math.floor(Number(maxMessages) || 0));
  if (limit === 0) return [];

  // Canonicalize first so the window is selected from complete exchanges. In
  // particular, do not select a suffix that starts with a function response
  // and then rely on validation to discard the useful history after it.
  const validHistory = validateHistoryPairs(history);
  const isRequest = (msg) => {
    if (!msg || msg.role !== "user") return false;
    const parts = msg.parts || [];
    if (parts.some((part) => part && part.functionResponse)) return false;
    return parts.some((part) => typeof (part && part.text) === "string" && part.text.trim().length > 0);
  };

  let latestRequestIndex = -1;
  for (let i = validHistory.length - 1; i >= 0; i--) {
    if (isRequest(validHistory[i])) {
      latestRequestIndex = i;
      break;
    }
  }

  // A text request is the useful conversational anchor. If none exists, an
  // empty history is safer than returning a window that starts with a model
  // turn or an orphaned function response.
  if (latestRequestIndex < 0) return [];

  if (validHistory.length <= limit) return validHistory;

  const pairs = [];
  for (let i = 0; i < validHistory.length - 1; i++) {
    const msg = validHistory[i];
    if (msg.role === "model" && (msg.parts || []).some((part) => part && part.functionCall)) {
      let requestIndex = -1;
      for (let j = i - 1; j >= 0; j--) {
        if (isRequest(validHistory[j])) {
          requestIndex = j;
          break;
        }
      }
      if (requestIndex >= 0) {
        pairs.push({ callIndex: i, responseIndex: i + 1, requestIndex });
      }
      i++; // The validator guarantees that this next turn is its response.
    }
  }

  // Keep the latest real user request even when it follows a long tool turn.
  // Each retained exchange also keeps the most recent real user request before
  // that exchange, so a later tool call does not lose the instruction that
  // led to it. Never synthesize context.
  const selectedPairs = [];
  const selectedRequestIndices = new Set([latestRequestIndex]);
  for (let i = pairs.length - 1; i >= 0; i--) {
    const pair = pairs[i];
    const proposedRequestIndices = new Set(selectedRequestIndices);
    proposedRequestIndices.add(pair.requestIndex);
    const proposedSize = (selectedPairs.length + 1) * 2 + proposedRequestIndices.size;
    if (proposedSize <= limit) {
      selectedPairs.push(pair);
      selectedRequestIndices.add(pair.requestIndex);
    }
  }

  const selectedIndices = new Set(selectedRequestIndices);
  for (const pair of selectedPairs) {
    selectedIndices.add(pair.callIndex);
    selectedIndices.add(pair.responseIndex);
  }

  // Preserve trailing plain model output when there is spare capacity. It is
  // optional, so it never crowds out the request or a complete tool exchange.
  const lastSelectedIndex = Math.max(...selectedIndices);
  for (let i = validHistory.length - 1; i > lastSelectedIndex; i--) {
    const msg = validHistory[i];
    const parts = msg.parts || [];
    const hasToolSemantics = parts.some((part) => part && (part.functionCall || part.functionResponse));
    if (hasToolSemantics || msg.role !== "model" || selectedIndices.size >= limit) break;
    selectedIndices.add(i);
  }

  return [...selectedIndices]
    .sort((a, b) => a - b)
    .map((index) => validHistory[index]);
}

/**
 * Validates that function calls and responses are properly paired.
 *
 * In addition to enforcing adjacency, this also enforces that:
 * - If a model turn contains N function calls for a given tool name,
 *   the very next user turn must contain N function responses for that
 *   same tool name.
 * - There are no extra function responses for tools that were not called.
 *
 * This mirrors the behaviour described in the Gemini tooling docs and the
 * forum discussion you referenced, and strips out any legacy turns where
 * the counts didn't match (e.g. old code that only returned a single
 * functionResponse for multiple functionCalls).
 */
function validateHistoryPairs(history) {
  const validated = [];

  for (let i = 0; i < history.length; i++) {
    const msg = history[i];
    const parts = msg.parts || [];

    const hasFunctionCall =
      msg.role === "model" && parts.some((p) => p.functionCall);
    const isFunctionResponse =
      msg.role === "user" && parts.some((p) => p.functionResponse);

    // If validated is empty and this is a model turn, skip it
    // (Conversations must start with a user turn)
    if (validated.length === 0 && msg.role === "model") {
      console.warn(
        `Skipping model turn at index ${i} - cannot start history with a model turn.`
      );
      continue;
    }

    // --- Model turn with one or more function calls ---
    if (hasFunctionCall) {
      // CRITICAL: A model turn with function calls can ONLY come after a user turn
      // (either a regular text turn or a function response turn).
      // If the last message in validated is a model turn, this would cause:
      // "function call turn comes immediately after a user turn or after a function response turn" error
      const lastValidated = validated.length > 0 ? validated[validated.length - 1] : null;
      if (lastValidated && lastValidated.role === "model") {
        console.warn(
          `Removing function call at index ${i} - cannot follow another model turn. ` +
          `Last validated turn was role: ${lastValidated.role}. ` +
          `This would cause: "function call turn comes immediately after a user turn or after a function response turn" error.`
        );
        continue;
      }

      const nextMsg = i < history.length - 1 ? history[i + 1] : null;
      if (!nextMsg) {
        console.warn(
          `Removing orphaned function call at index ${i} (no following message).`
        );
        continue;
      }

      const nextParts = nextMsg.parts || [];
      const responseParts =
        nextMsg.role === "user"
          ? nextParts.filter((p) => p.functionResponse)
          : [];

      if (responseParts.length === 0) {
        console.warn(
          `Removing orphaned function call at index ${i} (no function responses in next turn).`
        );
        continue;
      }

      // Count how many times each tool was called in this turn
      const callCounts = {};
      parts.forEach((p) => {
        if (p.functionCall && p.functionCall.name) {
          const name = p.functionCall.name;
          callCounts[name] = (callCounts[name] || 0) + 1;
        }
      });

      // Count how many function responses we have per tool name
      const responseCounts = {};
      responseParts.forEach((p) => {
        const fr = p.functionResponse;
        const name = fr && fr.name;
        if (name) {
          responseCounts[name] = (responseCounts[name] || 0) + 1;
        }
      });

      let mismatch = false;

      // Every called tool must have exactly as many responses
      Object.keys(callCounts).forEach((name) => {
        if (callCounts[name] !== (responseCounts[name] || 0)) {
          mismatch = true;
        }
      });

      // And there must not be responses for tools that were never called
      Object.keys(responseCounts).forEach((name) => {
        if (!callCounts[name]) {
          mismatch = true;
        }
      });

      if (mismatch) {
        console.warn(
          `Removing mismatched function call/response pair at index ${i}. ` +
          `Calls: ${JSON.stringify(callCounts)}, ` +
          `Responses: ${JSON.stringify(responseCounts)}`
        );
        // Drop this model turn, and if the next turn is its response, drop that too.
        if (nextMsg.role === "user" && responseParts.length > 0) {
          i++; // Skip the mismatched response as well
        }
        continue;
      }

      // Pair looks good: keep both the model functionCall turn and the user functionResponse turn
      validated.push(msg);
      validated.push(nextMsg);
      i++; // Skip the response since we already added it
      continue;
    }

    // --- User turn with function responses but no preceding call in validated history ---
    if (isFunctionResponse) {
      const prevMsg = validated.length > 0 ? validated[validated.length - 1] : null;
      const prevParts = prevMsg && prevMsg.parts ? prevMsg.parts : [];
      const prevHasCall =
        prevMsg &&
        prevMsg.role === "model" &&
        prevParts.some((p) => p.functionCall);

      if (!prevHasCall) {
        console.warn(
          `Removing orphaned function response at index ${i} (no preceding function call in validated history).`
        );
        continue;
      }
    }

    // Regular message (no function call/response semantics to enforce)
    validated.push(msg);
  }

  return validated;
}

function sanitizeHistory(history) {
  if (!history || history.length === 0) return history;

  // Use the validation function to clean up the history
  return validateHistoryPairs(history);
}

/**
 * Append a model functionCall turn and its functionResponse turn atomically,
 * validating that the pair is well-formed BEFORE either turn enters history.
 *
 * This makes it structurally impossible to push a mismatched
 * function-call/response pair into history (the condition the tier recovery
 * ladder exists to clean up after the fact). On any mismatch it throws and
 * leaves `history` untouched; the caller's existing catch paths handle it, and
 * the repair ladder remains as a net.
 *
 * Non-functionCall parts in the model turn (e.g. text/thought parts) are allowed
 * and ignored — only functionCall/functionResponse counts are enforced.
 *
 * @param {Array} history
 * @param {object} modelTurn - { role: "model", parts: [...] } with >=1 functionCall part
 * @param {object} userTurn  - { role: "user", parts: [...] } with the matching functionResponse parts
 * @returns {Array} the same history array (mutated)
 * @throws {Error} if the pair is malformed or per-name call/response counts differ
 */
function appendFunctionExchange(history, modelTurn, userTurn) {
  if (!Array.isArray(history)) {
    throw new Error("appendFunctionExchange: history must be an array.");
  }
  if (!modelTurn || modelTurn.role !== "model" || !Array.isArray(modelTurn.parts)) {
    throw new Error('appendFunctionExchange: modelTurn must be { role: "model", parts: [...] }.');
  }
  if (!userTurn || userTurn.role !== "user" || !Array.isArray(userTurn.parts)) {
    throw new Error('appendFunctionExchange: userTurn must be { role: "user", parts: [...] }.');
  }

  // Count functionCalls per tool name in the model turn.
  const callCounts = {};
  for (const p of modelTurn.parts) {
    if (p && p.functionCall && p.functionCall.name) {
      const name = p.functionCall.name;
      callCounts[name] = (callCounts[name] || 0) + 1;
    }
  }
  const totalCalls = Object.values(callCounts).reduce((a, b) => a + b, 0);
  if (totalCalls === 0) {
    throw new Error("appendFunctionExchange: modelTurn contains no functionCall parts.");
  }

  // Count functionResponses per tool name in the user turn.
  const responseCounts = {};
  for (const p of userTurn.parts) {
    const fr = p && p.functionResponse;
    if (fr && fr.name) {
      responseCounts[fr.name] = (responseCounts[fr.name] || 0) + 1;
    }
  }

  // Every called tool must have exactly as many responses...
  for (const name of Object.keys(callCounts)) {
    if (callCounts[name] !== (responseCounts[name] || 0)) {
      throw new Error(
        `appendFunctionExchange: function call/response mismatch for "${name}" ` +
        `(calls=${callCounts[name]}, responses=${responseCounts[name] || 0}).`
      );
    }
  }
  // ...and there must be no responses for tools that were not called.
  for (const name of Object.keys(responseCounts)) {
    if (!callCounts[name]) {
      throw new Error(
        `appendFunctionExchange: functionResponse for uncalled tool "${name}".`
      );
    }
  }

  history.push(modelTurn);
  history.push(userTurn);
  return history;
}

/**
 * Tier 2 Recovery: Remove ALL function call/response pairs from history
 * Keeps only regular text messages
 */
function removeAllFunctionPairs(history) {
  return history.filter(msg => {
    const parts = msg.parts || [];
    const hasFunctionCall = parts.some(p => p.functionCall);
    const hasFunctionResponse = parts.some(p => p.functionResponse);
    return !hasFunctionCall && !hasFunctionResponse;
  });
}

/**
 * Tier 3 Recovery: Create fresh start with minimal context
 * Returns new history with just the original user message
 */
function createFreshStartWithContext(originalUserMessage) {
  return [{
    role: "user",
    parts: [{ text: originalUserMessage }]
  }];
}

export {
  maintainHistoryWindow,
  validateHistoryPairs,
  sanitizeHistory,
  appendFunctionExchange,
  removeAllFunctionPairs,
  createFreshStartWithContext
};
