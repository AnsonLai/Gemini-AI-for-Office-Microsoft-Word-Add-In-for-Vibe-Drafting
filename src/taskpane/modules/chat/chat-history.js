/**
 * Maintains a rolling window of chat history while preserving function call/response pairs
 */
function maintainHistoryWindow(history, maxMessages) {
  if (history.length <= maxMessages) {
    return history;
  }

  // Start from the end and work backwards, keeping complete pairs
  let newHistory = [];
  let i = history.length - 1;

  while (i >= 0 && newHistory.length < maxMessages) {
    const msg = history[i];

    // If this is a function response, we must include its preceding function call
    const isFunctionResponse = msg.role === "user" && msg.parts && msg.parts.some(p => p.functionResponse);

    if (isFunctionResponse && i > 0) {
      const prevMsg = history[i - 1];
      const hasFunctionCall = prevMsg.role === "model" && prevMsg.parts && prevMsg.parts.some(p => p.functionCall);

      if (hasFunctionCall) {
        // Add both the function call and response together
        newHistory.unshift(msg);
        newHistory.unshift(prevMsg);
        i -= 2;
        continue;
      }
    }

    // If this is a function call, check if its response is already included
    const hasFunctionCall = msg.role === "model" && msg.parts && msg.parts.some(p => p.functionCall);

    if (hasFunctionCall && i < history.length - 1) {
      const nextMsg = history[i + 1];
      const hasResponse = nextMsg.role === "user" && nextMsg.parts && nextMsg.parts.some(p => p.functionResponse);

      if (hasResponse && !newHistory.includes(nextMsg)) {
        // Skip this function call since its response isn't in our window
        i--;
        continue;
      }
    }

    newHistory.unshift(msg);
    i--;
  }

  // Final validation: remove any orphaned function calls or responses at the boundaries
  return validateHistoryPairs(newHistory);
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
  removeAllFunctionPairs,
  createFreshStartWithContext
};
