# Taskpane state machines

**Written:** 2026-11-07, from a code reading of the working tree on top of `d410b60` (release marker v2.3.0.1), including the uncommitted Part A changes from [the 2026-11-01 review plan](plans/2026-11-01-review-fixes-and-follow-ups.md). Nothing was run in Word.

The taskpane has no explicit state machine. Its behaviour is spread across module variables, DOM flags, `localStorage` keys and the chat history array. This document names each of those implicit machines, lists the variables that make up its state, compares the states the code means to be in with every combination the variables allow, and records the transitions that reach a bad state.

Line numbers point at the code as of the date above. The fix plan is [2026-11-07-state-machine-fixes.md](plans/2026-11-07-state-machine-fixes.md).

## Terms

- **Valid state:** a combination the code is designed to be in.
- **Possible combinations:** the product of every variable's values, whether reachable or not.
- **Bad state:** a reachable combination that is invalid, a **dead end** (no way out except a reset), or a **race** (two things that should not overlap can).

## Summary

| # | Machine | Possible | Valid | Bad states found |
| --- | --- | --- | --- | --- |
| 1 | Request lock (chat, revert, refresh) | 144 | 5 | 3 |
| 2 | Conversation memory vs. chat transcript | 12 | 3 | 3 |
| 3 | Edit receipt and Revert button | 9 | 3 | 2 (1 by design) |
| 4 | Revert confirmation | n/a | 4 | 2 |
| 5 | Glance cards | 112 | ~9 | 3 |
| 6 | Settings form | 8 | 3 | 4 |
| 7 | Glance question drag-and-drop | 6 | 1 | 1 |

Machines with no problems found are listed at the end.

---

## 1. Request lock

One `AbortController` slot guards every operation that writes to Word from the chat pane.

**Variables**

| Variable | Where | Values |
| --- | --- | --- |
| `currentRequestController` | [taskpane.js:164](../src/taskpane/taskpane.js) | `null`, a chat request's controller, a revert's controller |
| `signal.aborted` of that controller | | true, false |
| `chat-input`, `send-button`, `think-button` `.disabled` | DOM | 8 combinations (meant to move together) |
| The request's bubble (`loadingMsg`) | [chat-ui.js:62](../src/taskpane/modules/chat/chat-ui.js) | dots with Cancel, status text, none |

**Possible:** 3 × 2 × 8 × 3 = 144. **Valid (5):**

1. **Idle:** no controller, controls enabled, no bubble.
2. **Busy, before tools:** chat controller, not aborted, controls disabled, dots with Cancel.
3. **Busy, tools running:** chat controller, controls disabled, status text.
4. **Cancelling:** chat controller, aborted, controls disabled (passes quickly).
5. **Reverting:** revert controller, controls disabled, no bubble.

**Bad states**

- **1a. No Cancel once a tool starts (dead end).** `tool_start` calls `updateSystemMessage(loadingMsg, …)` ([taskpane.js:1205](../src/taskpane/taskpane.js)). That re-renders the bubble from scratch (`renderSender` clears it, [chat-ui.js:21](../src/taskpane/modules/chat/chat-ui.js)), which deletes the typing dots and the Cancel button. For the rest of the request (up to `TOTAL_REQUEST_TIMEOUT_MS`, 5 minutes, and up to 14 model turns) the only way to stop it is the refresh button, which also clears the conversation.
- **1b. Refresh unlocks while Word is still writing (race).** `refreshChat` ([taskpane.js:435](../src/taskpane/taskpane.js)) aborts the controller, sets the slot to `null` and enables the controls. Abort can't stop a `context.sync()` that is already committing ([word-operation-runner.js:125-133](../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js), [checkpoint-restoration.js:44](../src/taskpane/modules/chat/checkpoint-restoration.js)). The pane then looks idle while a write is in flight, a state none of the variables above can express.
  - **During a revert:** the revert may still land, and its result message is suppressed because the slot no longer matches ([taskpane.js:997](../src/taskpane/taskpane.js)). A new chat request can start while the revert writes, the overlap that the guard at [taskpane.js:975](../src/taskpane/taskpane.js) exists to prevent.
  - **During an edit:** the edit lands with no receipt and no Revert button. Its checkpoint is in IndexedDB but nothing in the UI points to it.
- **1c. Cancel during the document read is reported as an error.** The source reader throws `AbortError` with "Request cancelled before reading document context." ([word-document-source.js:12](../src/taskpane/modules/chat/word-document-source.js)). The catch only treats the exact text "Request cancelled by user" as a cancel ([taskpane.js:1241](../src/taskpane/taskpane.js)), so the user sees a red "Sorry, I couldn't get a response. Error: Request cancelled…".

## 2. Conversation memory vs. chat transcript

What the model remembers (`chatHistory`) and what the user sees (`#chat-messages`) are kept separately and can disagree.

**Variables**

| Variable | Values |
| --- | --- |
| Last entry of `chatHistory` | empty, model reply, user text, user tool results |
| `#chat-messages` | welcome screen, empty, conversation |

**Possible:** 4 × 3 = 12. **Valid (3):** empty history with the welcome screen or an empty chat; history ending in a model reply with a conversation shown. (History ending in a user turn is valid only while a request runs.)

**Bad states**

- **2a. A cancelled instruction stays in memory.** `stop()` pops the user's prompt only for `timeout` ([document-agent-loop.js:171](../src/taskpane/modules/chat/document-agent-loop.js)). After a cancel before any tool ran, the prompt stays as the last history entry and is sent again with the next request. Cancel "delete section 3", then ask "what does section 2 say?", and the model receives both instructions.
- **2b. Memory wiped while the transcript stays.** `provider_error` and `refresh_failed` set `state.history = []` ([document-agent-loop.js:208](../src/taskpane/modules/chat/document-agent-loop.js), [:374](../src/taskpane/modules/chat/document-agent-loop.js)) so tool calls are never replayed. The chat still shows the whole conversation and nothing says the model has forgotten it.
- **2c. Error cleanup can remove another request's prompt.** On a failure with no tool records, the catch pops the last history entry if it is a user turn ([taskpane.js:1250](../src/taskpane/taskpane.js)). If the failure happened before this request pushed its prompt (for example while reading the document), the entry belongs to an earlier, cancelled request.

## 3. Edit receipt and Revert button

Each request has one bubble. It shows progress, then the edit receipt, and carries the Revert button.

**Variables:** bubble text (progress, receipt, stop message) × Revert button (none, latest checkpoint, earlier checkpoint). **Possible:** 9. **Valid end states (3):** receipt with the latest checkpoint's Revert; stop message with or without Revert; a refusal receipt with no Revert.

**Bad states**

- **3a. Progress text replaces the receipt.** If the model inspects or searches after an edit, `tool_start` sets the text to "Inspecting the document…" or "Searching the web…". Read-only tools end with `display.show = false` ([tool-registry.js:93](../src/taskpane/modules/chat/tool-registry.js)), so `tool_end` never puts the receipt back, and `response` keeps the bubble because an edit ran ([taskpane.js:1214](../src/taskpane/taskpane.js)). The final state is a progress message next to a Revert button.
- **3b. Only the last edit batch has a Revert button.** A second verified edit replaces the first Revert button ([chat-ui.js:151-153](../src/taskpane/modules/chat/chat-ui.js)). This is the intended design (review plan A1: "the message's Revert button belongs to the last edit"), listed here only for completeness.

**Related bad exit: errors after a confirmed edit.** Some failures are thrown instead of ending through `stop()`: "contained no candidates" ([document-agent-loop.js:215](../src/taskpane/modules/chat/document-agent-loop.js)), "missing content.parts" for a `MAX_TOKENS`, `SAFETY` or `OTHER` ending ([:28](../src/taskpane/modules/chat/document-agent-loop.js)), a failed tools-module load ([:255](../src/taskpane/modules/chat/document-agent-loop.js)), or a non-cancel executor exception ([:315](../src/taskpane/modules/chat/document-agent-loop.js)). If an edit was already written, the user sees "Sorry, I couldn't get a response" with no "inspect the document" warning, and for "missing content.parts" a **Retry** button that pastes the original request back ([taskpane.js:1256-1271](../src/taskpane/taskpane.js)). Sending it again can apply the edit twice. A failure from `requestModel` itself is handled correctly through `provider_error`; these paths are not.

## 4. Revert confirmation

`restoreCheckpoint` ([taskpane.js:973](../src/taskpane/taskpane.js)) reads the checkpoint, compares tokens, and either writes, or asks with "Revert anyway" / "Keep document".

**Valid states (4):** idle, reverting, waiting for confirmation (any number of open prompts), done (verified or not verified).

The token check is sound: a stale "Revert anyway" asks again, and several open prompts can't discard more than the user was shown.

**Bad states**

- **4a. A refused confirmation is used up.** `addConfirmButtons` removes the buttons before calling the handler ([chat-ui.js:242](../src/taskpane/modules/chat/chat-ui.js)). If a chat request is running, the handler refuses ("Wait for the current request…") and the prompt is gone. The user has to start again from the original Revert button.
- **4b. "Revert anyway" saves nothing first.** The changes it discards (later edits and the user's typing) get no checkpoint. Whether Word's own Undo can bring them back has not been checked.

## 5. Glance cards

**Variables (per card, plus two shared)**

| Variable | Values |
| --- | --- |
| Indicator class | gray, green, yellow, red |
| Summary text | "Waiting for analysis…", "Press refresh…", "Checking…", a result, "No answer…", "Error running check.", "too long" |
| `glanceController` | null, a run |
| `glanceAutoRun` setting | on, off |

**Possible:** 4 × 7 × 2 × 2 = 112. **Valid (about 9):** idle (gray, idle text matching the setting, no run), checking (gray, "Checking…", a run), each result colour with its summary, and the three gray failure texts with no run.

**Bad states**

- **5a. Returning from Settings wipes results.** `showMainView` always calls `renderGlanceMain` ([taskpane.js:431](../src/taskpane/taskpane.js)), which resets every card to the idle text ([:684](../src/taskpane/taskpane.js)). Finished results disappear. With auto-run on, cards say "Waiting for analysis…" although no run is coming (auto-run fires only at startup and after Save). A card added in Settings shows the same.
- **5b. A running check is shown as idle.** Going to Settings and back during a check resets the cards to the idle text while the check is still running; its results then appear later without a "Checking…" phase.
- **5c. No API key: refresh does nothing (dead end).** `runGlanceChecks` returns silently without a key ([taskpane.js:834](../src/taskpane/taskpane.js)) while the cards say "Press refresh to run this check."

## 6. Settings form

**Variables:** two groups of fields. Group A saves on change: Glance questions, Glance auto-run, redlines, redline author. Group B saves only on Save: API key, Fast and Slow models, system prompt. Each group is clean or edited, and the pane shows the main view or Settings. **Possible:** 2 × 2 × 2 = 8. **Valid (3):** both clean in either view; edits only while Settings is open.

**Bad states**

- **6a. Back half-saves.** Back keeps group A and silently drops group B. Already tracked as C5 in the review plan.
- **6b. Save errors are invisible.** "API Key cannot be empty." and the custom model error are added to the chat ([taskpane.js:480](../src/taskpane/taskpane.js), [:499](../src/taskpane/taskpane.js)), which is hidden behind Settings. Save appears to do nothing.
- **6c. Settings change under a running request (race).** The redline author saves on every keystroke ([taskpane.js:276](../src/taskpane/taskpane.js)), and each edit reads the author and the redline toggle when it writes ([document-tools.js:29](../src/taskpane/modules/chat/document-tools.js)). An edit that lands while the user types gets a half-typed author; toggling redlines mid-request changes mode between batches.
- **6d. The key can't be removed.** Save refuses an empty key, so there is no way back to the no-key state from the UI.

## 7. Glance question drag-and-drop

**Variables:** `dragSrcEl` (null, a card on screen, a card already replaced by a re-render) × what is dropped (a card's handle, anything else). **Possible:** 6. **Valid:** 1 (a card on screen, dropped from its handle).

**Bad state**

- **7a. Dropping anything else moves the last question.** `dragSrcEl` starts as `null` and is never cleared ([taskpane.js:767](../src/taskpane/taskpane.js), [:819](../src/taskpane/taskpane.js)). Dropping text (for example dragged out of another question's box) onto a card gives `indexOf(dragSrcEl) === -1`, and `splice(-1, 1)` takes the last question and moves it to the drop position ([taskpane.js:806-811](../src/taskpane/taskpane.js)). Confirmed in Node: `['A','B','C']` becomes `['C','A','B']`.

## No problems found

- **Main view / Settings switch** ([taskpane.js:364](../src/taskpane/taskpane.js)): the 200 ms fade uses an uncancelled timer, but the button that would start the opposite switch is hidden until the timer fires, so two switches can't overlap.
- **Lazy module loader** ([lazy-module-loader.js](../src/taskpane/modules/utils/lazy-module-loader.js)): shares a pending load and allows a retry after a failure.
- **Latest Glance run wins** ([taskpane.js:839-842](../src/taskpane/taskpane.js)) and **superseded chat requests** (every UI write checks `currentRequestController === requestController`).
- **No-progress loop guard** ([mutation-loop-guard.js](../src/taskpane/modules/chat/mutation-loop-guard.js)).
- **Checkpoint migration flag** ([taskpane.js:927](../src/taskpane/taskpane.js)): set before its awaits, so a concurrent save skips ahead, but pruning only removes other sessions' checkpoints and the new checkpoint has the highest id, so nothing is lost.
