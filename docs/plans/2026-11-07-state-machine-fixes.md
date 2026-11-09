# State machine fixes

**Created:** 2026-11-07. **Source:** [taskpane state machines](../taskpane-state-machines.md), which describes each finding referenced here by its number (1a, 2a, …).
**Status:** Tiers 1 and 2 done (unit-tested where possible; manual Word checks M1 to M9 not yet run). Tier 1 is committed as `3408704`; Tier 2 is committed as `788b442`. Tier 3 waits on the owner. Builds on Part A of [the 2026-11-01 review plan](2026-11-01-review-fixes-and-follow-ups.md), committed as `110070a`.

## How each finding is triaged

A finding is fixed here only when all of these hold:

- **Can fix:** the fix is in add-in code, not `@ansonlai/docx-redline-js` (library defects are reported, not patched; see [KNOWN_BUGS.md](KNOWN_BUGS.md)).
- **Should fix:** a user can reach the bad state through normal use, and it either loses or misreports work, or leaves them stuck.
- **No open design question:** the right behaviour is obvious from the code's own intent. When it isn't, the finding goes to Tier 3 for the owner.

Each fix follows the same steps:

1. **Confirm.** Where the code has unit tests (agent loop, chat UI, stores), write a test that fails first. `taskpane.js` has no unit tests, so there the confirmation is the code path in the description plus a manual check in Word.
2. **Fix** with the smallest change that removes the bad state.
3. **Verify:** `npm test` and `npm run build` pass, and the manual check is added to the list at the end.

## Tier 1: low-hanging fruit

Small, local, no design question. Each is a few lines.

| # | Finding | Fix | Verify |
| --- | --- | --- | --- |
| T1 | 7a: dropping text on a Glance question moves the last question | Clear `dragSrcEl` on drag end; ignore a drop whose source isn't a card in the list | Manual (M1) |
| T2 | 1c: Cancel during the document read shows as an error | Treat any failure after this request's signal was aborted as a cancel | Manual (M2) |
| T3 | 2c: error cleanup can remove an earlier request's prompt | Pop the last history entry only if this request pushed it | Code reading (one flag) |
| T4 | 2a: a cancelled instruction stays in memory | When a request is cancelled before any tool ran, drop its prompt, as a timeout already does | Loop unit test |
| T5 | 6c: author or redline setting changes under a running request | Read both once when the request starts | Code reading |
| T6 | 6b: Save errors are hidden behind Settings | Show them in a status line inside the Settings view | Manual (M3) |
| T7 | 5c: Glance refresh does nothing without a key | Say "Add your Gemini API key in Settings to run checks." on the cards | Manual (M4) |
| T8 | 5a, 5b: returning from Settings wipes Glance results | Rebuild the cards only when the questions changed; one idle text, "Press refresh to run this check." | Manual (M4) |
| T9 | 3a: a later inspect or search replaces the edit receipt | Remember the last edit receipt and put it back when a read-only step ends | Manual (M5) |

## Tier 2: worth fixing, more work

Each needs more than a few lines or touches timing, so each gets a test where possible and a Word check.

| # | Finding | Proposed fix | Verify |
| --- | --- | --- | --- |
| T10 | 1a: no Cancel once a tool starts | Keep the Cancel button when the bubble's text changes; remove it when the request ends | chat-ui test with jsdom; manual (M6) |
| T11 | Errors after a confirmed edit offer Retry and no warning | In the loop, end those paths through `stop('provider_error', …)` when tool records exist, as a failed `requestModel` already does; never offer Retry after a write | Loop unit tests; manual (M7) |
| T12 | 1b: refresh unlocks while Word is writing | Refuse refresh during a revert (as Revert is refused during a request). During a chat request, make refresh cancel and clear only after the request has settled | Manual (M8) |
| T13 | 4a: a refused "Revert anyway" is used up | Keep the prompt when the revert is refused; remove it only once the revert starts | chat-ui or manual (M9) |

## Tier 3: owner decisions

Not fixed here. Each needs a choice about behaviour.

- **2b:** after `provider_error` or `refresh_failed`, the model's memory is wiped but the chat still shows the conversation. Options: say so in the stop message ("I've cleared this conversation from my memory…"), or clear the transcript too, or keep history minus the tool calls.
- **4b:** "Revert anyway" discards later changes without a checkpoint. First check whether Word's Undo restores them; if not, save a checkpoint of the current body before the write.
- **6a, 6d:** settings save model and removing the key. Already tracked as C5 in the 2026-11-01 plan; fold these in there.

## Not fixing

- **3b:** only the last edit batch in a request has a Revert button. Intended (review plan A1).

## Manual checks in Word

Add to Part B of the 2026-11-01 plan once the matching fix is in.

- **M1:** in Settings, select text in one Glance question and drag it onto another card. The question order doesn't change.
- **M2:** send a request on a long document and press Cancel at once. Expect the grey "Request cancelled." message, not a red error.
- **M3:** in Settings, clear the API key field and press Save; then choose a custom model with an invalid id and press Save. Each shows an error in Settings.
- **M4:** without a key, press Glance refresh: the cards say to add a key. With a key, run Glance, open Settings, press Back: the results are still there. Add a question, press Back: the cards are rebuilt with "Press refresh to run this check."
- **M5:** ask for an edit and then a check of the result in one request (for example "bold the title, then confirm it's bold"). The bubble ends with the edit receipt and its Revert button.
- **M6:** start a request that inspects and edits; press Cancel while "Inspecting the document..." shows. The request stops.
- **M7:** hard to trigger on demand; covered by unit tests.
- **M8:** click Revert on a large document and press refresh at once. Refresh is refused until the revert finishes. Then start a request that edits, and press refresh while "Validating the document edit batch..." shows: the refresh button greys out, the request stops, and only then does the chat clear.
- **M9:** get the "Revert anyway" prompt, start a new request, press "Revert anyway". Expect the refusal message, and the prompt is still there afterwards.
- **M10:** hard to trigger on demand (needs a reply with no content after an inspect, for example in a dev build). After a failure that follows only reading, the bubble goes, a red error appears, and where Retry is offered it resends the question with no duplicate. After a failure that follows an applied edit, the bubble keeps its receipt and Revert, there is no Retry, and the message ends "I no longer have the earlier conversation in memory…".
- **M11:** needs a request that does not settle (a dev build, or a low `REFRESH_SETTLE_TIMEOUT_MS`). Press refresh during an edit. After 15 seconds the refresh button comes back with "Word is still finishing the last request…", the chat is not cleared and the input stays locked. Refresh again once the request settles: the chat clears.
- **M12:** send a request that edits and press refresh as the write lands. The fresh chat shows "Chat history cleared…" and then "An edit from the cancelled request was applied to the document. Use Revert to undo it." with a Revert button that restores the document. A refresh with no landed edit shows only the first message.
- **M13:** Settings has no Save button. Change the Fast model, Slow model, system prompt and API key, tabbing out of each: each shows "Saved." and survives Back and reopening. A custom model id of `../x` shows the error and is not saved. Clearing the key shows "API key removed."; after Back, Glance says to add a key and sending a message asks for one. Re-enter the key and the pane works again.
- **M14:** with Glance auto-run on, change a question or a model and press Back: Glance runs once. Open Settings and press Back with no changes: results are kept and nothing re-runs.

## Progress

- **2026-11-07:** plan written. Tier 1 (T1 to T9) done in the working tree:
  - **T4** in [`document-agent-loop.js`](../../src/taskpane/modules/chat/document-agent-loop.js): `stop()` drops the unanswered prompt for `cancelled` as well as `timeout` when no tool ran. A new case in `tests/document_agent_loop_tests.mjs` failed before the change and passes after.
  - **T9** in `taskpane.js` `onEvent`: a request-local `shownToolResult` puts the last shown tool result back when a read-only step ends. `tests/agentic_navigation_dispatch_tests.mjs` now carries that variable across events and checks that an inspection after an edit restores the receipt.
  - **T1, T2, T3, T5, T6, T7, T8** in [`taskpane.js`](../../src/taskpane/taskpane.js), plus a `#settings-status` line in `taskpane.html` and `.settings-status` in `taskpane.css`:
    - **T2:** a failure after Cancel is reported as a cancel.
    - **T3:** history cleanup pops only this request's `promptTurn`, by identity.
    - **T5:** the author and the redline setting are read once per request.
    - **T6:** Save errors appear in Settings and clear when Settings opens.
    - **T7:** Glance refresh without a key says so.
    - **T8:** `renderGlanceMain` rebuilds only when the questions changed, aborts a check asking the old questions, and always uses "Press refresh to run this check." as the idle text.
    - **T1:** the drop handler needs both cards in the list, and drag end clears `dragSrcEl`.
  - **Also changed:** a cancel after a tool ran keeps the bubble (its receipt and Revert button) and says "Request cancelled. Inspect the document before retrying."
  - **Verified:** `npm test` 51 suites pass, 0 fail; `npm run build` succeeds with only the two known bundle-size warnings (A8). Manual checks M1 to M5 not run.
  - Committed as `3408704`, on top of Part A (`110070a`).
- **2026-11-07:** Tier 2 (T10 to T13) done:
  - **T10** in [`chat-ui.js`](../../src/taskpane/modules/chat/chat-ui.js): `updateSystemMessage` keeps a request's Cancel button when the status text changes. New `removeCancelButton` runs in `sendChatMessage`'s `finally`. A `.chat-message.system > .cancel-request-btn` rule in `taskpane.css` shows it inline after the text.
  - **T11** in [`document-agent-loop.js`](../../src/taskpane/modules/chat/document-agent-loop.js): a new `catch` ends any failure after a tool ran (no candidates, a `MAX_TOKENS` or `SAFETY` ending with no content, a tools-module load failure, an executor exception) through `stop('provider_error', …)`, as a failed `requestModel` already did. History is cleared and the message says to inspect the document and includes the error. Before any tool ran, errors still reach the taskpane, which may offer Retry. So Retry now never follows a tool run.
  - **T12** in `taskpane.js`:
    - **During a revert:** refresh is refused ("Wait for the revert to finish…"), using a new `revertInProgress` flag.
    - **During a chat request:** refresh cancels it, disables the refresh button, and waits on `currentRequestSettled` (resolved at the end of `sendChatMessage`'s `finally`) before clearing. The request unlocks the UI itself, so nothing can start while a Word write may still be committing.
    - **Trade-off:** if a `Word.run` never returns, refresh now waits with it instead of unlocking the pane.
  - **T13:** `addConfirmButtons` keeps the question open when its handler returns `false`. The "Revert anyway" handler returns `false` when a request is running.
  - **Tests:**
    - New `tests/chat_ui_controls_tests.mjs` (Cancel kept through status updates and removed at the end; Revert kept; refused confirmation stays open; "Keep document" closes). It fails on the pre-fix `chat-ui.js`.
    - Three new cases in `tests/document_agent_loop_tests.mjs` (empty and `MAX_TOKENS` replies after an edit end as `provider_error`; before any tool, the error is still thrown). They failed before the fix.
  - **Verified:** `npm test` 52 suites pass, 0 fail; build succeeds with only the two known warnings. Manual checks M6, M8 and M9 not run.
- **2026-11-09: review of Tier 2.**
  - **Hardened in `taskpane.js`:** `settleRequest()` now runs in its own `finally`, so a throw in the request's cleanup cannot leave refresh waiting forever. `revertInProgress` is now cleared unconditionally when a revert ends.
  - **Open, owner to decide:**
    - **(a) T11 fires on read-only tool records too.** The catch runs when any tool record exists, including a read-only inspect or search, so a failure after only reading also loses Retry and clears memory. Gating it on a write attempted or confirmed would match "never offer Retry after a write".
    - **(b) Contradictory stop message.** In that case the stop message can say "The completed tool results remain available in this conversation" while history is cleared on the same line. The older `provider_error` path does the same.
    - **(c) T12 can wait forever.** If a `Word.run` never returns, refresh waits with it. A timeout on the wait (for example 15 seconds, then a message) would restore a way out.
    - **(d) T12 loses the receipt.** When the cancelled request's edit lands, refresh still clears the chat, so the receipt and Revert button are lost. Either skip the clear or leave a note in the fresh chat.
    - **(e) Untracked "Also changed" behaviour.** The cancel-after-a-tool behaviour listed under "Also changed" has no row in the tables and no manual check.
- **2026-11-09: owner decisions applied** (not yet run in Word; checks M10 to M14 added):
  - **(a) and Tier 3 2b:** only a write attempted or confirmed ends a failure as `provider_error`. After only reads or refused edits, the error is thrown with the history as it was when the request started, so the taskpane removes the bubble, drops the prompt and may offer Retry. This covers both a failed `requestModel` and an unusable reply. Every stop that clears memory now ends "I no longer have the earlier conversation in memory, so restate anything I need."
  - **(b):** `completedWork` no longer says tool results remain available when the history is cleared.
  - **(c):** refresh waits at most 15 seconds (`TIMEOUT_LIMITS.REFRESH_SETTLE_TIMEOUT_MS`), then says Word is still finishing and leaves the chat and the lock as they are.
  - **(d):** the request settles with `{ wrote, checkpointIndex }`. After a refresh that cancelled a request whose edit landed, the fresh chat carries a note with that edit's Revert button.
  - **6a, 6d (C5):** every Settings field saves on change. An empty API key removes the stored key. The Save button is gone; Back returns to the chat and runs Glance once if a relevant setting changed and auto-run is on.
  - **4b:** left as is. The owner expects to replace Revert with Word's Undo now that edits are applied step by step.
  - **(e):** still open: the cancel-after-a-tool behaviour has no manual check.
  - **New code:** `modules/chat/request-outcome.js` and `modules/config/settings-save.js`, with `tests/request_outcome_tests.mjs` and `tests/settings_save_tests.mjs`.
