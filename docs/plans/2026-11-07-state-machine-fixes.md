# State machine fixes

**Created:** 2026-11-07. **Source:** [taskpane state machines](../taskpane-state-machines.md), which describes each finding referenced here by its number (1a, 2a, …).
**Status:** Tier 1 done and committed (unit-tested where possible; manual Word checks M1 to M5 not yet run). Tier 2 in progress. Builds on Part A of [the 2026-11-01 review plan](2026-11-01-review-fixes-and-follow-ups.md), committed as `110070a`.

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
- **M8:** click Revert on a large document and press refresh at once. Refresh is refused until the revert finishes.
- **M9:** get the "Revert anyway" prompt, start a new request, press "Revert anyway". Expect the refusal message, and the prompt is still there afterwards.

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
