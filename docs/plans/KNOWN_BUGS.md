# Known bugs and host-boundary exceptions

**Updated:** 2026-10-18, on `@ansonlai/docx-redline-js` 0.11.0 (exact pin). The browser demo and MCP server were removed the same day.
Previous: 2026-10-12, after the 0.9.0 upgrade ([validation](../validation-reports/2026-10-12-docx-redline-v090.md)).

This file tracks two things:

- Defects and gaps we know about but have not fixed.
- Every place production code uses the Word JavaScript API (Office.js) instead of the library.

The goal is to use the library wherever possible. Each Word API use must say why it exists and what would let us remove it.

**Library defects** (`@ansonlai/docx-redline-js`, which the owner maintains) get a short report here or in the library repo, with a portable reproducer that uses only public APIs and exits 0 while the defect reproduces and 1 once fixed. Never work around a library defect in the add-in. Resolved reports from the former `docs/library-issues/` folder are summarized in the [closeout](../validation-reports/2026-10-17-library-issues-closeout.md).

## Open bugs and gaps

### KB-1: Body-final table append via live Word insertion leaves an extra paragraph on Reject All

- **Where:** the production Word insertion of a library-prepared table append at the end of the body, in [`word-operation-runner.js`](../../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js). The library's own output is correct.
- **Evidence (2026-10-12, Word 16 desktop):**
  - Exporter: [`export-table-append-reject-all.mjs`](../../scripts/word-desktop/export-table-append-reject-all.mjs) (`--export-dir <dir>` writes the fixtures and `manifest.json`).
  - Oracle command: `verify-wp6-word.ps1 -CaseName public-library-table-append-reject-all`.
  - The source, tracked, Accept All and Reject All checks pass on the 0.9.0 package: 7 paragraphs and no revisions after `Revisions.RejectAll()`.
  - The oracle's native-insert lane fails Reject All with 8 paragraphs against 7, plus a trailing empty paragraph. That lane inserts the production adapter's payload into a live document with COM `Range.InsertXML`.
  - Word appears to add an untracked empty paragraph after an inserted body-final table.
- **User impact today:** none. The agent refuses structural table edits and table appends. This is a latent blocker for bringing table append back.
- **Exit condition:** an Office.js run of a body-final table append through the production adapter. Its saved export must give 7 paragraphs and 0 revisions under `Revisions.RejectAll()`, checked by the independent oracle. Any fix must stay in the transport (for example, how the replace range ends). No OOXML rewriting of library output.
- **Not yet established:** whether Office.js `body.insertOoxml` behaves like COM `InsertXML` here. COM evidence does not certify Office.js.

### KB-2: Whole-document `applyRedlineToOxml` corrupts unrelated paragraphs (unsupported usage)

- `applyRedlineToOxml` is for paragraph, range or table scope. A complete `document.xml` goes through `applyOperationsToDocumentXml`, as production does.
- Given a whole body, the library parses the modified text as Markdown. Unchanged literal text such as `7. REMEDIES`, `A. …`, `____` and `**x**` is reinterpreted, so unrelated paragraphs change in Accept All. This happens in 0.8.4 and 0.9.0.
- In this path, 0.9.0 also drops an inserted `2`, producing "two () copies".
- Production never uses this shape.
- `tests/addin/integration_tests.mjs` used it and passed only through OOXML substring assertions. Those cases now use targeted `applyOperationsToDocumentXml` and assert exact per-paragraph Accept All and Reject All text.
- **Follow-up (upstream, optional):** have `applyRedlineToOxml` refuse a multi-paragraph body input rather than return `ok`.
- **Exit condition:** no test or production caller uses the whole-body shape, which is met as of 2026-10-12. An upstream refusal would make this impossible to repeat.

### KB-3: `Document.RejectAllRevisions()` differs from the ribbon

- After a body-final table append, Word COM `Document.RejectAllRevisions()` and `AcceptAllRevisions()` leave the tracked last paragraph mark behind. `Document.Revisions.RejectAll()` and `AcceptAll()`, which the ribbon uses, resolve it fully.
- This is a Word API quirk, also noted in the library's 0.9.0 changelog.
- [`verify-wp6-word.ps1`](../../scripts/verify-wp6-word.ps1) now uses `Revisions.AcceptAll/RejectAll`.
- Oracle reports produced before 2026-10-12 used the `Document.*` methods.

### KB-4: List conversion and list-format changes (resolved for the exposed scope, 2026-10-14)

- 0.9.0 ships `list-format` (`convert-plain`, `change-format`, `join`) and `paragraph-format.numbering`. It also supports a composed `redline` plus `list-format` for changed-text headers, and returns explicit refusal codes.
- **2026-10-14:** `apply_document_edits` now exposes `convert_to_list`, `change_list_format` and `continue_list` as one-to-one maps onto `list-format` (contract in [`agentic-tool-contracts.md`](../agentic-tool-contracts.md)). The add-in parses no markers and writes no numbering. Offline evidence: [`document_list_intents_tests.mjs`](../../tests/document_list_intents_tests.mjs) (Accept All and Reject All on synthetic and Word-authored fixtures, library refusals with no write) and the direct golden step 12 replay.
- **Host evidence (2026-10-14):** five new Office.js capability cases plus one refusal control all pass. The Word oracle gives 88 PASS / 0 FAIL, with labels read from `ListFormat.ListString`. The real-Word NDA replay passes 16/16 (51/51 checks) with step 12 applied. See the [2026-10-14 list-format validation](../validation-reports/2026-10-14-list-format-intents.md).
- **Still open:** the live-provider lane and Word Online (release track). Browser/MCP facade parity was dropped with those consumers (2026-10-17). Not exposed: range-scoped changes, `paragraph-format.numbering`, depth/indent changes, table lists, `detach`.
- Tracked in the [canonical list migration follow-up](completed/2026-09-30-canonical-list-migration-follow-up.md) (WP1 to WP3).
- **Closed (2026-10-17):** the [list follow-up](completed/2026-09-30-canonical-list-migration-follow-up.md) is complete and archived.

### KB-5: `TRACKING_POLICY_MISMATCH` guard (retired by phase-2 WP3, 2026-10-17)

- The add-in refused an untracked batch whose prepared output carried new own-author revisions, a workaround for the 0.8.4 surgical-insertion bug.
- 0.9.0 fixed that bug and enforces the same rule itself (`REVISION_FIDELITY_VIOLATION` / `UNTRACKED_CREATED_REVISION`). The guard never fired on the certified cases during the 0.9.x cycle.
- **Retired:** the guard is removed from `document-tools.js`; the library's refusal is relied on. The untracked capability cases in `document_tools_tests.mjs` still pass.

### KB-6: Revision-fidelity refusals reach the model unmapped (closed 2026-10-30, no mapping needed)

- 0.9.0 refuses edits that would break Reject All (`REVISION_FIDELITY_VIOLATION`, default `enforce`). We never pass `verifyRevisionFidelity`, which is correct.
- [`document-tools.js`](../../src/taskpane/modules/chat/document-tools.js) passes the library's `error.recovery` through verbatim. No retained case hits this in the current gates, but the model gets the library's recovery text, not a product-specific instruction.
- **Exit condition:** a deterministic test showing the agent's response to a `REVISION_FIDELITY_VIOLATION`. Map it only if that test shows the model mishandles the library guidance.
- **Closed (2026-10-30):** `document_agent_loop_tests` now feeds the loop a receipt with the library's real refusal shape (`REVISION_FIDELITY_VIOLATION`, `recovery.action: narrow_or_change_operation`, `sameArgumentsSafe: false`, a hint). The model receives the code, the recovery action and the unsafe-retry flag. Nothing is written. A second attempt with the same arguments stops with `no_progress`, and a narrowed retry is allowed and succeeds. The library guidance is enough, so no product-specific mapping was added.

### KB-7: Library Markdown pre-processor collapses literal underscore runs (closed 2026-10-17, fixed in 0.10.0)

- **Where:** `@ansonlai/docx-redline-js` 0.8.4 and 0.9.0, `pipeline/markdown-processor.js` (`preprocessMarkdown`, the `__x__` and `_x_` patterns). Report: `2026-10-12-table-reconciliation-underscore-collapse.md` (in git history, see the [closeout](../validation-reports/2026-10-17-library-issues-closeout.md)).
- **Evidence (2026-10-12):**
  - `tests/addin/integration_tests.mjs`, case "Add Date Row", prints `KNOWN_LIBRARY_DEFECT:` under its PASS line. A `table-reconciliation` on the sample NDA signature table shrinks `_________________________` to `_` in Accept All, including cells the edit did not change.
  - The case keeps its strict assertions. The marker is printed only when the sole mismatch is this underscore collapse. Any other mismatch still fails.
  - Reproducer: `2026-10-12-table-reconciliation-underscore-collapse-reproducer.mjs` (in git history; exit 0 while it reproduces; 0.8.4 behaves the same).
- **User impact today:** the table-reconciliation route is not reachable. `rewrite_paragraph` and `replace_text` forbid line breaks, and the route needs a multi-line pipe table. Insertions refuse targets and text that `preprocessMarkdown` would change (`plainParagraph`, `checkedInsertionTarget`). `replace_text` is safe because the library refuses a localized replacement it cannot round-trip (`PATCH_ROUNDTRIP_MISMATCH`). **However, `rewrite_paragraph` is exposed through the same root cause.** It sends a `redline` with `structuredContent: false`, and the library still runs `preprocessMarkdown` on it. Rewriting a paragraph whose text holds an underscore run (signature or date line, `__x__`) returns `status: 'ok'` with the run shrunk. Example: `_________________________` with ` (sign)` appended accepts as `_____ (sign)`. `document-tools.js` verifies Word against the prepared text, not the requested text, so nothing catches it. Signature-block rewrites are the likely trigger.
- **Exit condition:** a published library fix, with the reproducer exiting 1. Then remove the known-defect gate from "Add Date Row" (the `squash`/`knownDefect` logic and the marker). Consider a guard that compares the requested `rewrite_paragraph` text with the prepared text before the interim period ends.
- **Closed (2026-10-17):** fixed in the published 0.10.0. Underscore emphasis follows CommonMark flanking rules, and a caller's `structuredContent: false` keeps inline Markdown literal. The reproducer exits 1, the "Add Date Row" known-defect gate is removed, and the report was moved to `library-issues/completed/` (folder since removed). The Office.js, Word oracle and NDA replay lanes pass on 0.10.0 ([validation](../validation-reports/2026-10-17-docx-redline-v0100.md)).
- **Related observation (not filed):** a `list-change` whose kept item carries `-` markers on a decimal list converts the list to a new bullet list (new `numId`, bullet labels). The only diagnostic is "Paragraph expanded to list fragment". The 0.9.0 README ("List operations") and `AGENT_KNOWLEDGE_BASE.md` say retyped markers keep the old numbering and warn `LIST_FORMAT_NOT_APPLIED`. Checked on the sample's decimal list with a single-paragraph target. This looks like a gap against the documented contract, probably on the single-paragraph or fragment path. It is not filed because Reject All and the multi-paragraph target form were not checked. The add-in avoids it by deriving markers from the target list's kind (`listInsertionMarkers`).

### KB-10: Live Gemini Flash runs out of model turns while inspecting (first live run, 2026-10-28)

- **Evidence:** the first live NDA run (`--mode live --profile fast --model gemini-flash-latest`, `.cache/golden/live-fast`) finished 16/16 steps but failed 16 of 51 checks. In steps 5 (bold signature labels), 10 (unbold "British Columbia"), 12 (headers to list), 13 (highlight the date) and 14 (review comments), the model made one narrow `inspect_document` search per turn (for example `"1."`, `"2."`, `"3."`…) and used all six model turns without an edit. The user saw "Stopped after the maximum number of model turns without a final response." Step 1 also hit the limit after a verified edit.
- **Why replay didn't catch it:** the deterministic replay batches read-only calls into one turn; the live model does not. No earlier live baseline exists.
- **User impact:** about a third of ordinary requests do nothing with the default fast model. **Release blocker for v2.3.0.0.**
- **Options (owner decision):** let read-only inspection turns count against a separate, larger inspection cap instead of the six-turn budget (mutation budgets unchanged); have the editing guide ask for one broad inspection or parallel searches; or raise the turn limit. Re-run both live profiles afterwards.
- **Fix (2026-10-28, owner chose it):** a turn whose calls are all known read-only `inspect_document` calls uses an allowance of eight inspection turns before the six-turn budget, so a request is still bounded at 14 turns plus the 180-second timeout. Mutation and failure budgets are unchanged. Editing guide v6 asks for distinctive search words, several inspections in one response, and one paged query for many targets. Pinned by `document_agent_loop_tests`. **Still to confirm:** re-run both live profiles. Replay host lanes on the fix (2026-10-28): Office.js passed (20 cases), Word oracle 88/0, NDA replay 16/16.

### KB-11: `convert_to_list` on a soft-break paragraph makes one list item (live step 3)

- **Evidence:** live step 3 ("make the recitals a list"). The recitals are one paragraph with soft line breaks (`A. …` / `B. …` / `C. …`). `convert_to_list` succeeded and was verified, but Accept All shows one list item `A.` that still contains the B and C lines, and the model told the user the recitals were now an A, B, C list.
- **Cause:** library `list-format convert-plain` strips only the first typed marker and does not refuse a paragraph whose soft-break lines carry further markers. Post-commit verification passes because Word committed exactly what was prepared.
- **Options:** the plugin refuses `convert_to_list` on targets containing soft breaks (a product-scope rule like the insertion refusals), and/or the library refuses multi-marker soft-break paragraphs (`LIST_MARKER_UNSUPPORTED`). Library work is parked by the owner.
- **Fix (2026-10-28, owner chose it):** `convert_to_list` refuses before preparation (`UNSUPPORTED_DOCUMENT_EDIT_TARGET`) when the target, the `through` paragraph or any paragraph between them contains a soft line break, and asks the user to split it into separate paragraphs. Editing guide v6 states the rule. Pinned by `document_edit_intents_tests`. The library-side refusal is still worth raising when library work resumes. **Still to confirm:** live step 3. Replay host lanes on the fix (2026-10-28): Office.js passed (20 cases), Word oracle 88/0, NDA replay 16/16.

### KB-12: Second live run, both models (2026-10-29)

- **Runs:** fast `gemini-flash-latest` (`.cache/golden/live-fast-2`) failed 8 of 51 checks, down from 16. Steps 3, 10, 12, 13 and 14 now pass, which confirms KB-10 and KB-11. Slow `gemini-2.5-pro` (`.cache/golden/live-slow`, first run) failed 20 of 51.
- **Plugin issues fixed (2026-10-29):**
  - An empty `search` was refused ("Search must contain 1 to 1000 characters"). Fast step 1 opened with `search: ""` and never recovered. A blank search now reads from the top.
  - Unsupported-field errors now list the action's fields. Both models sent `position` with `insert_list_item` and `replace` with `rewrite_paragraph`, then needed a retry (fast/slow step 7, slow step 8) or failed twice (slow step 16).
  - `insert_list_item` accepts `position: "after"` (its only position). `"before"` gets a recovery hint.
  - A `rewrite_paragraph` containing line breaks now says to use `replace_text` for a paragraph that already has Shift+Enter breaks. Slow step 16 tried to rewrite the soft-break recitals.
- **Model behaviour addressed in editing guide v7:** the slow model asked for content instead of inspecting or drafting (steps 1, 2), declined a document review as legal advice (step 14), and gave up on a bold label in a new paragraph rather than inserting and then formatting (fast step 1). Guide v7 asks it to inspect before replying, never ask for text it can find or draft, treat review comments as expected, use only each action's fields, insert-then-format new formatted text, and use `replace_text` in soft-break paragraphs.
- **Not product defects:** slow step 3 (correctly refused the soft-break recitals and explained why), steps 9 and 11 (correctly refused unsupported structure with no change), slow step 13 (asked which date, since the NDA has no calendar date), slow step 10 (asked whether "BC" means "British Columbia"). The live checker scores some correct refusals as failures because it looks for refusal wording or the absence of any `apply_document_edits` call.
- **Unexplained:** fast step 5 timed out at 207 s after two provider requests returned no status. That looks like provider latency or retries; the edit itself was well formed. Watch it on the next run.
- **Replay host lanes on the fixes (2026-10-29):** Office.js passed (20 cases), Word oracle 88/0, NDA replay 16/16.
- **Third run on guide v7 (2026-10-30):** fast (`live-fast-3`) failed 8 of 51 checks, and steps 1, 7, 10, 13 and 14 now pass. Slow (`live-slow-2`) failed 9 of 51, down from 20.
  - **Remaining product issue (KB-13):** fast steps 5 and 11 ran out of time. In both, Gemini Flash requests that think for thousands of tokens got no response within the client's 60 s per-request timeout. Each step failed after three attempts, so step 5 hit the 180 s request limit and step 11 ended with "The AI response failed after tool execution". Fast step 5 also timed out in the second run.
  - **Model judgement, not product defects:** slow step 1 asks for the instruction text instead of drafting it; slow step 10 asks whether "BC" means "British Columbia"; steps 3, 4 and 11 refuse unsupported structure with no change, which the checker scores as a failure when the refusal wording is missing or an edit was attempted; fast step 2 recovered from a stale handle with a fresh inspection.
  - **Minor:** step 12 (both models) converted all nine headers correctly, then kept inspecting and ended with "Stopped after the maximum number of model turns" instead of a summary.
- **Slow run on the deadline fix (`live-slow-3`, 2026-10-30, gemini-2.5-pro):** all 16 steps finished, and 47 of 51 checks passed (about 12 minutes; no timeouts). The failures had four causes:
  - **Empty closing turns (steps 5, 7, 8, 14, 15, 16):** after a verified edit, Gemini returned `STOP` with no content. The add-in already answers with the edit receipt, so the user saw a correct reply. The live observer now counts these as `emptyClosingTurns` instead of invalid responses; an empty turn anywhere else is still a model problem.
  - **Step 10:** the model searched for "BC" literally, found "British Columbia", and asked instead of editing. Guide v8 asks it to read clear abbreviations as the user means them.
  - **Step 11:** the second attempt (a rewrite plus three inserts after the same paragraph) was correctly refused (one text change per paragraph per batch), and the loop guard then stopped with a generic line. The stop now says that no changes were made and gives the last refusal reason. The checker accepts that stop as the last call's outcome, because its receipt never reaches a later request.
  - **Step 13 → 14:** step 13 ("highlight the date") also added a comment nobody asked for, which pushed step 14's comment count to three. Guide v8 asks for only the requested changes. The comments check now counts only the comments its own step added, and highlight-only steps (NDA 13, privacy 06) check that no comment was added.
  - Replay lanes on these changes: NDA 16/16 (52 checks), privacy 10/10 (34 checks). The guide changes still need a live run.
- **Slow run on guide v8 (`live-slow-4`, 2026-10-30):** all 16 steps finished, and 46 of 52 checks passed with no model problems. Steps 10 ("BC") and 14 (comments) now pass. The remaining failures:
  - **Step 3:** a correct refusal ("each clause needs to be in its own separate paragraph"), which the checker's wording list missed. The list now covers "needs to be in its own paragraph".
  - **Step 12:** all nine headers were converted, but the model searched for one heading per turn (13 searches) despite the guide, and its edit used the last allowed turn. The user saw "Stopped after the maximum number of model turns". When the turn limit lands right after a verified edit, the stop now reports that edit. The evidence check still fails, because the receipt never reaches a later request; the slow searching is the finding.
  - **Model judgement (open):** steps 9 and 11 substituted a smaller edit for an unsupported request (a same-level item instead of a 2.2.1 sub-item, and one rewritten paragraph without the bullets). They then closed with an empty turn, so the user saw only "Applied and verified 1 requested edit(s)" and was not told what was left out. **Fix (2026-10-30, owner chose it):** after an empty closing turn following a write in the same request, the loop asks the model once for a summary of what it changed and what it did not or could not do. A second empty turn still falls back to the receipt. Guide v9 forbids swapping in a different structure for an unsupported part, and asks the model to name anything it left out. Needs a live run. Step 13 asked for the agreement date, since the NDA has no calendar date.
- **Harness: "File In Use" hangs.** Each launch opened the same sideload document (named after the add-in ID). A Word left from an earlier run kept it open, so the new Word waited on a "File In Use" dialog and the add-in never started. That is the hang and `GeneralException` seen on 2026-10-30. The golden runner and the Office.js lane now open a per-run copy in the artifacts folder. Confirmed with a leftover Word open: NDA replay 16/16, privacy 10/10, Office.js passed (20 cases).

### KB-13: Long-thinking Gemini requests exceed the 60 s per-call timeout (live run, 2026-10-30)

- **Where:** `TIMEOUT_LIMITS` in `taskpane.js`: `FETCH_TIMEOUT_MS` 60 s per provider call (three attempts) and `TOTAL_REQUEST_TIMEOUT_MS` 180 s per user request.
- **Evidence:** `live-fast-3` steps 5 and 11 (`gemini-flash-latest`). The turns that succeeded around the failures used up to about 6,000 thinking tokens. The failures took about 60 s each, which is consistent with the client timeout rather than rate limiting (the driver records only `requestError`, so the code is inferred).
- **Options (owner decision):** raise the per-call timeout (for example to 120 s) and the request limit (for example to 300 s); cap the thinking budget for the fast profile; or accept the occasional timeout. The UI keeps working throughout, and a timeout never writes or replays an edit.
- **Fix (2026-10-30, owner chose it):** `FETCH_TIMEOUT_MS` 60 s → 120 s and `TOTAL_REQUEST_TIMEOUT_MS` 180 s → 300 s; the loop's own default matches. The request limit is checked between turns, so a call already in flight can run past it (at most three 120 s attempts); the user can cancel at any time, and nothing is replayed. The live NDA lane's step timeout is raised to 480 s to cover that. Replay host lanes on the change (2026-10-30): Office.js passed (20 cases), Word oracle 88/0, NDA replay 16/16.
- **Fourth fast live run (`live-fast-4`, 2026-10-30):** steps 1–10 passed, apart from step 3's correct soft-break refusal; step 5 now finished in 128 s. The run then stopped: step 11 had three calls with no response, and with the longer per-call timeout one model turn could run three 120 s attempts, past the harness's 480 s step limit.
- **Deadline fix (2026-10-30):** `requestGeminiJson` takes an absolute `deadline`. Each attempt's timeout is capped by the time left, no attempt or retry starts with under five seconds left, and the error reads as a timeout. The agent loop passes the request's 300 s deadline, so a request can no longer outlast its limit. The live driver now records `requestErrorKind` (`aborted` for the client's own timeout, deadline or cancel; `network` otherwise) and `elapsedMs` per call. Pinned by `gemini_client_tests`. Replay host lanes: Office.js passed (20 cases), Word oracle 88/0, NDA replay 16/16. **Confirmed (2026-10-30, `live-fast-5`):** all 16 steps finished with 47 of 51 checks passing, and no provider call failed (median call 2.9 s, p90 17 s, slowest 69 s). The four failed checks are checker strictness: step 2 recovered from a stale handle, step 3 is the correct soft-break refusal, and in step 5 the model bolded exactly "By" and "Title" as asked while the checker expects the colons too.

### KB-14: List item insertion after the body's last paragraph tracks the final mark (main, library 0.11.0)

- **Where:** main's `insert_list_item` encoding, a `list-change` with typed markers and `structuredContent: true`, which goes through the library's list regeneration.
- **Evidence (2026-10-30, offline probe on 0.11.0):** with a decimal list item as the body's last paragraph, inserting an item after it gives the new final paragraph a tracked inserted mark (`<w:pPr>…<w:rPr><w:ins/></w:rPr></w:pPr>`), the same mark Word cannot reject cleanly (see KB-9). It also redlines the target item as deleted and re-typed. Library Accept All and Reject All read correctly, so only Word shows the problem. No gate covers this shape: the NDA's list insertion is mid-list.
- **User impact:** adding a list item after a list that ends the document can leave a stray paragraph after Reject All in Word, plus a noisier review. Mid-document list insertion is unaffected.
- **Fix path:** the library's 0.11.1 pre-release (`caeba20`) fixes structured list regeneration at a container's end, and the plugin moves `insert_list_item` to plain-text redlines, which avoid list regeneration entirely.
- **Verified on 0.11.1-pre.2 (2026-10-30):** branch `plain-insertions-0111` (current main plus plain-text insertions, `74cf986`, guide wording, regression tests `6edd61e`) passes offline 47/47, Office.js (20 cases), Word oracle 88/0 and the NDA replay 16/16. The new test that appends at body end fails on 0.11.0 and passes on pre.2. **Waiting on:** 0.11.1 on npm; then pin it, merge the branch, and re-run the lanes and live checks.

### KB-15: Tracked edits drop footnote references after Accept All (library 0.11.0 and 0.11.1-pre.2; upstream #6, #7, #8)

- **Where:** any text edit in a paragraph with footnote references, through `applyOperationsToDocumentXml`. The library deletes the word before a reference together with the reference run, then reinserts the word without the reference.
- **Evidence (2026-10-30, own probe `.cache/fn-repro/probe.mjs`):**
  - a reference at the paragraph's end: 1 → 0 after Accept All (#6);
  - a bookmark before a mid-paragraph reference: 1 → 0 (#7);
  - five references with an edit to the first sentence: 5 → 1 (#8).
  Reject All gives back every reference, so the library's fidelity check passes, and every operation returns `ok`. The behaviour is the same on 0.11.0 and on 0.11.1-pre.2 (`caeba20`).
- **Impact:** silent data loss for a reviewer who accepts the AI's edits in a document with footnotes. No gate covers it, because neither golden document has footnotes.
- **Status:** reported to the library agent on 2026-10-30, asking for a fix in 0.11.1 (or an Accept All fidelity refusal). There is no add-in workaround, by policy. After the fix: add a footnoted paragraph to the oracle or golden fixtures.
- **Fixed in 0.11.1-pre.3 (library 851b03c, 2026-10-30):** references are mapped through the diff and kept at their boundary. The library also refuses any edit whose Accept All view would lose a footnote or endnote reference, a drawing or an embedded object (`REVISION_FIDELITY_VIOLATION`, invariant `ACCEPT_NON_TEXT_CONTENT`). The add-in passes that refusal and its recovery to the model unchanged. pre.3 also adds the library-side KB-11 refusal (`LIST_MARKER_UNSUPPORTED`, reason `soft-break-markers`). Verified on branch `verify-0111-pre3` (main plus `plain-insertions-0111`, pre.3 installed with --no-save): probe 1/1, 1/1, 5/5; new `document_edit_intents_tests` footnote cases (two `replace_text` edits and one `rewrite_paragraph`); offline 48/48; Office.js passed (20 cases); Word oracle 88/0; NDA replay 16/16; privacy replay 10/10. **Waiting on:** 0.11.1 on npm; then pin it exactly and merge the branch into main.

### KB-9: Plain-text "after" insertion tracks the new paragraph's mark (library 0.11.0)

- **Where:** `@ansonlai/docx-redline-js` 0.11.0, the plain-adjacency insertion route used by a `structuredContent: false` redline whose `modified` adds a paragraph after the target.
- **Behaviour:** the new paragraph's own mark gets the `<w:ins>`. The structured route (`structuredContent: true`) marks the target's mark, which is Word-native. At body end the plain route therefore tracks the document's final paragraph mark, which Word cannot reject cleanly. Library Reject All still reports the right text.
- **Evidence (2026-10-18):** with the plugin's insertions switched to plain text, the Word oracle failed `direct-basic-underline-format` (Reject All 55 vs 56 paragraphs) because a previous case's tracked-append seed left `<w:p><w:pPr><w:rPr><w:ins w:author="Other reviewer"/></w:rPr></w:pPr></w:p>` at body end, carried over by Word's `insertFileFromBase64` Replace. Reproduced twice; the structured encoding passes 88/0. A minimal reproducer was sent to the library session.
- **User impact today:** none. The plugin reverted to the structured insertion recipe (`c82c190`).
- **Exit condition:** a library release that marks the target's mark for "after" insertions on the plain route. Then re-apply the plain-text insertions (phase-2 WP1 step 6) and pass Office.js, the Word oracle and the NDA replay.

### KB-8: Test XML provider was chosen silently; the `jsdom` lane fails 5 suites (all provider-specific, none a user defect)

- **Where:** [`tests/setup-xml-provider.mjs`](../../tests/setup-xml-provider.mjs). `d5809b9` (the issue #2 fix) added `jsdom` as a dev dependency, and the setup then preferred it whenever it was importable.
- **Resolved (2026-10-12):** the provider is now explicit. The default is `@xmldom/xmldom` (48/48 suites pass). `DOCX_TEST_XML_PROVIDER=jsdom` selects `jsdom` and fails loudly if it is missing. Unknown values throw. [`TESTING.md`](../TESTING.md) documents it. Since 2026-10-13 the default lane configures nothing and uses the library's own xmldom fallback, with no test-installed DOM globals.
- **Production environment:** the add-in never calls `configureXmlProvider`, so the library uses the host's native `DOMParser`/`XMLSerializer`. In Word that is the WebView2 (Chromium) DOM. Every failure below was re-run against real Chromium (Microsoft Edge 155, headless) through the library's browser bundle (`dist/docx-redline.bundle.js`, 0.9.0), alongside xmldom and `jsdom` in Node. Same scenarios, same inputs.
- **Root cause of 4 of 5 failures (class B, `jsdom` deviation):**
  - When `jsdom` serializes a subtree, an attribute whose namespace is not declared on that subtree gets a generated prefix (`w14:paraId` becomes `ns1:paraId`, `r:id` becomes `ns1:id`). Chromium and xmldom keep the original prefix.
  - The library serializes paragraph, `pPr`, `rPr` and run fragments and strips every `xmlns:` declaration with a regex before re-embedding them (`pipeline/serialization.js` `stripNamespaceDeclarations`, `pipeline/patching.js`, `engine/rpr-helpers.js`, `engine/table-cell-context.js`). That is only safe when the serializer keeps the prefixes the surrounding document declares, as Chromium does.
  - Under `jsdom` the stripped `ns1` is unbound (parse error), or `r:id` stops being found (see below).
  - Proof that this is the only cause: a scratch shim that maps generated prefixes back to `w14`/`r` makes all four suites pass under `jsdom`.
- **Per-failure classification:**

| Suite | Class | Evidence |
| --- | --- | --- |
| ~~`browser_document_session_tests`~~ | B | Suite removed with the browser demo (2026-10-17). The same `w14:paraId` serializer deviation still shows in the remaining class-B suites. Since then the jsdom lane reports 4 known failures, not 5. |
| `docx_cross_host_parity_tests` ("portable core") | B | A localized replace in a paragraph with `<w:hyperlink r:id>` is refused with `REVISION_FIDELITY_VIOLATION` under `jsdom` (`r:id` serialized as `ns1:id`, so the rel id reads back empty: `hyperlink:rIdLink` vs `hyperlink:`). xmldom and Chromium: `ok`, `r:id="rIdLink"` kept. The guard fails closed, so no corruption results. |
| `ooxml_formatting_visual_tests` | B | Same hyperlink `r:id` refusal, in the tab/line-break/hyperlink paragraph (`combined`, line 186). |
| `addin/integration_tests` ("Parties to Table") | B | Text-to-table over two NDA paragraphs: `PACKAGE_OPERATION_FAILED`, unbound `ns1` (the later `w:p` elements use `ns1:paraId` without a declaration). xmldom and Chromium: `ok`, Accept All resolves. |
| `phase4/golden-guardrail` | C | Correct output, provider-specific bytes. `table_reconcile`: xmldom and Chromium emit redundant `xmlns:w` on imported children, and their output is byte-identical after id/date normalization (3085 chars). `jsdom` omits them (2304). `comment_package_injection`: xmldom keeps the XML declaration plus a newline, Chromium keeps the declaration without the newline (1 char shorter), and `jsdom` drops the declaration. All are equivalent XML. The baseline is an xmldom-byte baseline, so any other provider (Chromium included) mismatches on `comment_package_injection`. Hashes were not changed. |

- **Class A (library defects that also occur in Chromium):** none found. No library report was filed.
- **Library robustness (optional, not a defect):** the `xmlns:` strip regexes assume the provider keeps prefixes. Re-using the original prefixes explicitly, or checking that every prefix in a stripped fragment is bound in the target context, would make the library tolerant of spec-style serializers such as `jsdom`'s. Only worth doing if a `jsdom`/Node-based consumer matters. It is not a user issue.
- **User impact:** none observed. Production runs on Chromium, where all five scenarios pass. Not tested: Word hosts that fall back to the legacy IE11/EdgeHTML webview, which this add-in does not target.
- **Exit condition (recommendation):**
  - Do not make the `jsdom` lane a required gate. It reports false failures, and a fix would force either a library change that no user needs or a prefix shim in the tests.
  - Keep xmldom as the required gate. Replace `jsdom` as the browser-like signal with a real-browser lane: headless Chromium or Edge loading `dist/docx-redline.bundle.js` and running the isomorphic scenarios (multi-op batch on a `w14:paraId` fixture, hyperlink `r:id` edit, text-to-table on the NDA sample, accept/reject). Make it advisory first, then required once stable in CI.
  - Until that lane exists, keep `DOCX_TEST_XML_PROVIDER=jsdom` as an opt-in, non-gating diagnostic. Its 5 known failures are classified above.
  - Keep `jsdom` as a devDependency: `chat_ui_security_tests` and `model_settings_tests` use it for UI DOM tests, independent of the XML provider.

## Word JS API usage (host boundary)

Inventory of `src/` at `a707b85` plus the 0.9.0 working tree. Dev-only entries are excluded: `scripts/officejs-validation.js` and `scripts/golden/golden-driver.js`.

All document *transformation* happens in the library. The only Word call that changes content is a whole-body `insertOoxml(..., 'Replace')`. The tables below give every use, why it exists, and its exit condition.

**Necessary:** the library cannot reach the live document. **Expedient:** the library could do it, or the use is leftover.

### Source read (Flat OPC body)

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `chat/word-document-source.js` | `body.getOoxml()` | Library inspection needs the live body package | Necessary | Only a Word-supported way to read the document package directly |
| `docx-redline-js-integration/word-operation-runner.js` | `body.getOoxml()` | Source for the single batch | Necessary | Same |
| `chat/document-tools.js` | `getOoxml` (wrapped) | Freshness re-read after checkpoint, before commit | Necessary | Same |

### Commit

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `word-operation-runner.js` | `body.insertOoxml(payload, 'Replace')` | One whole-body write of the library-prepared package | Necessary | Only a different host API. The complete-DOCX facade and binary bridge were rejected by the owner |

### Tracking mode save, disable and restore

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `docx-redline-js-integration/word-ooxml.js`, `word-operation-runner.js`, `chat/checkpoint-restoration.js` | `document.changeTrackingMode` load, set to off, restore | Word would wrap the replacement in its own revisions. The library already produced the tracked output | Necessary until disproved | Evidence that `insertOoxml('Replace')` preserves the package's revisions unchanged with tracking on |

### Selection, identity and platform

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `chat/word-document-source.js` | `document.getSelection()`, `load('text')` | "User highlighted text" hint. The library has no selection concept | Necessary | None. A future improvement could map the selection to inspected paragraph handles instead of Word text |
| `taskpane.js` | `Office.context.document.url`, `Office.onReady`, `Office.context.platform` | Document-identity guard, host gate, and the platform passed to the library | Necessary | None |
| `chat/word-document-source.js` → `word-source-identity.js` `detectExportTailIndex` | `body.paragraphs.load('$none')`, `items.length` | Host paragraph count, used only to recognise the empty trailing paragraph Word adds to `getOoxml` exports | **Expedient** | The library recognises the Word export tail from Flat OPC alone, or the extra paragraph is shown never to occur. Then remove the count and `detectExportTailIndex` |

### Verification and checkpoint restore

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `chat/document-tools.js` | `body.getOoxml()` after commit | Evidence of what Word actually committed, compared through library inspection | Necessary | None |
| `chat/checkpoint-restoration.js` | `body.insertOoxml(saved, 'Replace')`, then `getOoxml`, then `paragraphs.load` | Whole-body restore and token comparison | Necessary (the paragraph count is the expedient tail check above) | None for the restore itself |

### Author lookup and glance checks

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `taskpane.js` (~L561) | `document.properties.load('lastAuthor, author')` | Pre-fills the redline author on refresh | Convenience, host-only | Optional. An explicit author setting exists |
| `taskpane.js` (glance checks) | none since WP4 (2026-10-16) | Prompt text is built from library inspection (`inspection.paragraphs[].exactText`) through the shared source reader, not Word's `body.text` | Retired | — |

### Local package logic

This is not Word API use, but it is outside the library:

- **`consumer-core.js` and `word-source-identity.js`:** Flat OPC part and relationship handling, comment and numbering merge, insertion-package assembly, and Word export-noise normalisation. **They stay in the plugin** by owner decision (2026-10-17): the library stays host-independent. Their remaining issue is imports of unexported library internals (see phase-2 WP1).
- Moved by phase-2 WP3 (library 0.10.0): the foreign-revision policy choice (now `existingRevisions: 'auto'`) and the post-commit review-segment comparison (now `compareDocumentInspections`).

### Expedient uses to retire (priority order)

1. Host paragraph count for the export tail: `detectExportTailIndex` needs `hostParagraphCount` to certify that Word's trailing empty paragraph is an export artefact and not real content. Removing it needs Word evidence that the tail is always present, or a Flat OPC marker that identifies it. This is plugin work now that the Word layer stays here.

Retired by phase-2 WP3 (2026-10-17): the `TRACKING_POLICY_MISMATCH` guard (KB-5).

Retired by phase-2 WP4 (2026-10-16): the glance-check `body.text` read, the unreachable checkpoint-capture `getOoxml` fallback in `taskpane.js`, the unused `baseTrackingMode` option, the unimported `src/taskpane/debug_structure.js`, and the Outlook-template `src/commands/commands.js` (the manifest's `FunctionFile` still points at the now script-free `commands.html`).
