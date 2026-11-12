# Known bugs and host-boundary exceptions

**Updated:** 2026-11-09 (KB-20 and KB-21 added, KB-17 status). Baseline 2026-10-30, after v2.3.0.0 shipped on `@ansonlai/docx-redline-js` 0.11.1 (exact pin). Closed entries KB-2 to KB-16 are archived, unchanged, in [closed entries through v2.3.0.0](completed/2026-10-30-known-bugs-closed-v2.3.0.0.md).

This file tracks three things:

- Defects and gaps we know about but have not fixed.
- Work deliberately deferred until after v2.3.0.0.
- Every place production code uses the Word JavaScript API (Office.js) instead of the library.

The goal is to use the library wherever possible. Each Word API use must say why it exists and what would let us remove it.

**Library defects** (`@ansonlai/docx-redline-js`, which the owner maintains) are reported to the library (its repository or the library session) with a portable reproducer that uses only public APIs. Never work around a library defect in the add-in. Word-specific glue, such as Flat OPC handling and Word export noise, stays in the add-in.

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


### KB-21: Deleting a paragraph is refused when the same author's pending inserted paragraphs follow it (library, 0.11.1)

Found 2026-11-11 in a live `gemini-4-carbon` run (see KB-17) and reproduced with the library alone. With track changes on, insert new paragraphs after a paragraph, then delete that paragraph as the same author: the delete is refused with `REVISION_FIDELITY_VIOLATION` ("Reject All of this operation's output would not restore the source (REJECT_PARAGRAPH_TEXT … expected \"Old provision.\", got \"\")") and the atomic batch rolls back. Nothing is written.

- **Cause, as far as seen:** the insertion marks the *original* paragraph's paragraph mark as a pending insertion (`<w:pPr><w:rPr><w:ins/></w:rPr></w:pPr>`), which is how Word itself records paragraphs added after one. A delete of a paragraph whose mark is a pending same-author insertion then produces output that fails the library's own Reject All check.
- **Controls that work:** the same delete on the untouched source; the delete after the insertions are accepted; deleting a paragraph elsewhere in the document; and inserting the new paragraphs *before* the paragraph, then deleting it. As a different author the delete is refused with `EXISTING_REVISIONS`, which is the expected policy refusal.
- **Impact:** replacing one paragraph with several in two stages (insert the new ones after it, then delete the old one) cannot be finished in the same review session. The model is left with the old paragraph and the new ones both in the document, and the user has to revert.
- **Add-in:** no workaround (library defect or missing case). The refusal reaches the model as `BATCH_OPERATION_FAILED` with the failed operation's message.
- **Library ask:** let a same-author delete go through when the target's paragraph mark is a pending insertion, or refuse it with a specific code and a recovery hint (for example, insert before the paragraph instead).
- **Reproducer** (public API, `@ansonlai/docx-redline-js/standalone-runner`): body `Before.` / `Old provision.` / `After.`; options `{ atomic: true, structuredContent: true, pairReplacements: true, generateRedlines: true, strictTargets: true }`; author `A`. Stage 1: `{ type: 'redline', target: { exactText: 'Old provision.', index: 2 }, modified: 'Old provision.
New opening:
Point one.
Point two.
Point three.', structuredContent: false }` returns `ok`. Stage 2 on that output: `{ type: 'delete', target: { exactText: 'Old provision.', index: 2 } }` returns `BATCH_OPERATION_FAILED`.

### KB-19: A replacement drops a pending same-author bold in the same paragraph (library, 0.11.1)

Found 2026-10-31 by the privacy replay in real Word. With track changes on, bold a phrase (a pending `rPrChange` by the add-in's author), then change other wording in the same paragraph with `replace_text`: the bold is gone after Accept All. Nothing reports it, because the prepared package itself has no bold, so the add-in's after-write check matches. It happens with one replacement as well as several. Settled (accepted) bold is kept, and with a different author the library refuses the replacement instead.

**Impact:** a user who bolds words and then edits that paragraph's wording in the same review session silently loses the bold. **Add-in:** no workaround (library defect). `tests/document_edit_intents_tests.mjs` prints a `KNOWN_LIBRARY_DEFECT: KB-19` line while it reproduces. **Library ask:** keep pending same-author run formatting through a localized replacement. Reproducer: paragraph runs `This notice explains how ` + `Maple Clinic` (`<w:b/>` with `<w:rPrChange w:author="A"><w:rPr/></w:rPrChange>`) + ` collects your information.`; operation `{ type: 'replace', replacements: [{ find: 'collects', replace: 'gathers' }] }` with author `A`; Accept All has no bold.

### KB-18: Overlapping format operations fail with an internal receipt error (library, 0.11.1)

Two `format` operations on one paragraph whose text overlaps (for example `WHEREAS` and `WHEREAS, Provider`, or bold on and bold off for the same word) are refused safely, but with `RECEIPT_RECONCILIATION_FAILED` ("Committed revision id '1000' (kind: rPrChange) was not found in word/document.xml"). Nothing is written. Separate, non-overlapping format operations on one paragraph work and pass Accept All and Reject All.

**Add-in today:** the intent compiler refuses overlapping `format_text` spans in one paragraph with a clear message before the library runs (2026-10-31). That check stays as a validation rule either way. **Library ask:** refuse overlapping format spans with a specific code and message. Reproducer: one paragraph `RECITALS: WHEREAS, Provider owns it.`, operations `{ type: 'format', textToFormat: 'WHEREAS', properties: { bold: true } }` and `{ type: 'format', textToFormat: 'WHEREAS, Provider', properties: { bold: true } }` in one batch.

### KB-17: Live-model judgement gaps (live runs, 2026-10-30)

Seen with gemini-2.5-pro on guide v9 (`live-slow-5`). None involves the library, and every edit was verified, so nothing was written incorrectly:

- **Staged substitution:** asked to rewrite a provision with bullets, the model built it from a rewrite plus plain paragraphs over several turns, instead of declining the bullets. Guide v9 forbids this; it still happens with this model.
- **Overclaiming:** the model converted 6 of 9 headers and said it had converted all of them. The edit receipt was accurate ("Applied and verified 6 requested edit(s)").
- **Clarifying question:** the model asked which date to highlight when the NDA has no calendar date (also a release-notes known issue).

**Next step, if pursued:** compare the receipt's count with the request in the model's closing turn, or have the loop's summary quote the receipt. Re-check with the next live runs.

**Status 2026-11-11 (live runs in Word, NDA scenario):** still open. Every edit that landed was verified.

| Model | Prompts | Checks | Miss |
| --- | --- | --- | --- |
| `gemini-flash-latest` | 16 | 51 of 52 | step 11 |
| `gemini-flash-latest` (second run) | first 11 | 33 of 34 | step 11 |
| `gemini-4-carbon` | first 11 | 33 of 34 | step 11 |
| `gemini-4-carbon` | first 5 | 16 of 17 | step 3 (scorer wording only) |
| `gemini-2.5-pro`, `gemini-3.8-flash` | first 5 each | 17 of 17 | none |

Step 11 (`11-required-disclosure-rewrite`) asks for a rewritten provision plus a few bullets; the scenario expects the whole request to be refused.

- **Flash, both runs: split request.** It rewrote the paragraph (one call, verified), added no bullets and said so in its reply. It did not substitute plain paragraphs and did not overclaim, but the guide says not to split an unsupported mixed request.
- **Carbon: staged attempt, left half done.** (1) A rewrite plus an insertion on one paragraph was refused (`CONFLICTING_DOCUMENT_EDITS`). (2) It inserted the new opening and three points as plain paragraphs after the original (verified). (3) It then sent `delete_paragraph` for the original with `convert_to_list` (bullet) for the new points; the library refused the delete with `REVISION_FIDELITY_VIOLATION` (`REJECT_PARAGRAPH_TEXT`: Reject All would give "" for the deleted paragraph) and the batch rolled back. (4) The next model call hit the loop's 5-minute total limit, so the request ended with "Confirmed edits remain in the document…" and Revert. The document was left with the old provision followed by the new plain paragraphs.
- **Open question:** whether step 11 should still be "unsupported" now that staged edits and `convert_to_list` with bullets exist. The delete refusal in (3) is reproduced with the library alone and tracked as KB-21.

Other observations: carbon was slow and unreliable on the day (steps of 60 to 250 s; in the 5-prompt run one 120 s client timeout and one 503 "high demand", both retried by the add-in; in the 11-prompt run one more failed provider response at step 1 that was not looked into). At step 3 in the 5-prompt run carbon declined correctly with "I could not convert…", which the refusal pattern in `scripts/golden/golden-checks.mjs` does not recognise.

**Guide v12 (2026-11-11), step 11 now a supported edit.** Three guide changes: the "never split an unsupported mixed request" phrase is gone (it contradicted "do only the supported part and say what you left out"); a new list from new text is described as two stages (insert plain paragraphs, inspect, `convert_to_list` with markers "none" and a format); replacing one paragraph with several is described as rewrite first, then insert the rest (no delete, so KB-21 is not on the path). The scenario's step 11 is now a three-stage replay scored on "intro paragraph then bullets", and the refusal pattern accepts "could not". Results in Word the same day: replay 16/16 (53 checks); live first 11 prompts 35 of 35 on both `gemini-flash-latest` (step 11 in 38 s) and `gemini-4-carbon` (step 11 in 133 s). A later run the same day with `gemini-flash-latest` on v12 passed all 16 NDA prompts (53 of 53) and the privacy scenario (16/16, 65 of 65). One run each; the slow profile and Carbon were not run past prompt 11 on v12. The overclaiming and clarifying-question items above were not re-tested.

### KB-20: "Revert anyway" discards later changes with no way back in the add-in

When the document changed after an edit, Revert asks first ("The document changed after this edit. Reverting will also undo those later changes, including any of your own typing."). Choosing "Revert anyway" replaces the body with the checkpoint taken before the edit, so everything done since, including the user's own typing, is gone. The add-in saves no checkpoint of the state it discards and offers no redo. Finding 4b in [taskpane state machines](../taskpane-state-machines.md).

- **Impact:** a user who confirms the prompt loses later work unless Word's own Undo brings it back. Whether Undo restores it after the body replacement has not been checked in Word.
- **What works:** the warning is shown, the revert is refused while a request runs and the prompt is kept (check M9 in the taskpane lane, 2026-11-09).
- **Owner direction (2026-11-09):** leaning towards replacing the Revert system with Word Undo now that edits are applied step by step, so this was left as is.
- **Exit condition:** either Revert is replaced by Undo, or the add-in saves a checkpoint of the current body before a "Revert anyway" write (or it is confirmed in Word that Undo restores the discarded changes, and the prompt says so).

### KB-8: The `jsdom` XML lane is a diagnostic only

The default test lane uses `@xmldom/xmldom`. `DOCX_TEST_XML_PROVIDER=jsdom` is an opt-in diagnostic that reports 4 known, provider-specific failures and no user defect, because Word runs the add-in in Chromium (WebView2). Don't make it a required gate. The full classification is in the [archive](completed/2026-10-30-known-bugs-closed-v2.3.0.0.md#kb-8-test-xml-provider-was-chosen-silently-the-jsdom-lane-fails-5-suites-all-provider-specific-none-a-user-defect).

## Deferred after v2.3.0.0

Known and intentionally left for later (owner decision, 2026-10-30; moved here from the archived [phase-2 plan](completed/2026-10-12-library-first-simplification-phase-2.md)):

- **Issue #3, automatic model fallback.** After transient retries run out, the add-in does not switch to another model on its own. Custom model IDs in Settings are the workaround; decide later whether to fall back automatically.
- **Word Online.** Not re-checked for v2.3.0.0, by the owner's choice; listed as a known issue in the release notes.
- **Host paragraph count for the export tail.** The last expedient Word API use (see "Expedient uses to retire" below).
- **KB-1** (body-final table append, not exposed) and **KB-8** (`jsdom` diagnostic lane).
- **Lint.** 1,572 Prettier errors predate this work.
- **Dev certificate.** The local HTTPS dev certificate lasts 30 days, and host lanes hang when it expires. Renew it with `npx office-addin-dev-certs install`; the current one expires on 2026-11-11.

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

- **`consumer-core.js` and `word-source-identity.js`:** Flat OPC part and relationship handling, comment and numbering merge, insertion-package assembly, and Word export-noise normalisation. **They stay in the plugin** by owner decision (2026-10-17): the library stays host-independent. Since 0.11.0, `src/` imports only the package root and `/standalone-runner`.
- Moved by phase-2 WP3 (library 0.10.0): the foreign-revision policy choice (now `existingRevisions: 'auto'`) and the post-commit review-segment comparison (now `compareDocumentInspections`).


### Expedient uses to retire (priority order)

1. Host paragraph count for the export tail: `detectExportTailIndex` needs `hostParagraphCount` to certify that Word's trailing empty paragraph is an export artefact and not real content. Removing it needs Word evidence that the tail is always present, or a Flat OPC marker that identifies it. This is plugin work now that the Word layer stays here.

Retired by phase-2 WP3 (2026-10-17): the `TRACKING_POLICY_MISMATCH` guard (KB-5).

Retired by phase-2 WP4 (2026-10-16): the glance-check `body.text` read, the unreachable checkpoint-capture `getOoxml` fallback in `taskpane.js`, the unused `baseTrackingMode` option, the unimported `src/taskpane/debug_structure.js`, and the Outlook-template `src/commands/commands.js` (the manifest's `FunctionFile` still points at the now script-free `commands.html`).

