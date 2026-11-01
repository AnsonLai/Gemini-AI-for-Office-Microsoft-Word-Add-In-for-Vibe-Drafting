# Known bugs and host-boundary exceptions

**Updated:** 2026-10-30, after v2.3.0.0 shipped on `@ansonlai/docx-redline-js` 0.11.1 (exact pin). Closed entries KB-2 to KB-16 are archived, unchanged, in [closed entries through v2.3.0.0](completed/2026-10-30-known-bugs-closed-v2.3.0.0.md).

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


### KB-18: Overlapping format operations fail with an internal receipt error (library, 0.11.1)

Two `format` operations on one paragraph whose text overlaps (for example `WHEREAS` and `WHEREAS, Provider`, or bold on and bold off for the same word) are refused safely, but with `RECEIPT_RECONCILIATION_FAILED` ("Committed revision id '1000' (kind: rPrChange) was not found in word/document.xml"). Nothing is written. Separate, non-overlapping format operations on one paragraph work and pass Accept All and Reject All.

**Add-in today:** the intent compiler refuses overlapping `format_text` spans in one paragraph with a clear message before the library runs (2026-10-31). That check stays as a validation rule either way. **Library ask:** refuse overlapping format spans with a specific code and message. Reproducer: one paragraph `RECITALS: WHEREAS, Provider owns it.`, operations `{ type: 'format', textToFormat: 'WHEREAS', properties: { bold: true } }` and `{ type: 'format', textToFormat: 'WHEREAS, Provider', properties: { bold: true } }` in one batch.

### KB-17: Live-model judgement gaps (live runs, 2026-10-30)

Seen with gemini-2.5-pro on guide v9 (`live-slow-5`). None involves the library, and every edit was verified, so nothing was written incorrectly:

- **Staged substitution:** asked to rewrite a provision with bullets, the model built it from a rewrite plus plain paragraphs over several turns, instead of declining the bullets. Guide v9 forbids this; it still happens with this model.
- **Overclaiming:** the model converted 6 of 9 headers and said it had converted all of them. The edit receipt was accurate ("Applied and verified 6 requested edit(s)").
- **Clarifying question:** the model asked which date to highlight when the NDA has no calendar date (also a release-notes known issue).

**Next step, if pursued:** compare the receipt's count with the request in the model's closing turn, or have the loop's summary quote the receipt. Re-check with the next live runs.

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

