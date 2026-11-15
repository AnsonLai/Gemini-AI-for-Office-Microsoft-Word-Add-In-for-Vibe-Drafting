# Known bugs: closed entries in v2.4.0.1

Closed on 2026-11-14 by `@ansonlai/docx-redline-js` 0.11.3 (0.11.2 was never published; 0.11.3 is the same code). The entries are moved here from [KNOWN_BUGS](../KNOWN_BUGS.md), with the original text unchanged except for the added closing note.

### KB-21: Deleting a paragraph is refused when the same author's pending inserted paragraphs follow it

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

**Closed 2026-11-14 by library 0.11.3.** Verified with the library alone (`.cache/repro/kb21-library-only.mjs`): the stage 2 delete as author A now returns `ok` (the control on the untouched source also `ok`; as author B it is still refused, the expected `EXISTING_REVISIONS` policy). Deleting your own pending inserted paragraph now removes it.

### KB-19: A replacement drops a pending same-author bold in the same paragraph

Found 2026-10-31 by the privacy replay in real Word. With track changes on, bold a phrase (a pending `rPrChange` by the add-in's author), then change other wording in the same paragraph with `replace_text`: the bold is gone after Accept All. Nothing reports it, because the prepared package itself has no bold, so the add-in's after-write check matches. It happens with one replacement as well as several. Settled (accepted) bold is kept, and with a different author the library refuses the replacement instead.

**Impact:** a user who bolds words and then edits that paragraph's wording in the same review session silently loses the bold. **Add-in:** no workaround (library defect). `tests/document_edit_intents_tests.mjs` prints a `KNOWN_LIBRARY_DEFECT: KB-19` line while it reproduces. **Library ask:** keep pending same-author run formatting through a localized replacement. Reproducer: paragraph runs `This notice explains how ` + `Maple Clinic` (`<w:b/>` with `<w:rPrChange w:author="A"><w:rPr/></w:rPrChange>`) + ` collects your information.`; operation `{ type: 'replace', replacements: [{ find: 'collects', replace: 'gathers' }] }` with author `A`; Accept All has no bold.

**Closed 2026-11-14 by library 0.11.3.** `tests/document_edit_intents_tests.mjs` now asserts the pending same-author bold is kept after a replacement in the same paragraph, and passes; the `KNOWN_LIBRARY_DEFECT` print is gone.

### KB-18: Overlapping format operations fail with an internal receipt error

Two `format` operations on one paragraph whose text overlaps (for example `WHEREAS` and `WHEREAS, Provider`, or bold on and bold off for the same word) are refused safely, but with `RECEIPT_RECONCILIATION_FAILED` ("Committed revision id '1000' (kind: rPrChange) was not found in word/document.xml"). Nothing is written. Separate, non-overlapping format operations on one paragraph work and pass Accept All and Reject All.

**Add-in today:** the intent compiler refuses overlapping `format_text` spans in one paragraph with a clear message before the library runs (2026-10-31). That check stays as a validation rule either way. **Library ask:** refuse overlapping format spans with a specific code and message. Reproducer: one paragraph `RECITALS: WHEREAS, Provider owns it.`, operations `{ type: 'format', textToFormat: 'WHEREAS', properties: { bold: true } }` and `{ type: 'format', textToFormat: 'WHEREAS, Provider', properties: { bold: true } }` in one batch.

**Closed 2026-11-14 by library 0.11.3.** Verified with the library alone (`.cache/repro/kb18-library-only.mjs`): the two overlapping bold `format` operations on `RECITALS: WHEREAS, Provider owns it.` now return `ok` with two `rPrChange` revisions, not `RECEIPT_RECONCILIATION_FAILED`. The add-in's own overlapping-span refusal in the intent compiler stays as a validation rule.
