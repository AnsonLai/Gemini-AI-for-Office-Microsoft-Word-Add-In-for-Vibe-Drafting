# List-format edit intents (phase-2 WP5)

**Date:** 2026-10-14
**Library:** exact `@ansonlai/docx-redline-js@0.9.0`
**Host:** desktop Word 16.0.20430
**Plans:** [phase-2 WP5](../plans/completed/2026-10-12-library-first-simplification-phase-2.md), [list migration follow-up](../plans/completed/2026-09-30-canonical-list-migration-follow-up.md) WP1–WP2, [KNOWN_BUGS KB-4](../plans/KNOWN_BUGS.md)

## What changed

`apply_document_edits` has three new actions. Each maps one-to-one onto a 0.9.0 `list-format` operation; the add-in parses no markers and writes no numbering.

| Action | Library operation |
| --- | --- |
| `convert_to_list` (`target`, optional `through`, `format`, `markers: typed\|none`, `start` with `none` only) | `list-format` `convert-plain` (`strip: marker\|none`, optional `targetEnd`) |
| `change_list_format` (`target` list item, `format`, optional `level`, `start`) | `list-format` `change-format`, `scope: 'list'` |
| `continue_list` (`target` later-list item, `from` earlier-list item) | `list-format` `join`, `scope: 'list'` |

**Same-paragraph pairs.** Exactly one `rewrite_paragraph` plus one single-paragraph `convert_to_list` may target the same paragraph (a changed-text header such as `1. Old` → `1. New`). Every other same-paragraph pairing with a list action is refused before preparation with `UNSUPPORTED_DOCUMENT_EDIT_COMBINATION`.

**Receipts.** Failed receipts now carry `error.failedOperations` (`editIndex`, `code`, `message`, `recovery`), so the library's specific refusal reaches the model instead of a bare `BATCH_OPERATION_FAILED`.

**Editing guide.** Now v5 (2026-10-14). List conversion and list-format changes are supported. Depth/indent changes, range-scoped splits, table lists and table structure stay refused.

**Not exposed:** range-scoped `change-format`, `paragraph-format.numbering`, depth/indent, `detach`, table lists.

**Size:** production code went from 4,437 to 4,549 lines (+112, 20 files) for three capabilities. All of it is validation and mapping.

## Evidence

| Lane | Result |
| --- | --- |
| Offline `npm test` | 49/49 suites, including the new `document_list_intents_tests.mjs`. Production tools with the real library check Accept All and Reject All on synthetic and Word-authored fixtures: non-contiguous decimal headers form one list with `w14:paraId` kept; a `C.` start; bullets → decimal; join continuing a count; the changed-text header pair; and library refusals (`LIST_MARKER_UNSUPPORTED`, `LIST_MARKER_AMBIGUOUS`, `LIST_FORMAT_STYLE_NUMBERED`) with no write. |
| Actual Office.js | Passed. The default manifest is now 20 cases: the original 14 unchanged, plus 5 list capability cases and 1 refusal control, with 26 host checks. |
| Independent Word oracle | 88 PASS, 0 FAIL, 22 not applicable (engine-reference packages). List labels are read from Word's `ListFormat.ListString` after `Revisions.AcceptAll()` / `RejectAll()`. A negative control that corrupted two expected labels failed as intended. |
| 16-prompt NDA replay (real Word) | 16/16 steps, 51/51 checks (was 49), 40 chat requests, 0 auxiliary. Step 12 ("turn all the headers into a proper ordered list") is now an applied edit, no longer a scripted refusal. All nine headers share one decimal list, and the manual numbers are removed. |

### New host cases

All sources were authored in Word through COM ([`create-agentic-list-action-fixtures.ps1`](../../scripts/create-agentic-list-action-fixtures.ps1)). Expectations are independent of the engine.

- **L1 `list-convert-noncontiguous-typed-headers`:** three typed headers with body paragraphs between them. Accept All gives 1./2./3.; Reject All restores the original typed text with no numbering.
- **L2 `list-convert-letter-headers-start-at-c`:** labels C./D.
- **L3 `list-change-format-bullets-to-decimal`:** Accept All gives 1./2./3.; Reject All restores the bullets.
- **L4 `list-continue-second-list-after-first`:** Accept All gives 3./4.; Reject All restores 1./2.
- **L5 `list-rewrite-and-convert-changed-header`:** Accept All gives label 1. with `New title`; Reject All gives the literal `1. Old title`.
- **L6 `list-convert-ambiguous-roman-marker-refusal`** (safety control, not a capability pass): `LIST_MARKER_AMBIGUOUS`, zero content writes, zero checkpoints, source unchanged.

## Notes and limits

- **Golden replay shape.** The replay serves step 12's nine header handles from one model turn of parallel read-only `inspect_document` calls. A live model may inspect differently, and the live fast/slow provider lane was **not run**.
- **`change-format` scope.** It redefines only the target's level, so nested bullet levels stay bullets. This is the library's documented behaviour.
- **Not run:** Word Online, and browser/MCP facade parity for `list-format` (list plan acceptance item).
- **No library defects found.**
