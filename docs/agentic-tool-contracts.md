# Word agent tool contracts

The Word agent exposes **two tools** through one registry and a bounded local loop using the existing Gemini client. One versioned editing guide defines the workflow. Mutation tools make no nested model calls. Browser and MCP consumers retain their own contracts.

## Inspect current source

`inspect_document({search?,around?,limit?,cursor?,includeDocumentEnd?})` reads current body Flat OPC and uses the library's accepted-view inspector. `around` is 0–5; `limit` is at most 8. A cursor continues its original query and cannot be combined with new search options.

Targets contain opaque handles, human locations, text, visible truncation, and compact list/table/revision/comment cues. Strong canonical descriptors remain private. Handles bind document instance, revision view and transported part tokens, and expire after writes, external changes, Undo/source changes or request completion. Truncated text cannot authorize editing its unseen remainder. The end handle binds the last real paragraph for append composition. Freshness checks do not lock Word against human edits.

Word regenerates session attributes and an export-only empty tail ID. These confirmed fields are normalized only in derived freshness tokens; raw source/checkpoints/payload remain intact. Meaningful paragraph/revision IDs, authors, text, styles, numbering, relationships and comments remain covered.

## Apply one typed batch

`apply_document_edits({edits:[...]})` accepts 1–40 intents. Each requires `type` and a fresh `target` handle. Unknown actions, extra fields, unsupported shapes, stale targets and unsafe combinations reject the **whole batch before writing**.

| Action | Additional fields | Scope |
| --- | --- | --- |
| `replace_text` | `find`, `replace`, optional `occurrence` | Exact literal text within one paragraph, including existing list/table-cell text. Occurrence is 1-based; no paragraph breaks. |
| `rewrite_paragraph` | `text` | Complete desired literal text for one existing paragraph. |
| `delete_paragraph` | None | Plain paragraphs; structural list/table deletion refuses. |
| `insert_paragraph` | `text`, `position:before` or `after` | One plain paragraph beside an inspected nonempty plain paragraph. |
| `append_paragraph` | `text` | End handle; one paragraph at a supported plain end. |
| `insert_list_item` | `text` | Same-level insertion after an active bullet/decimal item at library levels 0–1, outside tables. |
| `comment` | `text`, optional `anchor` | One comment on exact original inspected text; existing comments and threads are preserved when adding more. |
| `format_text` | `find`, optional `occurrence`, `properties` | Boolean bold/italic/underline/strikethrough; highlight `yellow` or `none`. |
| `convert_to_list` | optional `through`, `format`, `markers`, `start` | Plain, non-list, non-table paragraphs become a new list through the library's `list-format` `convert-plain`. `target` (and a later `through` for a contiguous run) are fully inspected handles. `markers` `typed` (default) strips the typed marker (`1.`, `A.`, `iv.`, `-`) as a tracked deletion and keeps its value as the label; `format` is optional and disambiguates `I.`/`C.`. `markers` `none` keeps the text, requires `format` and accepts `start`. One `convert_to_list` per header converts non-contiguous headers; consecutive values join one list. |
| `change_list_format` | `format`, optional `level` (0-8), optional `start` | Changes the numbering format of the whole list (`change-format`, scope `list`) of an inspected list item outside tables. The library redefines one level (the target's by default) in a new numbering definition; range scope is not exposed. |
| `continue_list` | `from` | The list of `target` continues the earlier list of the inspected list item `from` (`join`, scope `list`). |

Inserted content is one plain paragraph. Literal replacements/rewrites use `structuredContent:false`; paragraph/list insertion uses explicit canonical structured recipes. List conversion and format changes map one-to-one onto library `list-format` operations; the add-in parses no markers and writes no numbering. `format` values are `decimal`, `upperLetter`, `lowerLetter`, `upperRoman`, `lowerRoman`, `bullet`. These actions never change the paragraph count. No font family/size, layout, table creation/rows, list depth/indent changes, list range splits, revision acceptance, arbitrary OOXML, research or navigation executor is exposed.

Each source paragraph has one text-bearing desired state per batch. Combine wording changes into its complete rewrite. Text plus formatting on that paragraph requires a deliberate dependent stage with fresh inspection. Rewrite plus comment is allowed only when its original anchor remains unique in the resulting paragraph. Comments on newly created text require fresh inspection. Format plus comment can share unchanged source text. A list action may share a source paragraph with exactly one `rewrite_paragraph`, and only as a single-paragraph `convert_to_list` (no `through`): the rewrite keeps the typed marker (`1. Old` becomes `1. New`) and the library runs it first. Every other combination of a list action with another action on one paragraph, including two list actions, is refused as `UNSUPPORTED_DOCUMENT_EDIT_COMBINATION`. Cross-item conflicts of a list-wide rebind are left to the library (`LIST_FORMAT_COMPOSITION_CONFLICT`). Unsupported combinations never silently drop intents or become progressive writes.

## Commit and completion

All handles resolve from one fresh source. The tool prepares one atomic library batch, verifies complete operation accounting, persists the exact-source checkpoint, checks freshness again, and makes at most one body OOXML insertion. Native tracking is disabled for transport and its prior state restored. Fresh canonical text/structure, pending review segments/authors and comments (including reply relationships and resolved state) must match preparation; requested formatting must reapply as a pure no-op. Review-segment comparison tolerates Word revision-ID renumbering and equivalent adjacent carrier splitting while requiring the same text and view offsets.

A library refusal reaches the model unchanged and makes no write. The atomic wrapper (`BATCH_OPERATION_FAILED`) hides which operation failed, so the receipt's `error.failedOperations` lists each refused operation as `{ editIndex, code, message, recovery }` (for example `LIST_MARKER_AMBIGUOUS` with `recovery.allowedValues`). Recovery fields use library names: `list.format` is the intent's `format`.

Receipts include requested/accounted counts, status/success/completion/verification, write evidence, mutation outcome, checkpoint identity, counters and useful errors. Preparation/freshness source reads and verification reads are separate from content writes. Checkpoint failure blocks a new write. Verified no-ops succeed without writing or checkpoint capture. Identity literal replacements refuse; an already-satisfied complete rewrite is the supported no-op confirmation.

Checkpoint save resolves only after its IndexedDB transaction commits. Manual restore replaces the body once with native tracking disabled, restores its prior mode, and compares all transported parts. The restoration comparison omits confirmed Word session metadata and regenerated body/comment paragraph IDs. It additionally requires identical public canonical comment identity, including replies, resolved flags, authors, dates, text and anchors. Metadata references must resolve to inspected comments; durable comment IDs and all remaining part content stay covered. Normal source freshness still covers real IDs; stored XML is untouched. Failed or uncertain restoration preserves chat history and requires inspection. Verified restoration resets model history and the visible transcript together.

New-comment output from 0.8.4 is a merge delta. The Word transport composes the
published comments merge and sibling reconciliation helpers over in-memory
parts, then assembles the existing Flat OPC insertion payload. It does not
replace existing comment definitions with the delta or switch to a full-DOCX
Word facade. The [dated transport report](validation-reports/2026-10-17-library-issues-closeout.md) (in git history; see the closeout)
records the regression and correction evidence.

Pending revisions alone do not prohibit an edit. For single-paragraph wording
changes on targets with foreign or unattributed revisions, the adapter selects
the library's `slice-cross-author` policy per operation to preserve those review
carriers. Own-only targets retain the documented same-author baseline merge.
Structural/list actions
keep their established library policies. The add-in's tracked-edit setting is
respected. Localized
replacements and formatting preserve unaffected revisions. Changes inside
reviewed text use the library's surgical slice semantics; the affected span can
change while unrelated review marks remain pending. Whole same-author rewrites retain the library's
established merge policy: re-diff against that author's prior baseline, producing
clean text when tracking is disabled. Specific unsupported revision structures
still produce library refusals. No accept-all-first policy is selected.

Published 0.8.4 had an untracked surgical-insertion defect: a pure insertion
could create new own tracked text despite tracking being disabled. It is fixed in
0.9.0, and untracked pure insertion beside a foreign revision now succeeds (the
original capability gate passes, 2026-10-12). The library enforces the rule
itself (`UNTRACKED_CREATED_REVISION`), so the add-in's `TRACKING_POLICY_MISMATCH`
backstop was retired in phase-2 WP3 (KB-5 in [KNOWN_BUGS.md](plans/KNOWN_BUGS.md)).
Plain-text `replace_text`/`rewrite_paragraph` operations send
`existingRevisions: 'auto'`; the library slices around another reviewer's
pending revisions and merges the user's own. Foreign review carriers are
allowed. WP5 is closed.

Uncertain writes, verification errors, tracking restoration errors and cancellation preserve observed host evidence and require inspection. There is no native fallback or automatic mutation replay. The loop pairs every function response, rejects multiple mutation calls and mixed inspect/mutate exchanges before execution, and retains six model turns, two failed mutation cycles and one no-write stale refresh. Read-only calls do not reset mutation failures. Malformed recovery accepts only complete JSON calls for declared names. An unresolved refusal overrides a fabricated final success.

See the [plan and milestone evidence](plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md). Historical native/library reports remain dated evidence; retiring a route does not certify its former structural edits or close an upstream issue. Live fast/slow model checks are required before shipping the prompt. Desktop is the initial target; Word Online editing is unverified.
