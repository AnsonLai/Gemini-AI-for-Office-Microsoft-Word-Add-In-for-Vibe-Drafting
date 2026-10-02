# Agentic Tool Contracts (Current Consumer)

This inventory records the tool declarations and execution paths currently wired in `src/taskpane/taskpane.js` and `src/taskpane/modules/commands/agentic-tools.js`. The library contract references are from the installed and locked `@ansonlai/docx-redline-js@0.8.4` package. This is a description of the current consumer; it does not imply that every tool already submits a canonical operation.

The engine extraction predates the four August 29 plans. The later redline
batch migration delegates `apply_redlines` to one atomic library batch, removing
the former iterative Word loop and manual table recovery for that tool.
Canonical list delegation remains limited to tracking-on insertion into active
bullet/decimal lists through `insert_list_item` at levels 0–1. `edit_list`
retains its legacy `applyRedlineToOxml` path; header conversion and unsupported
insertion shapes remain native Word work. Provider retry, host outcome
reporting, and lazy tool loading are consumer responsibilities, not library
offloads. See
[package boundaries](package-boundaries.md) and the
[library offload review](library-offload-review.md).

## Source indexes and target conventions

- The enhanced context enumerates `context.document.body.paragraphs` in document order and labels them `[P1]`, `[P2]`, and so on. These are 1-based indexes into Word's body paragraph collection, not visual line numbers. Paragraphs inside tables also receive a `P` index; their metadata includes `T:<row>,<cell>`.
- Other context metadata includes the paragraph style, list kind and nesting level (`ListNumber` or `ListBullet`, `L:<level>`), and section markers (`§` or `§N`). Metadata is for tool selection and targeting; it is not paragraph text.
- Table row and column arguments (`targetRow`, `targetColumn`) are 0-based. The `T:` marker positions come from `parentTableCellOrNullObject.rowIndex` and `cellIndex`.
- Paragraphs can shift after an edit. The chat carries a canonical prompt-time baseline into `apply_redlines` and `insert_list_item`; execution compares it before a write. The optional insertion baseline preserves compatibility for callers that omit it. Other structural tools resolve their supplied `P` indexes against the live Word paragraph collection and do not inherit that full-source freshness contract.
- If enhanced context extraction fails, the chat path falls back to plain `body.text`. In that case there are no `[P#]` anchors for the redline input validator to parse. The redline mapper still resolves its operation indexes against the live source read by the Word adapter.
- After a complete tool exchange reports a confirmed successful mutation, the taskpane rereads document text and the canonical baseline before the next model turn. A proven no-write `STALE_DOCUMENT_CONTEXT` refusal additionally permits one fresh snapshot per request, followed by model replanning without replaying the refused batch or resetting the failed-mutation budget. It does not refresh between calls already queued for one source snapshot; other refusals and uncertain writes do not trigger this recovery. A failed refresh stops further mutation; the external-edit stale-source guard remains in force. Safe mismatch diagnostics distinguish unavailable/invalid baselines, missing/invalid targets, count, text and fingerprint mismatches without logging source text or paragraph IDs.

## Mutating tool map

| Tool and declared arguments | Executed arguments and operation | Current Word transport |
| --- | --- | --- |
| `apply_redlines({ instruction: string })` | The instruction and anchored document are passed to the diff generator. It returns changes with `paragraphIndex`, `anchorText`, `operation`, and either `newContent`, `replacements`, or the applicable `content` / `originalText` / `replacementText`; `replace_range` also has `endParagraphIndex`. The consumer maps these to library `{ type: 'redline', target: { index, exactText, paragraphId?, fingerprint? }, targetRef: 'P#', structuredContent: true, modified }` or `{ ..., replacements: [{ find, replace, occurrence? }] }`. `replace_range` adds `targetEndRef`. | `applyRedlineChangesToWordContext` reads the body OOXML once, plans all operations against that source, and calls `executePureOoxmlBatch` / the library `applyOperationsToDocumentXml` with `atomic: true`. A successful batch is inserted once into the Word body with `insertOoxml(..., 'Replace')`. |
| `insert_comment({ instruction: string })` | The suggestion generator returns `{ paragraphIndex, textToFind, commentContent }`. For each suggestion, the consumer calls the library with `{ type: 'comment', targetRef: 'P1', target: targetParagraph.text || textToFind, textToComment: textToFind, commentContent }`. `P1` is local to the single-paragraph OOXML scope; `paragraphIndex` selected that scope from the document. | One Word run iterates the suggestions. Each suggestion gets its own `executePureOoxmlBatch` read/prepare/write through the shared paragraph adapter; the library inserts comment anchors and comment package parts. The suggestions are not one atomic batch. |
| `highlight_text({ instruction: string, color?: enum })` | The suggestion generator returns `{ paragraphIndex, textToFind }`. The consumer calls the library with `{ type: 'highlight', targetRef: 'P1', target: targetParagraph.text || textToFind, textToHighlight: textToFind, color: normalizedColor.toLowerCase() }`. | Same per-suggestion shared paragraph adapter and OOXML `Replace` transport as comments. The consumer turns native track changes off while applying generated OOXML, then restores the prior tracking state. |
| `edit_list({ startParagraphIndex, endParagraphIndex, newItems, listType, numberingStyle? })` | `startParagraphIndex` and `endParagraphIndex` are 1-based inclusive. `newItems` are strings; `listType` is `bullet` or `numbered`; `numberingStyle` is `decimal`, `lowerAlpha`, `upperAlpha`, `lowerRoman`, or `upperRoman`. The consumer passes items through the library's `normalizeListItemsWithLevels` and `buildListMarkdown`, then calls the older `applyRedlineToOxml` helper, passing the document's `numberingXml`. **No canonical `DocumentOperation` is constructed here.** | Source-independent request validation runs first; the executor also checks both indexes against the live paragraph count before resolving the range. It reads range OOXML, computes reconciled list OOXML, then replaces that range once with Word `insertOoxml`. |
| `insert_list_item({ afterParagraphIndex, text, indentLevel? })` | Strict 1-based target and relative indentation -1/0/+1. With redlining enabled, active bullet/decimal lists at source and resolved levels 0–1 map through `planAgenticListOperations` to a source-bound canonical redline insertion; outdent requires a following same-list root sibling. | The verified subset uses one body OOXML read and one atomic batch insertion. Plain anchors, other levels/styles, tracking-off requests and unsupported outdent contexts retain native paragraph/list insertion. Native capability selection happens before engine preparation; engine/host errors never replay natively. |
| `edit_table({ paragraphIndex, action, content?, targetRow?, targetColumn? })` | The schema uses nested `content: string[][]`. `paragraphIndex` is a strict positive 1-based paragraph index; row and column indexes are strict non-negative 0-based integers. `replace_content` overlays the provided 2D cells; `add_row` appends exactly one row; `delete_row` uses `targetRow`; `update_cell` uses both target coordinates and exactly one cell. The validator normalizes accepted legacy flat add-row / scalar or one-element update-cell inputs to nested arrays and refuses inputs that would be silently truncated. A Markdown-formatted `update_cell` alone maps to `{ type: 'redline', targetRef: 'P1', target: currentCellText, modified: normalizedValue }` in a cell-range scope. **Other actions do not construct canonical operations.** | The executor validates request shape before Word access, then validates the target paragraph and row / per-row cell dimensions from Word before staging cell edits. Direct Office.js `Table` / `TableRow` / `TableCell` APIs handle replacement, row insertion, row deletion, and plain-text cell updates. Markdown cell updates use the shared range adapter and OOXML replacement. The action enum has no column insertion or deletion. |
| `edit_section({ sectionHeaderIndex, newHeaderText?, newBodyParagraphs?, preserveSubsections? })` | `sectionHeaderIndex` is 1-based and must resolve to a list item. `newHeaderText` omits the list marker; `newBodyParagraphs` supplies replacement body paragraphs; `preserveSubsections` controls the section scan. **No canonical `DocumentOperation` is constructed here.** | Direct Office.js calls update the header paragraph, delete existing body paragraphs, and insert new paragraphs. The tool can perform multiple separate writes within one Word run. |
| `convert_headers_to_list({ paragraphIndices, newHeaderTexts?, numberingFormat? })` | `paragraphIndices` are 1-based. `numberingFormat` is `arabic`, `lowerLetter`, `upperLetter`, `lowerRoman`, or `upperRoman`. The request validator binds each supplied text to its original index before sorting, rejects duplicate indexes, and returns the sorted paired records. Header text is derived by stripping a manual prefix unless `newHeaderTexts` is supplied. **No canonical `DocumentOperation` is constructed here.** | Direct Office.js calls replace each header's text, start a list on the first header, set its numbering, and attach later headers to that list. These are multiple separate writes within one Word run. |

### Source-independent request validation

`src/taskpane/modules/commands/agentic-request-validation.js` exports `validateListRequest(request, paragraphCount?)`. It returns `{ valid: true, request }` or `{ valid: false, error: { code: 'INVALID_LIST_REQUEST', message } }` without Word or Office dependencies.

- `edit_list` requires strict positive integer indexes, ordered inclusive bounds, non-empty string items, and exact list / numbering enum values. If a source paragraph count is supplied, both indexes must fit. Item strings are copied unchanged so the existing library list-level normalizer can interpret custom indentation.
- `convert_headers_to_list` requires distinct positive integer indexes and, when supplied, exactly one non-empty text string per index. It binds `{ paragraphIndex, text? }` records before sorting and returns sorted `paragraphIndices`, matching `newHeaderTexts`, and `headerRecords`.
- `insert_list_item` requires a strict positive integer `afterParagraphIndex`, non-empty single-paragraph `text`, and an integer `indentLevel` of exactly `-1`, `0`, or `1` (default `0`). The executor calls the validator before Word access and checks the index against the live paragraph count before reading the target OOXML.
- All three production list executors call the validator. The focused invocation suite verifies invalid preflight, live-range refusal before mutation, and that `[3, 1]` keeps its original text pair after sorting.

`src/taskpane/modules/commands/table-request-validation.js` exports `validateTableRequest(request, live?)` with the same result shape and `INVALID_TABLE_REQUEST` code. It requires a positive integer source paragraph, a supported action, and action-specific arguments. Delete/update row and update column coordinates must be actual non-negative integers; numeric strings and fractions are rejected. Replacement content must be a non-empty matrix of string cells; add-row input must describe one non-empty row; update-cell input must describe exactly one string cell. Empty cell strings remain valid so callers can clear cells. The normalized request uses 2D `content` for replace/add/update.

The optional `live` record can supply `paragraphCount`, `rowCount`, `columnCount`, and `rowCellCounts`. Production calls validate against the current paragraph collection and the loaded row cell counts before staging mutations. Replacement overlays may be ragged or partial, but every supplied row and cell must exist; add-row values cannot exceed the appended row's capacity; delete/update coordinates must fit the current table. The production suite covers oversized overlays refused before the first write, fractional-coordinate and multi-row add preflight, canonical nested add/update inputs, and the provider schema's actual nested-array shape / supported action list.

### Library operation types actually used

The 0.8.3 public `DocumentOperation` declarations include `redline`/`replace`, `restore`, `delete`, `comment`, comment-thread operations, `highlight`, character and paragraph formatting, `list-change`, `table-reconciliation`, and `insert`. The 0.8.3 release does not change API or schema declarations from 0.8.2; 0.8.4 adds the format operation's `textOccurrence` (used by `format_text`) and `applyRedlineToOxml`'s `numberingXml` option (passed by `edit_list`). The current consumer constructs only:

- `redline` for `apply_redlines`, the verified `insert_list_item` subset and Markdown-formatted table-cell updates;
- `comment` for `insert_comment`;
- `highlight` for `highlight_text`.

Markdown inline formatting and Markdown list/table parsing inside redline content do not create separate `format`, `list-change`, or `table-reconciliation` operations in this consumer. The dedicated list, section, and most table tools continue to use Word-specific code paths. The types above are from `node_modules/@ansonlai/docx-redline-js/services/standalone-operation-runner.d.ts`; the current mappings are in `word-redline-runner.js`, `word-operation-runner.js`, and `agentic-tools.js`.

## Redline input and exact-text contract

`apply_redlines` uses a full-document, immutable-source batch path. The verified
`insert_list_item` subset also submits a canonical body batch; its unsupported
shapes retain native Word execution selected before preparation.

- The model change schema supports `edit_paragraph`, `replace_paragraph`, `modify_text`, and `replace_range`. `edit_paragraph` accepts either a whole-paragraph `newContent` or non-empty localized `replacements`; `replace_paragraph` and `replace_range` use `content`.
- Each localized `find` must be non-empty and match the original paragraph exactly, including case and punctuation. `occurrence` is 1-based. A missing match or repeated match without `occurrence` is rejected before the batch write. A replacement may be the empty string to delete the matched text. Localized replacement cannot insert paragraph breaks; use a full-content operation for structural text.
- A non-empty `anchorText` is checked against the claimed paragraph and can correct a unique match within two paragraphs. If a supplied anchor cannot be verified unambiguously, the change is rejected. Missing or blank anchors remain a compatibility bypass in `verifyAnchor`, despite being required by the prompt/schema. With anchored source text, `replace_range` shifts the end index along with any corrected start index.
- The mapper resolves target descriptors from the batch's inspected paragraphs. It carries the source paragraph's exact text, index, and available `paragraphId` / fingerprint into the canonical operation. Appending at `paragraphCount + 1` is supported for paragraph-level content operations; the mapper resolves the last paragraph and adds the new content after it.
- One narrow coalescing case combines a compatible edit to the final paragraph with a single Markdown table append. The append anchor remains `paragraphCount + 1`; it is not rewritten to the last paragraph index. Both original change indexes map to the one library operation and its receipt. Incompatible edits and ambiguous overlap are refused before writing. As of 0.8.4, `UNSUPPORTED_TABLE_FORMATTING` refuses only changing the text of a paragraph that carries the author's own pending formatting while adding a table; new inline formatting with a table and appends after an underlined paragraph are allowed.
- Sanitization, anchor checks, and operation compilation happen before the one Word write. An invalid change may be rejected while other valid changes are applied if the remaining batch compiles and commits. A batch-level engine or host failure refuses the batch write or reports an uncertain host outcome; see the result section below.

## Outcome contracts exposed by the tools

The reliability fields describe different stages and should be preserved independently:

- `receipts` are library operation receipts (or the observer's collected single-operation receipts). They describe committed engine work; they do not alone prove that Word accepted a host write.
- `writeAttempted` records that the consumer queued or attempted a Word mutation. `written` means a host `context.sync()` confirmed at least one pending write. These flags are not interchangeable.
- `mutationOutcome` uses `noop`, `refused`, `prepared`, `rolled_back`, `applied`, `partial`, `indeterminate`, `applied_with_host_error`, or `failed` (the last is returned by the redline adapter's caught batch-runner exception) depending on engine and host evidence. `indeterminate`, `partial`, and `applied_with_host_error` require document inspection before any replay.
- `apply_redlines` returns library batch fields such as `status`, `results`, `receipts`, `written`, `writeAttempted`, `mutationOutcome`, plus `changesApplied`, `skipped`, `rejectedChanges`, `engineSkipped`, `error`, `message`, and `showToUser` as available. A valid mixed batch can have a confirmed write plus per-operation rejections.
- For a coalesced edit-plus-append, the consumer maps both original AI changes back to the shared library operation receipt. The receipt records engine work; host `written` evidence remains separate.
- Comment and highlight results spread `createMutationObserver().result()` into `{ status, success, written, writeAttempted, confirmedHostWrites, mutationOutcome, receipts, operationResults }` and add `message` / `showToUser`. Because suggestions are processed in sequence, a later failure can follow an earlier confirmed write; the result may be partial.
- The direct structural tools (`edit_list`, `insert_list_item`, `edit_table`, `edit_section`, `convert_headers_to_list`) return the observer fields plus `success` and `message`. A no-op may have `success: true` and `written: false`; the taskpane currently treats these tools' `success` as tool success. `apply_redlines`, comments, and highlights instead gate success on a confirmed write.
- Missing API keys return structured error objects across redline, comment, highlight, and navigation executors. Mutating tools report `MISSING_API_KEY`, `status: 'error'`, `success: false`, and a refused mutation outcome with no write attempted. Navigation also reports structured success/failure; success is true only after Word selects and syncs the target paragraph.

## Unsupported or behaviorally limited cases

- Exact replacement occurrence semantics currently exist only for localized `apply_redlines` changes. Comment and highlight suggestion objects have no occurrence field; the library rejects a missing or ambiguous exact anchor within the selected paragraph. The user-facing prompt also says anchors must match exactly.
- `apply_redlines` atomicity applies to that tool's library batch. It does not make a multi-call chat turn atomic across tools.
- `insert_comment` and `highlight_text` do not compile all suggestions into one batch. Each one reads and writes its target paragraph separately, so earlier suggestions can be committed before a later suggestion fails.
- `edit_list`, native-only `insert_list_item` shapes, `convert_headers_to_list`, `edit_section`, and most `edit_table` actions do not pass canonical document operations to the library. They use direct Office.js calls, OOXML stitching, or the older `applyRedlineToOxml` route. They therefore do not inherit the redline batch's all-source-at-once targeting, preflight, or atomic commit contract.
- `insert_list_item` on a non-list paragraph inserts a plain paragraph. Its request rejects relative indentation outside `-1..1`; its executor still clamps resolved Word levels to `0..8`.
- The 0.8.2 inspector could mistake `w:numPr` nested inside historical `w:pPrChange` properties for active list membership. The v0.8.3 release fixes that inspection behavior; its historical-inspection cutover test now requires historical numbering to be absent. The 0.8.2 reproduction remains historical evidence, not a current package behavior claim. The production observer still does not replay a native insertion after an engine failure. See [the historical list inspection issue](library-issues/completed/2026-09-30-historical-list-inspection.md) and [v0.8.3 validation report](validation-reports/2026-09-30-docx-redline-v083.md).
- `edit_table` now declares `content` as an array of arrays of strings, and the tool description lists only supported row/cell actions. The validator rejects unsupported column actions and oversized content rather than allowing the executor to ignore extra nested rows or cells. Table edits remain direct Office.js operations apart from Markdown-formatted cell updates.
- `convert_headers_to_list` now validates unique indexes and sorts paired text records before calling the current executor; direct Office.js mutation remains non-atomic. Its content is still changed through multiple host writes.
- `edit_list` content is list markdown generated by helper functions, not individual canonical list operations. Its request is checked against the live paragraph count before range resolution.
- `edit_section` recognizes section boundaries through Word list items and list levels. It does not accept an exact content anchor; a stale but still in-range index can target another list header.
- `apply_redlines` verifies `anchorText` against the original chat-turn document string. The taskpane also captures a prompt-time source baseline from the same Flat OPC inspector used by execution and passes it to both the redline adapter and `insert_list_item`. Before planning, the adapter compares each targeted live paragraph's canonical `exactText` and fingerprint against that baseline; it checks inclusive ranges, insert-before anchors, and the original count/final paragraph for appends. List insertion checks its source anchor before either canonical or native execution. A provided empty or malformed baseline and any stale/missing target fail closed as `STALE_DOCUMENT_CONTEXT` before a write. Callers that omit the optional baseline retain the legacy behavior. Other structural tools still use their own targeting paths.
- Column insertion/deletion, arbitrary table-to-table structural transformations through `edit_table`, non-contiguous section replacement beyond the tool's list-header scan, and canonical `list-change`/`table-reconciliation` calls are not exposed by the current tool mappings.
- Plain or text-changing header-to-list conversion and list-format changes still lack supported canonical operations. On bare `document.xml`, a marker-prefixed `1. Header` to `1. Header` operation can fail with `RECEIPT_RECONCILIATION_FAILED`.
- The 0.8.3 table append route lost underline on Accept All; 0.8.4 fixes it (see the [completed library issue](library-issues/completed/2026-10-01-table-append-tracked-formatting.md)). The remaining 0.8.4 limitation is refused before the Word write as `UNSUPPORTED_TABLE_FORMATTING`: changing the text of a paragraph that carries the author's own pending formatting while adding a table. The dated final 2026-10-01 Word run on 0.8.3 passed 13 of 20 checks: plain accepted table content and native insertion passed; tracked Reject All leaves an extra paragraph (still open on 0.8.4: one extra empty paragraph after a final table append) and the formatting fixture's Accept All lost underline. See the [host validation report](validation-reports/2026-10-01-table-creation-reliability.md) and [Reject All paragraph issue](library-issues/2026-10-01-word-reject-table-append-paragraph.md). Passing offline tests do not establish full Word fidelity. No full live chat/model reproduction is claimed.

## Current native Office.js mutation calls

These are the live Word calls that remain in tool execution, aside from the shared adapter's required OOXML boundary (`getOoxml`, `insertOoxml`, and `context.sync`):

- `insert_list_item`: `Paragraph.insertParagraph`, `ListItem.level`, and fallback `Range.insertOoxml`.
- `convert_headers_to_list`: `Paragraph.insertText`, `Paragraph.startNewList`, `List.setLevelNumbering`, and `Paragraph.attachToList`.
- `edit_table`: `TableCell.body.insertText`, `Table.addRows`, and `TableRow.delete`; Markdown-formatted `update_cell` uses the shared OOXML adapter.
- `edit_section`: `Paragraph.insertText`, `Paragraph.delete`, and `Paragraph.insertParagraph`.
- `edit_list`: not a canonical operation path; it calls `applyRedlineToOxml` (passing the document's `numberingXml`, read from `body.getOoxml()`) and replaces the selected Word range with the returned OOXML.

By comparison, `apply_redlines`, comments, highlights, and Markdown-formatted table-cell edits calculate changes through the installed library's standalone operation engine before using the shared OOXML Word adapter to commit them.

## Concrete gaps and recommended coverage

1. **Table tool contract:** schema nesting, strict source/row/cell bounds, unsupported column action refusal, and no-truncation behavior now have focused validation and production invocation tests. Continue covering ragged Word tables and Office.js failures after the executor starts staging native writes.
2. **Target freshness and atomicity:** `apply_redlines` now has same-inspector source freshness checks and a mixed-batch no-write refusal test. Migrate the remaining structural tools to source inspection plus a single canonical batch where supported. Verify whether native multi-sync tools report partial writes correctly.
3. **List operations:** cover relative indentation `-1/0/+1`, Word-level clamping at 0 and 8, non-list neighbors, numbering identity/restart/continuation, nested items, and out-of-range indexes. Specifically test unsorted and duplicate `paragraphIndices` paired with `newHeaderTexts`. The pPrChange inspection behavior recorded against 0.8.2 is reported fixed in 0.8.3; keep its dated evidence separate from any current validation.
4. **Exact anchors:** exercise missing and repeated `find` text, 1-based occurrence selection, empty `replace`, stale `anchorText`, range end correction, and a mixed batch with valid and invalid changes. Assert library receipts, `writeAttempted`, `written`, and `mutationOutcome` separately.
5. **Comment/highlight batches:** test missing and duplicate snippets within one paragraph, invalid paragraph indexes, a failure after an earlier suggestion committed, and whether the desired contract is per-suggestion partial success or one atomic batch.
6. **Outcome handling:** missing-key results and navigation success/failure now use structured objects. Continue covering no-op, preflight refusal, rollback, host failure before commit, and host error after a confirmed write across the remaining tool paths.

These are consumer-side gaps. If evidence points to a defect in `@ansonlai/docx-redline-js@0.8.4`, track it as a separate library issue; do not compensate for it in this consumer inventory or imply a consumer-side workaround is part of the library contract.

## v0.8.3 validation (dated record)

The 0.8.4 upgrade is recorded in the [0.8.4 validation](validation-reports/2026-10-02-docx-redline-v084.md) report.

The supported list matrix covers 12 cases with 48 actual Office.js checks; five
additional native routes pass 20 checks. The independent Word oracle reports
92 applicable checks passed, zero failed, and five not applicable across 17
exports. The two original public-facade Reject All cases pass 12 Word checks.
`npm test` passes 51 suites with zero failures and four exclusions. See the
[v0.8.3 report](validation-reports/2026-09-30-docx-redline-v083.md) and its
linked collector/oracle artifacts. This evidence does not add the missing
canonical header conversion or list-format operations.

The separate 2026-10-01 table incident has 54 passing offline suites and a
passing production build. Its final Word run passed 13 of 20 checks; tracked
Reject All paragraph structure and underline fidelity remained open on 0.8.3;
underline is fixed in 0.8.4 and the Reject All extra paragraph remains open. See the
[incident plan](plans/2026-09-30-table-creation-reliability.md) and
[validation report](validation-reports/2026-10-01-table-creation-reliability.md).
