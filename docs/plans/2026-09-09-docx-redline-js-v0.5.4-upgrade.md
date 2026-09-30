# `@ansonlai/docx-redline-js` v0.8.1 Direct Upgrade Plan

**Date:** 2026-09-23  
**Updated:** 2026-09-29

**Status:** WP0–WP3 implemented; WP4–WP6 remain — direct migration to v0.8.1. WP3 still needs Word Desktop validation in WP6.

**Target:** Upgrade both root add-in and `mcp/docx-server` directly to exact version `0.8.1`.

This plan supersedes and consolidates:
- `2026-09-02-docx-redline-js-v0.4.0-upgrade.md`
- `2026-09-07-docx-redline-js-v0.5.0-upgrade.md`

Because the four 2026-08-29 optimization plans were never executed against v0.5.4, an intermediate step through v0.8.0 is unnecessary. Upgrade directly to **v0.8.1** and use the document lifecycle introduced in v0.8.0 to retire temporary glue code (such as manual XML wrapping, synthetic `<w:document>` envelopes, and external JSZip handling). Verify those APIs and behaviors against v0.8.1 before removing existing code.

---

## 1. Summary of Major Library Capabilities (v0.5.0 &rarr; v0.8.1)

| Release Span | Architectural Breakthrough | Direct Impact on This Repository |
|---|---|---|
| **v0.5.1 – v0.5.4** | Multi-Author Revision Slicing (`slice-cross-author`), Paragraph Restoration (`restore`), and Baseline-Delta Validation. | Pre-existing document defects or existing tracked changes in user files no longer fail subsequent valid operations. |
| **v0.6.0 – v0.6.2** | **Caller-Order-Independent Batches**: Targets resolve against the immutable starting state of the document. | Prior splits or insertions in a batch no longer throw off later paragraph indexes, eliminating the need for manual bottom-up sorting or iterative single-paragraph loops in `word-redline-runner.js`. |
| **v0.6.0 – v0.6.2** | Compact Mutation Outputs & Receipts (`completion`, `written`, `status`, `receipts`). | Clean structured receipts allow caller verification without inspecting verbose OOXML dumps. |
| **v0.7.0 – v0.7.2** | **Localized Exact Replacements (`replacements: [{ find, replace }]`)**: Operations specify exact phrase substitutions inside a target paragraph. | **60–90% reduction in generated prompt tokens**; prevents model truncation and eliminates anchor mismatch rejections. |
| **v0.7.0 – v0.7.2** | Two-Turn Golden Path: Single computed success gate (`completion: true` + non-null output). | Streamlines machine recovery and verification. |
| **v0.8.0** | **Universal `openDocx` Facade & Zero-Node Architecture**: Runs natively in modern browsers without `node:zlib`, `node:crypto`, or `Buffer`. Ingests and serializes canonical `Uint8Array`. | Zero-Node packaging runs out-of-the-box in the Word taskpane and browser demo. |
| **v0.8.0** | **Built-in `fflate` ZIP Container**: ZIP extraction, OOXML part loading, and rebuilding are handled internally. | **Removes external `jszip` and `@xmldom/xmldom` dependencies** from `mcp/docx-server`. |
| **v0.8.0** | **Integrated Document Lifecycle**: `doc.inspect()`, `doc.applyOperations()`, and `doc.resolveRevisions('accept' \| 'reject')`. | Extracts paragraphs, applies batches, manages `numbering.xml`/`comments.xml`, and resolves revisions in one step. |
| **v0.8.1** | Fixes the `commentsExtended.xml` content type and comment-thread markers; adds comment resolve/reopen/delete-by-id and editing of existing header/footer parts. | Prevents Word repair prompts when editing commented files; requires regression coverage for comment preservation, thread operations, and body/part batch isolation. |

---

## 2. Migration Principles for Direct v0.8.1 Upgrade

1. **Pin `0.8.1` exactly** in both `package.json` and `mcp/docx-server/package.json`. Regenerate lockfiles via `npm install`.
2. **Eliminate Redundant Dependencies in WP5**: Keep `jszip` and `@xmldom/xmldom` until the MCP server stops importing them. Remove them with the WP5 facade refactor.
3. **Adopt Universal `openDocx` & Lifecycle Methods**:
   - In MCP: Use `openDocx(fs.readFileSync(path))` &rarr; `doc.inspect()` &rarr; `doc.applyOperations()` &rarr; `doc.toUint8Array()`.
   - In Word Add-in: Support both Document-Level Binary Single-Hop (`getFileAsync` &rarr; `openDocx` &rarr; `insertFileFromBase64`) and Range-Level OOXML Single-Hop (`applyOperationsToDocumentXml`).
4. **Transition to Order-Independent Batches**: Eliminate the 1,100-line iterative `for (const change of changes)` loop in `word-redline-runner.js`. Send operations as a single batch with `atomic: true`.
5. **Modernize Prompt & Validation for Localized Replacements**:
   - Update `redline-prompt.js` to encourage `replacements: [{ find, replace }]` for phrase/sentence edits.
   - Update `change-validation.js` to validate `find` targets within paragraph text rather than demanding full paragraph rewrites.
6. **Preserve Error Codes Generically**: Ensure structured error codes (`COMMENTED_CONTENT_DELETE`, `FOREIGN_PARAGRAPH_MARK_DELETION`, `GENERATED_OOXML_INVALID`, `PATCH_ROUNDTRIP_MISMATCH`, `UNSAFE_REVISION_BOUNDARY`, `TARGET_NOT_FOUND`, `AMBIGUOUS_ANCHOR`) pass cleanly through all host boundaries.
7. **Purge Obsolete Workarounds**: Delete the unused `word-structured-list.js` module and redline runner's manual table recovery. Retain active, supported list helpers and legacy bridge functions until their other callers migrate.

---

## 3. Work Packages

### WP0 — Establish Current Working Baseline
1. Check `git status --short`.
2. Confirm `@ansonlai/docx-redline-js@0.8.1` is published and compare its exports, types, dependencies, and behavior with v0.8.0. The user supplied the v0.8.1 release notes (commit `3e55a02`); npm registry lookup confirmed that exact version exists and that its package exports and direct dependency ranges match v0.8.0. Record the delta below and cover it in WP1.
3. Run existing suites:
   ```bash
   node tests/redline_result_contract_tests.mjs
   node tests/addin/word_operation_runner_adapter_tests.mjs
   node tests/addin/shared_operation_bridge_tests.mjs
   node tests/no_legacy_shared_operation_bridge_tests.mjs
   npm run build:dev
   ```
4. Record baseline before modifying dependencies.

**WP0 record (2026-09-29):** The root and MCP manifests and lockfiles initially pinned `0.5.4`. The working tree already contained documentation changes from the plan update. The four listed Node suites passed. `npm run build:dev` passed (webpack compiled successfully; non-blocking Application Insights telemetry errors appeared in the sandbox).

**v0.8.1 release delta (user-supplied notes, checked against the installed API declarations):** The CLI contract remains v8 and existing operation shapes remain valid. The release fixes the `word/commentsExtended.xml` content type, preserves untouched comment parts byte-for-byte, repairs the old type on a subsequent write, and rejects the old type during package validation. It corrects multi-paragraph thread keys and reply markers, preserves anchors during edits, and synchronizes optional comment sibling parts. It adds `comment_resolve`, `doc.resolveComment`, `doc.deleteComments({ ids })`, and opt-in `part` selectors for existing headers/footers; `inspect().headersFooters` exposes those parts. Body searches remain body-only. New failure codes include `COMMENT_NOT_FOUND`, `PARENT_ANCHOR_NOT_FOUND`, `FIELD_EDIT_REFUSED`, `COMMENT_IN_HEADER_FOOTER`, `PART_NOT_FOUND`, and `PART_AMBIGUOUS`. `artifactsChanged` may include comment sibling parts, content types, and header/footer parts. Limitations: no header/footer creation or removal, no header/footer image/text-box edits, and no field protection for body edits.

### WP1 — Dependency Upgrade & Compatibility Suite (`v0.8.1`)
1. In `package.json`, set `"@ansonlai/docx-redline-js": "0.8.1"`.
2. In `mcp/docx-server/package.json`:
   - Set `"@ansonlai/docx-redline-js": "0.8.1"`.
   - Retain `"jszip"` and `"@xmldom/xmldom"` until WP5 removes their imports.
3. Run `npm install` in root and `npm install --prefix mcp/docx-server`.
4. Create `tests/docx_redline_v081_compat_tests.mjs`:
   - Verify `openDocx` ingests and outputs canonical `Uint8Array`.
   - Verify `doc.inspect()` extracts 1-based paragraph index, `P#` reference, `paragraphId`, and exact text.
   - Verify order-independent batch execution with `atomic: true`.
   - Verify localized replacements (`replacements: [{ find, replace }]`).
   - Verify native revision resolution via `doc.resolveRevisions('accept')`.
   - Verify structured error pass-through (`COMMENTED_CONTENT_DELETE`, `GENERATED_OOXML_INVALID`, `PATCH_ROUNDTRIP_MISMATCH`).
   - Verify a body edit in a document with existing comments keeps untouched comment parts and relationships byte-identical and writes the correct `commentsExtended.xml` content type. Check repair of the old type and rejection by `validateDocxPackage`.
   - Verify reply range/reference markers and anchor preservation after a middle-of-paragraph edit. Cover Word-authored multi-paragraph threads and optional comment sibling parts in WP6 fixtures.
   - Verify `comment_resolve`, `resolveComment`, reopen, delete-by-id, root cascade, `no_change`, and `COMMENT_NOT_FOUND` rollback.
   - Verify `inspect().headersFooters`, an opt-in footer edit, body/part batch ordering and atomic rollback, package-wide revision ID uniqueness, and acceptance of revisions in parts. Verify body searches cannot match part text; cover rejection with Word-authored fixtures in WP6.
   - Verify `FIELD_EDIT_REFUSED`, `COMMENT_IN_HEADER_FOOTER`, `PART_NOT_FOUND`, `PART_AMBIGUOUS`, and `PARENT_ANCHOR_NOT_FOUND` pass through the consumer boundary without fallback.
   - Verify `artifactsChanged` reflects changed comment, content-type, and header/footer parts.
5. Run the v0.5.4 compatibility suite with its consumer guards and record any behavior tests that intentionally skip after the version change.

**WP1 progress (2026-09-29):** Both manifests and lockfiles now pin exact `0.8.1`, and both installed package copies report `0.8.1`. Existing root consumer/adapter/bridge tests, MCP service tests, table normalization and prompt tests, and `npm run build:dev` passed. The v0.5.4 compatibility suite passed its consumer guards and skipped package behavior by design because the installed version is now 0.8.1. `tests/phase4/golden-guardrail.mjs` reported changed list, table, and comment XML hashes against the v0.5.4 baseline. Its semantic flags and counts still match. The comment-package size increase is the new `w14:paraId` on comment paragraphs; the list/table deltas span the whole v0.5.4-to-v0.8.1 jump and need XML-level review before updating goldens. The generated latest fixture was restored to its previous tracked state. The new `tests/docx_redline_v081_compat_tests.mjs` passes. It covers package opening, inspection, localized atomic batches, comment content types and preservation, thread operations, header/footer editing and refusal, rollback, revision IDs, and artifact reporting. `npm ls @ansonlai/docx-redline-js --depth=0` confirms exact `0.8.1` in both projects. Word-authored thread fixtures, optional sibling parts, and full golden updates remain WP6 quality gates.

### WP2 — Purge Obsolete Workarounds & Dead Code
1. Delete the unused `src/taskpane/modules/docx-redline-js-integration/word-structured-list.js` and its add-in export.
2. Remove the direct `ReconciliationPipeline` structural fallback and unused `detectRequestedContentKind` table-intent heuristic from `agentic-tools.js`; use the supported package list helpers for list generation.
3. Guard against restoring the deleted module in `tests/no_legacy_shared_operation_bridge_tests.mjs` and verify list generation through the package API.

**WP2 record (2026-09-29):** The unused structured-list module, direct reconciliation fallback, and table-intent heuristic were removed. `tests/agentic_list_generation_tests.mjs` covers the supported list path. Investigation found that `normalizeListItemsWithLevels` and `buildListMarkdown` are supported v0.8.1 exports still used by active commands, and `list-level-utils.js` is an active list-level clamp helper, so these remain. The older `word-operation-runner.js` paragraph/scope helpers also remain because other agentic commands and adapter tests still call them. They are outside the WP3 redline batch path. Remove them only after those callers migrate.

### WP3 — Refactor Word Add-In to True Single-Hop Batch Runner
1. In `src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js`:
   - Implement `executePureOoxmlBatch(context, targetScope, operations, options)` accepting the full operations array.
   - Forward operations directly to `applyOperationsToDocumentXml` (or `doc.applyOperations`) with `atomic: true`, `structuredContent: true`, and `pairReplacements: true`.
   - Support both full `content` and `replacements: [{ find, replace }]`.
2. In `src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js`:
   - Replace the 1,100-line iterative `Paragraph` proxy walking loop with a direct invocation of `executePureOoxmlBatch`.
   - Remove manual table synthesis and range-expansion heuristics.

**WP3 record (2026-09-29):** `executePureOoxmlBatch` now reads scope OOXML once, inspects the immutable source, submits one atomic structured batch, preserves existing package parts, and writes once only when the batch succeeds. The redline caller plans operations from the inspected paragraph descriptors and uses package-localized replacements for `modify_text`; full paragraph and range edits remain supported. The iterative proxy loop, native insertion recovery, and manual table synthesis were removed. `tests/addin/word_pure_ooxml_batch_tests.mjs` covers the bridge, and `tests/word_redline_batch_plan_tests.mjs` covers planning, rollback, range/append behavior, and an end-to-end one-read/one-write mock Word flow. Word Desktop package insertion and commented-document opening remain WP6 manual gates.

### WP4 — Modernize Prompting & Schema for Localized Replacements
1. In `src/taskpane/modules/commands/redline-prompt.js`:
   - Update `REDLINE_DIFF_SCHEMA` to allow `replacements: [{ find, replace }]`.
   - Add prompt instructions: *"For word, phrase, or sentence edits inside an existing paragraph, return `replacements: [{ find, replace }]` instead of rewriting the entire paragraph."*
2. In `src/taskpane/modules/commands/change-validation.js`:
   - Update `verifyAnchor` and `sanitizeChangeSet` to validate localized replacements against target paragraph text.

### WP5 — Modernize `mcp/docx-server`
1. Remove direct `jszip` and `@xmldom/xmldom` dependencies only after replacing the services that import them.
2. In `mcp/docx-server/src/server.mjs`:
   - Refactor to use `openDocx(fs.readFileSync(path))` directly.
   - Wire `docx_list_paragraphs` to `doc.inspect().paragraphs`.
   - Wire `docx_edit_paragraph` and batch edits to `doc.applyOperations()`.
   - Wire `docx_save` to write `doc.toUint8Array()`.
3. Delete redundant services:
   - `mcp/docx-server/src/services/docx-package-service.mjs`
   - `mcp/docx-server/src/services/paragraph-targeting-service.mjs`

### WP6 — Verification & Quality Gates
1. Create `tests/ooxml_formatting_visual_tests.mjs` (verifying bold/italic inheritance, tabs, line breaks, hyperlinks, section layout, and localized replacements). Include Word-authored multi-paragraph comment threads, optional `commentsIds.xml`/`commentsExtensible.xml` synchronization, and Word Desktop open-without-repair checks for commented documents. Review the v0.5.4 golden XML differences before updating hashes.
2. Create `scripts/benchmark-ooxml-pipeline.mjs` to empirically verify sub-100ms batch processing.
3. Create `scripts/run-all-tests.mjs` and configure `"test": "node scripts/run-all-tests.mjs"`.

---

## 4. Definition of Done

- The v0.8.1 release delta is recorded and relevant changes are covered by compatibility tests.
- Both manifests and lockfiles resolve exact `0.8.1`.
- External `jszip` and `@xmldom/xmldom` are removed from `mcp/docx-server` in WP5 after the imports are replaced.
- The unused `word-structured-list.js` module and WP5 MCP services are deleted. Active list-level helpers remain until their callers no longer need them.
- Single-hop batch execution replaces the 1,100-line iterative proxy loop in `word-redline-runner.js`.
- Localized replacements (`replacements: [{ find, replace }]`) are fully functional in prompt, validation, and engine execution.
- All test suites execute offline via `npm test` and pass 100%.
- Webpack builds cleanly for development and production without Node polyfill errors.
