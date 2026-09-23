# `@ansonlai/docx-redline-js` v0.8.0 Direct Upgrade Plan

**Date:** 2026-09-23  
**Status:** Active — Supersedes the intermediate v0.5.4 plan; direct migration to v0.8.0.  
**Target:** Upgrade both root add-in and `mcp/docx-server` directly to exact version `0.8.0`.

This plan supersedes and consolidates:
- `2026-09-02-docx-redline-js-v0.4.0-upgrade.md`
- `2026-09-07-docx-redline-js-v0.5.0-upgrade.md`
- `2026-09-09-docx-redline-js-v0.5.4-upgrade.md`

Because the four 2026-08-29 optimization plans were never executed against v0.5.4, an intermediate step to 0.5.4 is unnecessary. Jumping directly to **v0.8.0** eliminates the need for temporary glue code (such as manual XML wrapping, synthetic `<w:document>` envelopes, and external JSZip handling) that 0.8.0 solves natively.

---

## 1. Summary of Major Library Capabilities (v0.5.0 &rarr; v0.8.0)

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

---

## 2. Migration Principles for Direct v0.8.0 Upgrade

1. **Pin `0.8.0` exactly** in both `package.json` and `mcp/docx-server/package.json`. Regenerate lockfiles via `npm install`.
2. **Eliminate Redundant Dependencies**: Remove `jszip` and `@xmldom/xmldom` from `mcp/docx-server/package.json`.
3. **Adopt Universal `openDocx` & Lifecycle Methods**:
   - In MCP: Use `openDocx(fs.readFileSync(path))` &rarr; `doc.inspect()` &rarr; `doc.applyOperations()` &rarr; `doc.toUint8Array()`.
   - In Word Add-in: Support both Document-Level Binary Single-Hop (`getFileAsync` &rarr; `openDocx` &rarr; `insertFileFromBase64`) and Range-Level OOXML Single-Hop (`applyOperationsToDocumentXml`).
4. **Transition to Order-Independent Batches**: Eliminate the 1,100-line iterative `for (const change of changes)` loop in `word-redline-runner.js`. Send operations as a single batch with `atomic: true`.
5. **Modernize Prompt & Validation for Localized Replacements**:
   - Update `redline-prompt.js` to encourage `replacements: [{ find, replace }]` for phrase/sentence edits.
   - Update `change-validation.js` to validate `find` targets within paragraph text rather than demanding full paragraph rewrites.
6. **Preserve Error Codes Generically**: Ensure structured error codes (`COMMENTED_CONTENT_DELETE`, `FOREIGN_PARAGRAPH_MARK_DELETION`, `GENERATED_OOXML_INVALID`, `PATCH_ROUNDTRIP_MISMATCH`, `UNSAFE_REVISION_BOUNDARY`, `TARGET_NOT_FOUND`, `AMBIGUOUS_ANCHOR`) pass cleanly through all host boundaries.
7. **Purge Bespoke Workarounds**: Delete `word-structured-list.js`, `list-level-utils.js`, and manual table recovery heuristics.

---

## 3. Work Packages

### WP0 — Establish Current Working Baseline
1. Check `git status --short`.
2. Run existing suites:
   ```bash
   node tests/redline_result_contract_tests.mjs
   node tests/addin/word_operation_runner_adapter_tests.mjs
   node tests/addin/shared_operation_bridge_tests.mjs
   node tests/no_legacy_shared_operation_bridge_tests.mjs
   npm run build:dev
   ```
3. Record clean baseline before modifying dependencies.

### WP1 — Dependency Upgrade & Compatibility Suite (`v0.8.0`)
1. In `package.json`, set `"@ansonlai/docx-redline-js": "0.8.0"`.
2. In `mcp/docx-server/package.json`:
   - Set `"@ansonlai/docx-redline-js": "0.8.0"`.
   - Remove `"jszip"` and `"@xmldom/xmldom"`.
3. Run `npm install` in root and `npm install --prefix mcp/docx-server`.
4. Create `tests/docx_redline_v080_compat_tests.mjs`:
   - Verify `openDocx` ingests and outputs canonical `Uint8Array`.
   - Verify `doc.inspect()` extracts 1-based paragraph index, `P#` reference, `paragraphId`, and exact text.
   - Verify order-independent batch execution with `atomic: true`.
   - Verify localized replacements (`replacements: [{ find, replace }]`).
   - Verify native revision resolution via `doc.resolveRevisions('accept')`.
   - Verify structured error pass-through (`COMMENTED_CONTENT_DELETE`, `GENERATED_OOXML_INVALID`, `PATCH_ROUNDTRIP_MISMATCH`).

### WP2 — Purge Obsolete Workarounds & Dead Code
1. **Delete**:
   - `src/taskpane/modules/docx-redline-js-integration/word-structured-list.js` (176 lines of hardcoded OPC XML).
   - `src/taskpane/modules/commands/list-level-utils.js` (regex marker parsing).
2. **Clean `agentic-tools.js`**:
   - Remove `detectRequestedContentKind`.
   - Remove `normalizeListItemsWithLevels` and `buildListMarkdown`.
   - Remove direct `ReconciliationPipeline` references.
3. **Clean `word-operation-runner.js`**:
   - Remove `enforceListBindingOnParagraphNodes`.
   - Remove `wrapParagraphNodesAsDocument`, `buildParagraphOnlyPackage`, and `wrapParagraphWithComments`.
4. Update `tests/no_legacy_shared_operation_bridge_tests.mjs` resurrection guards to protect against reintroducing deleted files.

### WP3 — Refactor Word Add-In to True Single-Hop Batch Runner
1. In `src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js`:
   - Implement `executePureOoxmlBatch(context, targetScope, operations, options)` accepting the full operations array.
   - Forward operations directly to `applyOperationsToDocumentXml` (or `doc.applyOperations`) with `atomic: true`, `structuredContent: true`, and `pairReplacements: true`.
   - Support both full `content` and `replacements: [{ find, replace }]`.
2. In `src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js`:
   - Replace the 1,100-line iterative `Paragraph` proxy walking loop with a direct invocation of `executePureOoxmlBatch`.
   - Remove manual table synthesis and range-expansion heuristics.

### WP4 — Modernize Prompting & Schema for Localized Replacements
1. In `src/taskpane/modules/commands/redline-prompt.js`:
   - Update `REDLINE_DIFF_SCHEMA` to allow `replacements: [{ find, replace }]`.
   - Add prompt instructions: *"For word, phrase, or sentence edits inside an existing paragraph, return `replacements: [{ find, replace }]` instead of rewriting the entire paragraph."*
2. In `src/taskpane/modules/commands/change-validation.js`:
   - Update `verifyAnchor` and `sanitizeChangeSet` to validate localized replacements against target paragraph text.

### WP5 — Modernize `mcp/docx-server`
1. In `mcp/docx-server/src/server.mjs`:
   - Refactor to use `openDocx(fs.readFileSync(path))` directly.
   - Wire `docx_list_paragraphs` to `doc.inspect().paragraphs`.
   - Wire `docx_edit_paragraph` and batch edits to `doc.applyOperations()`.
   - Wire `docx_save` to write `doc.toUint8Array()`.
2. Delete redundant services:
   - `mcp/docx-server/src/services/docx-package-service.mjs`
   - `mcp/docx-server/src/services/paragraph-targeting-service.mjs`

### WP6 — Verification & Quality Gates
1. Create `tests/ooxml_formatting_visual_tests.mjs` (verifying bold/italic inheritance, tabs, line breaks, hyperlinks, section layout, and localized replacements).
2. Create `scripts/benchmark-ooxml-pipeline.mjs` to empirically verify sub-100ms batch processing.
3. Create `scripts/run-all-tests.mjs` and configure `"test": "node scripts/run-all-tests.mjs"`.

---

## 4. Definition of Done

- Both `package.json` and `mcp/docx-server/package.json` resolve exact `0.8.0`.
- External `jszip` and `@xmldom/xmldom` are completely removed from `mcp/docx-server`.
- All obsolete files (`word-structured-list.js`, `list-level-utils.js`, `docx-package-service.mjs`, `paragraph-targeting-service.mjs`) are deleted.
- Single-hop batch execution replaces the 1,100-line iterative proxy loop in `word-redline-runner.js`.
- Localized replacements (`replacements: [{ find, replace }]`) are fully functional in prompt, validation, and engine execution.
- All test suites execute offline via `npm test` and pass 100%.
- Webpack builds cleanly for development and production without Node polyfill errors.
