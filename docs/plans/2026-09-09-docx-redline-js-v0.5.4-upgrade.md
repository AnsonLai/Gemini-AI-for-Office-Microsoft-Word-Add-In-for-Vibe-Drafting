# `@ansonlai/docx-redline-js` v0.8.2 Direct Upgrade Plan

**Date:** 2026-09-23  
**Updated:** 2026-09-30

**Status:** **Complete.** WP0–WP6 are implemented and verified on exact v0.8.2. Actual Office.js transport passed on 2026-09-30 during reliability WP1. PDF export is optional; measured large-document performance is accepted by the user. See the [v0.8.2 report](../validation-reports/2026-09-30-docx-redline-js-v0.8.2.md) and [host closure evidence](../validation-reports/2026-09-30-reliability-and-quality-gates.md).

**Target:** Upgrade both root add-in and `mcp/docx-server` directly to exact version `0.8.2`.

## Completion snapshot (2026-09-30)

| Work package | State |
| --- | --- |
| WP0: baseline | Complete |
| WP1: pins and compatibility | Complete; both consumers use exact 0.8.2 |
| WP2: obsolete workaround cleanup | Complete; active supported helpers retained |
| WP3: atomic Word redline batch runner | Complete for range/body Flat-OPC transport |
| WP4: localized prompt/schema validation | Complete |
| WP5: MCP document facade | Complete |
| WP6: verification | Complete: upgrade checks pass; reliability follow-up adds seven actual Office.js checks and 12 independent Word checks |

### Final closure record

Completed on Word `16.0.20326.20158`: a development-only validation entry invokes the production bridge with genuine `getOoxml` / `insertOoxml` calls for a localized body edit and threaded-comment reply. Seven transport checks verify confirmed insertion and no writes for empty batches, unchanged paragraph and invalid targets. Twelve independent Word reopen/no-repair/Accept All/Reject All checks pass on exported packages and reference files, preserving exact text and comment identities. See the reliability report for commands, snapshots and artifact hashes.

This check closed [reliability WP1](2026-08-29-reliability-and-quality-gates.md) and upgrade WP6 together. Upgrade changes were committed as `68da8a4`; reliability follow-up changes form a separate change set.

### Follow-on plan order

1. [Reliability and quality gates](2026-08-29-reliability-and-quality-gates.md): complete; host gate, provider recovery and mutation outcomes verified.
2. [Agentic tools and list reliability](2026-08-29-agentic-tools-and-list-reliability.md): align remaining tool/list paths and expand independent Word coverage.
3. [Package boundaries and integrations](2026-08-29-package-boundaries-and-integrations.md): extract portable consumer logic and migrate browser-demo editing; MCP migration is already complete.
4. [OOXML engine and performance](2026-08-29-oxml-engine-and-performance.md): profile settled paths, then optimize measured consumer costs; route library findings upstream.

The August plans now distinguish inherited completed work from remaining work. They do not require repeating this upgrade.

This plan supersedes and consolidates:
- `2026-09-02-docx-redline-js-v0.4.0-upgrade.md`
- `2026-09-07-docx-redline-js-v0.5.0-upgrade.md`

Because the four 2026-08-29 optimization plans were never executed against v0.5.4, an intermediate step through v0.8.0 is unnecessary. Upgrade directly to **v0.8.2** and use the document lifecycle introduced in v0.8.0 to retire temporary glue code (such as manual XML wrapping, synthetic `<w:document>` envelopes, and external JSZip handling). Verify those APIs and behaviors against v0.8.2 before removing existing code. The records below preserve the initial v0.8.1 implementation and its subsequent v0.8.2 verification.

---

## 1. Summary of Major Library Capabilities (v0.5.0 &rarr; v0.8.2)

| Release Span | Architectural Breakthrough | Direct Impact on This Repository |
|---|---|---|
| **v0.5.1 – v0.5.4** | Multi-Author Revision Slicing (`slice-cross-author`), Paragraph Restoration (`restore`), and Baseline-Delta Validation. | Pre-existing document defects or existing tracked changes in user files no longer fail subsequent valid operations. |
| **v0.6.0 – v0.6.2** | **Caller-Order-Independent Batches**: Targets resolve against the immutable starting state of the document. | Prior splits or insertions in a batch no longer throw off later paragraph indexes, eliminating the need for manual bottom-up sorting or iterative single-paragraph loops in `word-redline-runner.js`. |
| **v0.6.0 – v0.6.2** | Compact Mutation Outputs & Receipts (`completion`, `written`, `status`, `receipts`). | Clean structured receipts allow caller verification without inspecting verbose OOXML dumps. |
| **v0.7.0 – v0.7.2** | **Localized Exact Replacements (`replacements: [{ find, replace }]`)**: Operations specify exact phrase substitutions inside a target paragraph. | Avoids repeating unchanged paragraph text; token savings depend on the workload. Exact target and occurrence validation remain required. |
| **v0.7.0 – v0.7.2** | Two-Turn Golden Path: Single computed success gate (`completion: true` + non-null output). | Streamlines machine recovery and verification. |
| **v0.8.0** | **Universal `openDocx` Facade & Zero-Node Architecture**: Runs natively in modern browsers without `node:zlib`, `node:crypto`, or `Buffer`. Ingests and serializes canonical `Uint8Array`. | Zero-Node packaging runs out-of-the-box in the Word taskpane and browser demo. |
| **v0.8.0** | **Built-in `fflate` ZIP Container**: ZIP extraction, OOXML part loading, and rebuilding are handled internally. | **Removes external `jszip` and `@xmldom/xmldom` dependencies** from `mcp/docx-server`. |
| **v0.8.0** | **Integrated Document Lifecycle**: `doc.inspect()`, `doc.applyOperations()`, and `doc.resolveRevisions('accept' \| 'reject')`. | Extracts paragraphs, applies batches, manages `numbering.xml`/`comments.xml`, and resolves revisions in one step. |
| **v0.8.1** | Fixes the `commentsExtended.xml` content type and comment-thread markers; adds comment resolve/reopen/delete-by-id and editing of existing header/footer parts. | Prevents Word repair prompts when editing commented files; requires regression coverage for comment preservation, thread operations, and body/part batch isolation. |
| **v0.8.2** | Fixes hyperlink/plain-run boundary preservation and Reject All text ordering around manual line breaks. CLI contract 8, APIs, schemas and error shapes remain unchanged. | Resolves library issues #3/#4; localized and full-paragraph forms now run as required offline and real Word regressions. |

---

## 2. Migration Principles for Direct v0.8.2 Upgrade

1. **Pin `0.8.2` exactly** in both `package.json` and `mcp/docx-server/package.json`. Regenerate lockfiles via `npm install`.
2. **Eliminate Redundant Dependencies in WP5**: Keep `jszip` and `@xmldom/xmldom` until the MCP server stops importing them. Remove them with the WP5 facade refactor.
3. **Adopt Universal `openDocx` & Lifecycle Methods**:
   - In MCP: Use `openDocx(fs.readFileSync(path))` &rarr; `doc.inspect()` &rarr; `doc.applyOperations()` &rarr; `doc.toUint8Array()`.
   - In Word Add-in: Use Range/Body-Level OOXML Single-Hop (`applyOperationsToDocumentXml`) with Flat-OPC part preservation. Document-Level Binary transport (`getFileAsync` &rarr; `openDocx` &rarr; `insertFileFromBase64`) remains a future option in the package-boundaries plan; it is not implemented or required to close this upgrade.
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

### WP1 — Dependency Upgrade & Compatibility Suite (`v0.8.2`, unchanged v0.8.1 API)
1. In `package.json`, set `"@ansonlai/docx-redline-js": "0.8.2"`.
2. In `mcp/docx-server/package.json`:
   - Set `"@ansonlai/docx-redline-js": "0.8.2"`.
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

**WP1 progress record (2026-09-29; historical checkpoint):** Both manifests and lockfiles then pinned exact `0.8.1`, and both installed package copies reported `0.8.1`. Existing root consumer/adapter/bridge tests, MCP service tests, table normalization and prompt tests, and `npm run build:dev` passed. The v0.5.4 compatibility suite passed its consumer guards and skipped package behavior by design because the installed version was now 0.8.1. `tests/phase4/golden-guardrail.mjs` reported changed list, table, and comment XML hashes against the v0.5.4 baseline. Its semantic flags and counts still matched. The comment-package size increase was the new `w14:paraId` on comment paragraphs; the list/table deltas spanned the whole v0.5.4-to-v0.8.1 jump. This checkpoint tentatively called for XML review before updating goldens; the later [golden provenance correction](#wp6--verification--quality-gates) records the review and explains why the old attribution was incorrect. The generated latest fixture was restored to its previous tracked state. The new `tests/docx_redline_v081_compat_tests.mjs` passed and covered package opening, inspection, localized atomic batches, comment content types and preservation, thread operations, header/footer editing and refusal, rollback, revision IDs, and artifact reporting. `npm ls @ansonlai/docx-redline-js --depth=0` confirmed exact `0.8.1` in both projects at that checkpoint. Word-authored thread fixtures, optional sibling parts, and golden review were then outstanding WP6 gates; the later WP6 records below close them.

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

**WP3 record (2026-09-29; historical checkpoint):** `executePureOoxmlBatch` reads scope OOXML once, inspects the immutable source, submits one atomic structured batch, preserves existing package parts, and writes once only when the batch succeeds. The redline caller plans operations from the inspected paragraph descriptors and uses package-localized replacements for `modify_text`; full paragraph and range edits remain supported. The iterative proxy loop, native insertion recovery, and manual table synthesis were removed. `tests/addin/word_pure_ooxml_batch_tests.mjs` covers the bridge, and `tests/word_redline_batch_plan_tests.mjs` covers planning, rollback, range/append behavior, and an end-to-end one-read/one-write mock Word flow. At this checkpoint, Word Desktop package insertion and commented-document opening were still WP6 gates; later WP6 and the final closure record document their completion.

### WP4 — Modernize Prompting & Schema for Localized Replacements
1. In `src/taskpane/modules/commands/redline-prompt.js`:
   - Update `REDLINE_DIFF_SCHEMA` to allow `replacements: [{ find, replace }]`.
   - Add prompt instructions: *"For word, phrase, or sentence edits inside an existing paragraph, return `replacements: [{ find, replace }]` instead of rewriting the entire paragraph."*
2. In `src/taskpane/modules/commands/change-validation.js`:
   - Update `verifyAnchor` and `sanitizeChangeSet` to validate localized replacements against target paragraph text.

**WP4 record (2026-09-29):** The prompt now prefers `edit_paragraph` with `replacements: [{ find, replace }]` for local text edits and retains `newContent` for full rewrites. The Gemini response schema exposes the replacement array and optional 1-based `occurrence`. Sanitization validates item shape, exact case-sensitive matches, ambiguous repeated finds, and occurrences when source paragraph text is available. Anchor verification checks the finds again after any unambiguous paragraph-index correction. The add-in and eval harness pass the same source paragraph texts to sanitization. Prompt, validation, and end-to-end batch tests pass.

### WP5 — Modernize `mcp/docx-server`
1. Remove direct `jszip` and `@xmldom/xmldom` dependencies only after replacing the services that import them.
2. In `mcp/docx-server/src/server.mjs` and its document service:
   - Open file bytes with `openDocx` and keep a `DocxDocument` in each session.
   - Wire `docx_list_paragraphs` to `doc.inspect().paragraphs`.
   - Wire `docx_edit_paragraph`, comments, and `docx_apply_operations` batches to atomic `doc.applyOperations()`.
   - Wire the existing `docx_save_as` tool to write `doc.toUint8Array()`.
   - Use a small blank DOCX template for `docx_new`, since v0.8.1 does not expose a new-document factory.
3. Delete redundant services:
   - `mcp/docx-server/src/services/docx-package-service.mjs`
   - `mcp/docx-server/src/services/paragraph-targeting-service.mjs`
   - `mcp/docx-server/src/services/docx-redline-js-service.mjs`
   - `mcp/docx-server/src/services/xml-utils.mjs`

**WP5 record (2026-09-29):** The MCP server now stores `DocxDocument` sessions and uses the v0.8.1 lifecycle for opening, inspection, edits, comments, and saving. It exposes `docx_apply_operations` for atomic batches and returns structured engine errors. The four redundant services and direct `jszip`/`@xmldom/xmldom` dependencies were removed from its manifest and lockfile. A 2,022-byte blank template supports `docx_new`; a clean initial title is applied through the package. The stdio workflow test and both root MCP tests pass, including create, edit, batch, comment, rollback, save, and reopen.

### WP6 — Verification & Quality Gates
1. Create `tests/ooxml_formatting_visual_tests.mjs` (verifying bold/italic inheritance, tabs, line breaks, hyperlinks, section layout, and localized replacements). Include Word-authored multi-paragraph comment threads, optional `commentsIds.xml`/`commentsExtensible.xml` synchronization, and Word Desktop open-without-repair checks for commented documents. Review the v0.5.4 golden XML differences before updating hashes.
2. Create `scripts/benchmark-ooxml-pipeline.mjs` to measure median/p95 for explicit workloads, separating core batches from open/save and host latency. Treat sub-100ms as an optional workload-specific budget, not a universal claim.
3. Create `scripts/run-all-tests.mjs` and configure `"test": "node scripts/run-all-tests.mjs"`.

**WP6 record (2026-09-30; then-current 33-suite checkpoint):** The offline runner passed all 33 consumer suites and reported four excluded entrypoints plus the historical v0.5.4 package-behavior skip. Development and production builds passed; production retained bundle-size warnings. The fidelity suite covered exact accepted/rejected text, formatting and structure preservation, three Word-authored fixtures, thread lifecycle, optional sibling identities, a localized first-page footer edit and atomic no-write refusal. The add-in bridge delegated sibling synchronization and package metadata to the library's public helper/registry; tests covered old content-type repair. Later verification checkpoints are recorded below.

**Golden provenance correction:** The earlier WP1 statement attributing list/table/comment deltas to the v0.5.4-to-v0.8.1 jump was incorrect. The stored hashes mixed older in-repository engine output with a partial v0.5.4 refresh. Replaying all nine exact cases against upstream v0.5.4 tag commit `2871bca73d590bd626518d9c419aff81254a4219` produces raw XML identical to v0.8.1. Historical XML differences were reviewed (paragraph-mark tracking, table header/row properties, comment paragraph IDs) before refreshing baseline/latest. Golden `--verify` no longer rewrites latest output.

**Word evidence (expanded; historical COM checkpoint):** Word 16.0/build 16.0.20326 passed 45 live checks across six supported fixtures, two native insertion roundtrips and the invalid-content-type negative control. `npm run test:word` reads Word's own `Content.WordOpenXML`, runs exact fixture operations through the production bridge, performs native insertion, saves and reopens, checks recognized revisions, exact Accept All/Reject All text, comment parent identities/resolved state and first-page footer text. Engine-resolved files are checked for residual revisions before Word resolves anything. The earlier insertion timeout attribution was wrong: insertion returned and `SaveAs2` stalled. Boxed by-reference save arguments, as used upstream, fix the harness. PDF export is disabled by default and optional via `-Render`; it is not an upgrade feature or required gate. Office.js transport had not yet been tested at this checkpoint; the later closure record adds direct Office.js evidence.

**Performance:** Ten localized edits measured a 35.610 ms core median (39.126 ms p95) at 100 paragraphs and 296.499 ms (393.545 ms p95) at 1,000 paragraphs. The user accepts the larger workload's slower performance; it does not block the upgrade. Raw samples are retained in `scripts/ooxml-benchmark-latest.json`.

**Additional live structural checks:** The formatting fixture checks Word's bold/italic state, plain-run isolation, Calibri/12-point font inheritance, explicit right tab stop, portrait/landscape sections and column counts. Word's additional default tab stops are distinguished from the explicit stop. These assertions complement XML preservation checks without claiming rendered page appearance.

**v0.8.2 follow-up (2026-09-30; then-current 33-suite checkpoint):** Both manifests, lockfiles and installed copies resolved exact `0.8.2`. Library issues [#3](https://github.com/AnsonLai/docx-redline-js/issues/3) and [#4](https://github.com/AnsonLai/docx-redline-js/issues/4) were fixed upstream, with no consumer-side workaround. Their localized and full-paragraph forms were mandatory default regressions. At this checkpoint, all 33 offline suites, the portable reproducer, 69 live Word checks and both webpack builds passed. Word verified exact accepted/rejected text, hyperlink boundaries and recognized revisions for all four fixed cases; native plain-edit and threaded-comment roundtrips still passed. Existing golden hashes were unchanged. Historical v0.8.1 evidence remains in its separate report.

**WP6 closed (2026-09-30):** Actual Office.js transport and its exported Word package lifecycle pass in reliability WP1. No upgrade acceptance tasks remain. The earlier COM evidence and v0.8.2 report above are historical records; the closure report adds direct Office.js evidence. PDF export is optional and accepted performance does not block the upgrade.

### Current repository checkpoint (2026-09-30, `96b6d92`)

The September upgrade remains closed at exact `@ansonlai/docx-redline-js@0.8.2`: both root and MCP manifests and lockfiles pin that version. The later agentic-tools/list-reliability checkpoint records **44 offline suites passed, zero failed**, with four excluded entrypoints and the historical v0.5.4 package-behavior skip reported separately; see its [validation report](../validation-reports/2026-09-30-agentic-tools-and-list-reliability.md) and [follow-on plan](2026-08-29-agentic-tools-and-list-reliability.md). That follow-on remains in progress: broader list migrations and deeper/other-style/tracking-off coverage are open, and separately reported library fidelity/capability findings remain tracked there. Those findings are outside this upgrade's acceptance criteria and do not reopen or block the v0.8.2 upgrade.

---

## 4. Definition of Done

- [x] The v0.8.1/v0.8.2 release deltas are recorded and covered by compatibility/fidelity tests.
- [x] Both manifests and lockfiles resolve exact `0.8.2`.
- [x] Direct `jszip` and `@xmldom/xmldom` dependencies removed from MCP after import migration.
- [x] Unused structured-list module and obsolete MCP services deleted; active helpers retained.
- [x] Atomic single-hop batching replaces the iterative redline proxy loop.
- [x] Localized replacements work in prompting, validation and execution.
- [x] Offline consumer suites pass; exclusions/version-specific skips reported explicitly.
- [x] Word-authored no-repair/revision and native COM insertion checks pass; both library defects fixed upstream without consumer workarounds.
- [x] Actual Office.js single-hop insertion verified and closure evidence recorded.
- [x] Development/production builds pass without Node polyfill errors; existing production bundle warnings recorded.

PDF export is optional diagnostic evidence. The measured 1,000-paragraph performance is accepted. Live model evaluation and broader list/browser/portability work belong to the August follow-on plans.
