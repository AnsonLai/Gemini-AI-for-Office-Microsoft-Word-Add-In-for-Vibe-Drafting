# OXML Engine Refactor Plan (Phase 4: Complexity and Line Reduction)

Date: 2026-02-07

## Scope

Behavior-preserving refactor of the OOXML reconciliation stack with two priorities:

1. Reduce line count and cognitive complexity in core modules.
2. Capture safe performance gains where hotspots are currently O(n*m) or repeatedly parse/scan XML.

No functional changes are intended in this phase.

## Execution Status (Updated 2026-02-07)

- [x] P4.1 Add Guardrails Before Refactor
- [x] P4.2 Slim Router Further (`engine/oxml-engine.js`)
- [x] P4.3 Split Comment Engine by Responsibility
- [x] P4.4 Decompose Ingestion into Paragraph/Table Pipelines
- [x] P4.5 Refactor Reconstruction Mode Into Mapper + Writer
- [x] P4.6 Simplify Patching Hot Paths
- [x] P4.7 Pipeline Cleanup and Stub Resolution

## Running Notes

- P4.1 implementation completed:
  - Added golden guardrail harness: `tests/phase4/golden-guardrail.mjs`
  - Added perf harness with threshold-based regression checks: `tests/phase4/perf-harness.mjs`
  - Captured and committed baselines:
    - `tests/fixtures/phase4-golden-baseline.json`
    - `tests/fixtures/phase4-perf-baseline.json`
  - Added latest-run outputs (diagnostic artifacts):
    - `tests/fixtures/phase4-golden-latest.json`
    - `tests/fixtures/phase4-perf-latest.json`

- P4.2 implementation completed:
  - Extracted table orchestration from router into new module:
    - `src/taskpane/modules/reconciliation/engine/table-mode.js`
  - Slimmed `src/taskpane/modules/reconciliation/engine/oxml-engine.js`:
    - removed in-file table reconciliation/transformation functions
    - delegated table paths to `table-mode.js`
    - added shared no-change helper in router path
  - Performance action applied:
    - parser instance from `applyRedlineToOxml` is reused in table mode paths to reduce parser churn.

- P4.3 implementation completed:
  - Split comment responsibilities into focused modules:
    - `src/taskpane/modules/reconciliation/services/comment-builders.js`
    - `src/taskpane/modules/reconciliation/services/comment-locator.js`
    - `src/taskpane/modules/reconciliation/services/comment-package.js`
  - Rebuilt `src/taskpane/modules/reconciliation/services/comment-engine.js` as thin orchestration facade while preserving public exports.
  - Performance actions applied:
    - paragraph text indexing introduced for comment placement
    - lazy index creation + conditional index rebuild only when additional comments target the same paragraph.

- P4.4 implementation completed:
  - Split ingestion into dedicated modules:
    - `src/taskpane/modules/reconciliation/pipeline/ingestion-paragraph.js`
    - `src/taskpane/modules/reconciliation/pipeline/ingestion-table.js`
    - `src/taskpane/modules/reconciliation/pipeline/ingestion-xml.js`
  - Replaced `src/taskpane/modules/reconciliation/pipeline/ingestion.js` with a compatibility facade exporting the same public API.
  - Refactored paragraph ingestion recursive branching into a node-handler map in `ingestion-paragraph.js`.
  - Performance action applied:
    - table-cell block ingestion now uses `ingestParagraphElement(...)` on existing nodes, removing per-cell serialize+reparse cycles.

- P4.5 implementation completed:
  - Split reconstruction into:
    - `src/taskpane/modules/reconciliation/engine/reconstruction-mapper.js`
    - `src/taskpane/modules/reconciliation/engine/reconstruction-writer.js`
  - Replaced `src/taskpane/modules/reconciliation/engine/reconstruction-mode.js` with a thin orchestration layer.
  - Performance actions applied:
    - introduced cursor-based range lookups for paragraph/property maps in mapper
    - introduced sentinel index map (`start -> sentinels[]`) used by writer hot path
    - reduced repeated full-array scans in reconstruction write loop.

- P4.6 implementation completed:
  - Reworked `src/taskpane/modules/reconciliation/pipeline/patching.js` hot paths:
    - `splitRunsAtDiffBoundaries(...)` now uses pre-sorted unique boundaries + cursor advancement instead of per-run filtering/sorting
    - added `buildTextRunLookup(...)` with binary-search neighbors (`findRunBefore/findRunAfter`) for insertion style inheritance
    - added `createRangeCursorLookup(...)` for non-insert op coverage checks
    - extracted insertion complexity into focused helpers:
      - `processInsertionOperation(...)`
      - `parseInsertionLine(...)`
    - tracked last paragraph-start index during patching to avoid reverse scans when converting inserted lines to list items
  - Performance actions applied:
    - replaced repeated O(n) scans in patch loops with indexed/cursor lookups aligned to reconstruction strategy.

- P4.7 implementation completed:
  - Extracted list-generation flow into:
    - `src/taskpane/modules/reconciliation/pipeline/list-generation.js`
  - Added content-analysis helpers:
    - `src/taskpane/modules/reconciliation/pipeline/content-analysis.js`
    - implemented previously stubbed `detectContentType(...)` and `parseListItems(...)`
  - Simplified `src/taskpane/modules/reconciliation/pipeline/pipeline.js`:
    - delegates list generation and indentation detection to extracted module
    - re-exports content-analysis helpers from dedicated module
  - Performance actions applied:
    - list-generation path now computes and reuses per-line metadata once, avoiding repeated regex/parse passes across markdown lines.

- Verification runs after P4.1-P4.3:
  - `node tests/phase4/golden-guardrail.mjs` ✅
  - `node tests/phase4/perf-harness.mjs --verify` ✅
  - `node tests/comment_tests.mjs` ✅
  - `node tests/standalone_smoke.mjs` ✅
  - `node tests/no_word_api_standalone_check.mjs` ✅
  - `node tests/include_numbering_behavior.mjs` ✅
  - `node tests/table_tests.mjs` ✅
  - `node tests/list_tests.mjs` ✅
  - `node tests/integration_tests.mjs` ✅
  - `node tests/highlight_tests.mjs` ✅
  - `node tests/formatting_tests.mjs` runs with pre-existing FAIL messages (exit code 0) ⚠️

- Additional verification runs after P4.4-P4.5:
  - `node tests/phase4/golden-guardrail.mjs` ✅
  - `node tests/phase4/perf-harness.mjs --verify` ✅
  - `node tests/comment_tests.mjs` ✅
  - `node tests/table_tests.mjs` ✅
  - `node tests/list_tests.mjs` ✅
  - `node tests/integration_tests.mjs` ✅
  - `node tests/standalone_smoke.mjs` ✅
  - `node tests/include_numbering_behavior.mjs` ✅
  - `node tests/no_word_api_standalone_check.mjs` ✅
  - `node tests/highlight_tests.mjs` ✅
  - `node tests/formatting_tests.mjs` runs with pre-existing FAIL messages (exit code 0) ⚠️

- Additional verification runs after P4.6-P4.7:
  - `node tests/phase4/golden-guardrail.mjs` ✅
  - `node tests/phase4/perf-harness.mjs --verify` ✅
  - `node tests/comment_tests.mjs` ✅
  - `node tests/table_tests.mjs` ✅
  - `node tests/list_tests.mjs` ✅
  - `node tests/integration_tests.mjs` ✅
  - `node tests/standalone_smoke.mjs` ✅
  - `node tests/include_numbering_behavior.mjs` ✅
  - `node tests/no_word_api_standalone_check.mjs` ✅
  - `node tests/highlight_tests.mjs` ✅
  - `node tests/formatting_tests.mjs` runs with pre-existing FAIL messages (exit code 0) ⚠️

## Hotspot Snapshot (Current)

- `src/taskpane/modules/reconciliation/pipeline/pipeline.js` (~500 lines)
- `src/taskpane/modules/reconciliation/engine/surgical-mode.js` (~459 lines)
- `src/taskpane/modules/reconciliation/pipeline/ingestion-paragraph.js` (~347 lines)
- `src/taskpane/modules/reconciliation/engine/reconstruction-writer.js` (~276 lines)
- `src/taskpane/modules/reconciliation/engine/reconstruction-mapper.js` (~270 lines)
- `src/taskpane/modules/reconciliation/engine/oxml-engine.js` (~255 lines)
- `src/taskpane/modules/reconciliation/services/comment-engine.js` (~247 lines)

Historical baseline before P4.4/P4.5:

- `ingestion.js` ~585 lines -> split into facade + paragraph/table/xml modules
- `reconstruction-mode.js` ~485 lines -> split into orchestration + mapper + writer

## Target Outcomes

1. Reduce total lines across hotspot modules by 20-30% via decomposition + deduplication.
2. Reduce branching depth in `applyRedlineToOxml`, `injectCommentsIntoOoxml`, and `executeListGeneration`.
3. Improve runtime on large paragraphs/tables by reducing repeated scans and repeated XML parse work.
4. Keep public API and observable behavior unchanged.

## Phase 4 Workstreams

### P4.1 Add Guardrails Before Refactor (Low Risk)

- Add golden-output fixtures for representative scenarios:
  - format-only add/remove
  - mixed insert/delete with track changes
  - list generation (flat + nested)
  - table edit + text-to-table transform
  - paragraph and package comment injection
- Add a perf harness script for repeatable timings on:
  - long paragraph (5k+ chars, 200+ runs)
  - table (40x10 with mixed edits)
  - multi-comment injection (25+ comments)
- Record baseline timings before code movement.
- Define step-level perf budgets used by the remaining 6 steps:
  - no regression >5% on any baseline scenario
  - require at least one measurable win by the end of P4.7

### P4.2 Slim Router Further (`engine/oxml-engine.js`) (Low Risk)

- Extract table-specific orchestration into `engine/table-mode.js`:
  - `applyTableReconciliation`
  - `applyTextToTableTransformation`
- Keep `applyRedlineToOxml` as a pure decision router (target ~250 lines).
- Normalize repeated early-return serialization patterns into shared helpers (`success/no-change/parse-fail` helpers).
- Performance action:
  - consolidate parser/serializer lifecycle within router execution paths to avoid unnecessary object churn and duplicate parse branches

### P4.3 Split Comment Engine by Responsibility (Medium Risk)

- Split `services/comment-engine.js` into:
  - `services/comment-builders.js` (comment XML builders/markers)
  - `services/comment-locator.js` (text locating + run split logic)
  - `services/comment-package.js` (package part + rel updates)
  - keep `comment-engine.js` as thin orchestration facade
- Performance actions:
  - replace repeated subtree scans in text location/injection with indexed run-text maps built once per paragraph
  - add single-pass paragraph text indexing for multi-comment injection workloads
- Preserve current exports from `comment-engine.js` for compatibility.

### P4.4 Decompose Ingestion into Paragraph/Table Pipelines (Medium Risk)

- Split `pipeline/ingestion.js` into:
  - `pipeline/ingestion-paragraph.js` (run model ingestion)
  - `pipeline/ingestion-table.js` (virtual grid ingestion)
  - `pipeline/ingestion-xml.js` (shared XML/node helpers)
- Replace long `processNodeRecursive` conditional ladder with a node-handler map for `w:r`, `w:ins`, `w:del`, `w:hyperlink`, etc.
- Remove unused internal paths (`processHyperlink` helper if still unreferenced after split).
- Performance action:
  - avoid serializing + reparsing each table-cell paragraph by introducing node-based ingestion helpers for table paths

### P4.5 Refactor Reconstruction Mode Into Mapper + Writer (Higher Risk)

- Split `engine/reconstruction-mode.js` into:
  - `engine/reconstruction-mapper.js` (build property/sentinel/paragraph maps)
  - `engine/reconstruction-writer.js` (append/render diff segments)
- Performance actions (safe, behavior-preserving):
  - replace repeated `.find()` inside loops with cursor-based lookup over sorted ranges
  - pre-index sentinels by `start` offset (`Map<number, Sentinel[]>`)
  - cache paragraph/range metadata so append/write loops stay O(n) over diff segments
- Keep `applyReconstructionMode(...)` as public orchestrator only.

### P4.6 Simplify Patching Hot Paths (Medium Risk)

- In `pipeline/patching.js`:
  - pre-sort diff boundaries once; avoid re-filtering all boundaries per run
  - replace repeated `findRunBefore`/`findRunAfter` scans with precomputed nearest-text-run index
  - extract insertion line processing into a dedicated helper to reduce `applyPatches` complexity
- Keep existing patch semantics and list-marker behavior unchanged.
- Performance action:
  - add range-cursor lookup strategy aligned with P4.5 so patching and reconstruction share the same O(n) scanning model

### P4.7 Pipeline Cleanup and Stub Resolution (Low Risk)

- In `pipeline/pipeline.js`:
  - extract `executeListGeneration` into `pipeline/list-generation.js`
  - move markdown-table block detection/parsing helper out of method body
  - resolve `detectContentType` and `parseListItems` stubs:
    - either implement minimal real behavior, or internalize/remove exports if unused
- Performance actions:
  - reduce repeated markdown/list parsing by reusing parsed line metadata inside list generation
  - run full perf harness from P4.1 and fail the step if perf budgets are not met

## Performance Traceability (Recommendations to Steps)

1. Range cursor indexes in reconstruction/patching -> `P4.5`, `P4.6`
2. Single-pass paragraph text indexing in comment injection -> `P4.3`
3. Avoid re-parsing paragraph XML in table ingestion -> `P4.4`
4. Consolidate parser/serializer lifecycle where safe -> `P4.2`

## Validation Strategy

- Required suites after each workstream:
  - `tests/standalone_smoke.mjs`
  - `tests/no_word_api_standalone_check.mjs`
  - `tests/include_numbering_behavior.mjs`
  - `tests/comment_tests.mjs`
  - `tests/table_tests.mjs`
  - `tests/list_tests.mjs`
  - `tests/integration_tests.mjs`
  - `tests/highlight_tests.mjs`
- Run `tests/formatting_tests.mjs` and track existing known failures separately from regressions.
- Compare golden outputs from P4.1 before/after each phase.
- Perf harness pass criteria:
  - no regressions >5% in any measured scenario
  - target 10-25% improvement in at least one large-input scenario

## Sequencing

1. P4.1 Guardrails + baseline
2. P4.2 Router slimming
3. P4.3 Comment engine split
4. P4.4 Ingestion split
5. P4.6 Patching hot-path simplification
6. P4.5 Reconstruction mapper/writer split
7. P4.7 Pipeline cleanup/stub resolution

## Definition of Done

1. Planned modules are split with stable public exports.
2. Hotspot files reduced materially in size and branch depth.
3. Behavior remains equivalent across golden fixtures and existing tests.
4. Perf harness shows no material regressions and at least one clear gain.
