# OXML Engine Refactor Task List (Phase 1 First)

## Overall Status

✅ **Completed.** All Phase 1, Phase 2, and Phase 3 work items in this task
list are checked off.

## Status Legend

- [ ] Pending
- [~] In progress
- [x] Completed
- [!] Blocked

## Active Phase

Phase 1 (Minimal Risk)

## Phase 1 Work Items

- [x] P1.1 Create shared package-builder module for OOXML wrappers
  - Consolidate package construction now spread across:
    - `pipeline/serialization.js`
    - `engine/table-cell-context.js`
    - `services/comment-engine.js`
  - Define options for:
    - base document package
    - numbering part inclusion
    - comments part inclusion
    - paragraph-only packaging

- [x] P1.2 Migrate existing package wrapper call sites to package-builder
  - Replace direct template string package construction in each caller.
  - Keep output structure equivalent (no behavior change intended).

- [x] P1.3 Remove integration import cycle
  - Update `integration/integration.js` to import `ReconciliationPipeline` from `pipeline/pipeline.js` directly.
  - Verify no cycle via static import graph checks (`rg` + manual review).

- [x] P1.4 Create shared list marker parser utility
  - Move shared marker regex and parsing helpers into one module.
  - Replace duplicated logic in:
    - `engine/oxml-engine.js`
    - `pipeline/pipeline.js`
    - `pipeline/patching.js`

- [x] P1.5 Fix `includeNumbering` boolean handling in oxml-engine
  - Replace `result.includeNumbering || true` with nullish-safe handling.
  - Add regression coverage for `false`, `true`, and `undefined`.

- [x] P1.6 Phase 1 verification + docs update
  - Run and record:
    - `tests/standalone_smoke.mjs`
    - `tests/no_word_api_standalone_check.mjs`
    - impacted reconciliation tests (list/table/comment/formatting)
  - Update architecture/readme notes if module locations or responsibilities changed.

## Running Notes

- Added shared package module: `src/taskpane/modules/reconciliation/services/package-builder.js`.
- Added shared list marker utility: `src/taskpane/modules/reconciliation/pipeline/list-markers.js`.
- Migrated package construction call sites:
  - `pipeline/serialization.js`
  - `engine/table-cell-context.js`
  - `services/comment-engine.js`
- Broke barrel cycle by updating `integration/integration.js` to import pipeline directly.
- Fixed numbering inclusion fallback in `engine/oxml-engine.js` (`??` handling).
- Added regression test: `tests/include_numbering_behavior.mjs`.
- Added fallback smoke test: `tests/dom_fallback_smoke.mjs`.
- Added/updated fallback DOM setup: `tests/setup-xml-provider.mjs` now:
  - uses `jsdom` when available, otherwise `@xmldom/xmldom`
  - configures global DOM constructors and document
  - adds iterable NodeList fallback for `xmldom`
- Removed direct `jsdom` imports from tests and routed through shared fallback setup:
  - `tests/comment_tests.mjs`
  - `tests/formatting_tests.mjs`
  - `tests/integration_tests.mjs`
  - `tests/list_tests.mjs`
  - `tests/table_tests.mjs`
  - `tests/highlight_tests.mjs`
  - `tests/verify_fix.mjs`
  - `tests/debug_test6.mjs`
  - `tests/debug_extraction.mjs`
- Updated docs:
  - `src/taskpane/modules/reconciliation/ARCHITECTURE.md`
  - `src/taskpane/modules/reconciliation/README.md`
- Verification run results:
  - `node tests/include_numbering_behavior.mjs` ✅
  - `node tests/standalone_smoke.mjs` ✅
  - `node tests/no_word_api_standalone_check.mjs` ✅
  - `node tests/dom_fallback_smoke.mjs` ✅
  - `node tests/comment_tests.mjs` ✅
  - `node tests/table_tests.mjs` ✅
  - `node tests/list_tests.mjs` ✅
  - `node tests/integration_tests.mjs` ✅
  - `node tests/formatting_tests.mjs` runs, but contains pre-existing logical test failures unrelated to fallback wiring (process exits 0).
- Phase 2 updates:
  - Split `engine/format-application.js` into focused helpers:
    - `engine/format-paragraph-targeting.js`
    - `engine/format-span-application.js`
    - slimmed orchestration module in `engine/format-application.js`
  - Pruned stale exports/dead helpers:
    - removed unused legacy `checkOxmlForFormatting` from `engine/format-extraction.js`
    - narrowed helper exports in `engine/surgical-mode.js` and `engine/reconstruction-mode.js`
    - removed unused `w:pPrChange` builder and stale helper exports in `engine/run-builders.js` / `engine/rpr-helpers.js`
  - Phase 2 verification rerun:
    - `node tests/standalone_smoke.mjs` ✅
    - `node tests/no_word_api_standalone_check.mjs` ✅
    - `node tests/dom_fallback_smoke.mjs` ✅
    - `node tests/include_numbering_behavior.mjs` ✅
    - `node tests/comment_tests.mjs` ✅
    - `node tests/table_tests.mjs` ✅
    - `node tests/list_tests.mjs` ✅
    - `node tests/integration_tests.mjs` ✅
    - `node tests/highlight_tests.mjs` ✅
    - `node tests/formatting_tests.mjs` runs with pre-existing logical FAIL messages (exit code 0)
- P2.3 serialization options normalization:
  - Normalized `serializeToOoxml` internals to a single options-object contract:
    - `author`
    - `generateRedlines`
    - `font`
  - Removed mixed string-vs-object helper handling in serialization internals.
  - Standardized font application behavior for both hinted and non-hinted run serialization paths.
  - Normalized `wrapInDocumentFragment` options path to object contract (`DocumentFragmentOptions`) with compatibility for legacy boolean input.
  - Added shared typedefs in `core/types.js`:
    - `SerializationOptions`
    - `DocumentFragmentOptions`
  - Updated JSDoc contract usage in:
    - `pipeline/serialization.js`
    - `pipeline/pipeline.js`
    - `services/package-builder.js`
- P2.3 verification rerun:
  - `node tests/standalone_smoke.mjs` ✅
  - `node tests/no_word_api_standalone_check.mjs` ✅
  - `node tests/include_numbering_behavior.mjs` ✅
  - `node tests/table_tests.mjs` ✅
  - `node tests/list_tests.mjs` ✅
  - `node tests/integration_tests.mjs` ✅
  - `node tests/comment_tests.mjs` ✅
  - `node tests/highlight_tests.mjs` ✅
  - `node tests/formatting_tests.mjs` runs with pre-existing logical FAIL messages (exit code 0)
- P3.1 revision provider unification:
  - Added shared revision metadata helpers in `core/types.js`:
    - `getRevisionTimestamp()`
    - `createRevisionMetadata()`
  - Migrated track-change ID/date generation away from ad hoc random/date calls in:
    - `engine/run-builders.js`
    - `pipeline/serialization.js`
    - `services/table-reconciliation.js`
    - `services/comment-engine.js`
    - `engine/format-application.js`
    - `engine/oxml-engine.js` (text-to-table deletion wrappers)
- P3.4 paragraph offset policy unification:
  - Added shared paragraph boundary policy module:
    - `core/paragraph-offset-policy.js`
  - Migrated paragraph-boundary offset handling to shared policy in:
    - `pipeline/ingestion.js`
    - `engine/format-extraction.js`
    - `engine/reconstruction-mode.js`
    - `engine/surgical-mode.js`
    - `engine/format-paragraph-targeting.js`
  - Updated docs to reflect new core policy module:
    - `src/taskpane/modules/reconciliation/ARCHITECTURE.md`
    - `src/taskpane/modules/reconciliation/README.md`
- P3.1/P3.4 verification rerun:
  - `node tests/standalone_smoke.mjs` ✅
  - `node tests/no_word_api_standalone_check.mjs` ✅
  - `node tests/dom_fallback_smoke.mjs` ✅
  - `node tests/include_numbering_behavior.mjs` ✅
  - `node tests/comment_tests.mjs` ✅
  - `node tests/table_tests.mjs` ✅
  - `node tests/list_tests.mjs` ✅
  - `node tests/integration_tests.mjs` ✅
  - `node tests/highlight_tests.mjs` ✅
  - `node tests/formatting_tests.mjs` runs with pre-existing logical FAIL messages (exit code 0)
- P3.2 hot-path lookup indexing:
  - Added indexed operation lookups in `services/table-reconciliation.js` for:
    - row delete by `gridRow`
    - cell modify by `gridRow:gridCol`
    - sorted row inserts
  - Added indexed diff lookups in `pipeline/patching.js` for:
    - insert operations by `startOffset`
    - non-insert coverage operations for run patching
  - Added sweep-line format hint overlap lookup in `engine/format-application.js` to avoid full hint scans per span.
- P3.3 shared XML query helpers:
  - Added `core/xml-query.js` with shared helpers:
    - `getElementsByTag` / `getFirstElementByTag`
    - `getElementsByTagNS` / `getFirstElementByTagNS`
    - `getElementsByTagNSOrTag` / `getFirstElementByTagNSOrTag`
    - `getXmlParseError`
  - Migrated XML query call sites in:
    - `engine/oxml-engine.js`
    - `engine/table-cell-context.js`
    - `engine/format-extraction.js`
    - `engine/format-application.js`
    - `engine/reconstruction-mode.js`
    - `engine/surgical-mode.js`
    - `engine/run-builders.js`
    - `pipeline/ingestion.js`
    - `pipeline/pipeline.js`
    - `services/comment-engine.js`
- P3.2/P3.3 verification rerun:
  - `node tests/standalone_smoke.mjs` ✅
  - `node tests/no_word_api_standalone_check.mjs` ✅
  - `node tests/dom_fallback_smoke.mjs` ✅
  - `node tests/include_numbering_behavior.mjs` ✅
  - `node tests/comment_tests.mjs` ✅
  - `node tests/table_tests.mjs` ✅
  - `node tests/list_tests.mjs` ✅
  - `node tests/integration_tests.mjs` ✅
  - `node tests/highlight_tests.mjs` ✅
  - `node tests/formatting_tests.mjs` runs with pre-existing logical FAIL messages (exit code 0)

## Phase 2 Backlog (Medium Risk)

- [x] P2.1 Split `engine/format-application.js` into focused modules
- [x] P2.2 Remove dead code and prune stale exports/stubs
- [x] P2.3 Normalize serialization options contract

## Phase 3 Backlog (Higher Risk)

- [x] P3.1 Unify revision ID/date provider usage
- [x] P3.2 Add hot-path lookup indexing for patching/format/table loops
- [x] P3.3 Introduce shared XML query helpers and migrate call sites
- [x] P3.4 Unify paragraph offset policy across extraction/reconstruction

## Notes

- This list tracks execution order for the plan in:
  - `plans/oxml-engine-phase1-refactor-plan.md`
- Phase 1 aims to be behavior-preserving and low regression risk.
