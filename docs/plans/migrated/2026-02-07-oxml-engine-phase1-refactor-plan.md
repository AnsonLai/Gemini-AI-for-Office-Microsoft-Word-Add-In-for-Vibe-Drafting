# OXML Engine Refactor Plan (Phase 1 First)

> **Migrated on 2026-08-29:** Remaining engine/performance follow-up was
> consolidated into [`2026-08-29-oxml-engine-and-performance.md`](../2026-08-29-oxml-engine-and-performance.md).
> This document is retained in `migrated/` as historical detail.

## Scope

This plan captures the full set of refactors identified in the `reconciliation` stack and prioritizes a minimal-risk Phase 1 for immediate execution.

Date: 2026-02-07

## Objectives

1. Improve maintainability and module boundaries without changing core behavior.
2. Reduce coupling and duplication across engine/pipeline/services.
3. Improve predictability and testability for IDs, packaging, and XML handling.
4. Execute low-risk structural cleanup first, then progressively deeper refactors.

## Refactor Backlog (All Identified Items)

1. Consolidate OOXML package builders into one shared service.
   - Current duplication:
     - `src/taskpane/modules/reconciliation/pipeline/serialization.js`
     - `src/taskpane/modules/reconciliation/engine/table-cell-context.js`
     - `src/taskpane/modules/reconciliation/services/comment-engine.js`
   - Target: one `package-builder` module for document fragment packaging with options (`numbering`, `comments`, paragraph-only replacement).

2. Remove barrel import cycle in integration layer.
   - Current cycle:
     - `integration/integration.js` imports `ReconciliationPipeline` from `../index.js`
     - `index.js` re-exports integration helpers
   - Target: integration imports pipeline directly from `../pipeline/pipeline.js`.

3. Centralize list marker detection/parsing.
   - Duplicate list marker regex/parsing in:
     - `engine/oxml-engine.js`
     - `pipeline/pipeline.js`
     - `pipeline/patching.js`
   - Target: one shared parser utility (`list-parser` or `list-markers`) used everywhere.

4. Break down `engine/format-application.js` into smaller focused modules.
   - Current file remains large and multi-responsibility.
   - Target split:
     - paragraph targeting/matching
     - span splitting/boundary utilities
     - format-only orchestration
     - run formatting application helpers

5. Remove dead/legacy paths and tighten public API.
   - Candidates:
     - unused helper(s) in `pipeline/ingestion.js`
     - unused exported functions in `engine/format-application.js`
     - unfinished stub exports in `pipeline/pipeline.js` (`detectContentType`, `parseListItems`)
   - Target: remove or internalize unused APIs; deprecate or implement stubs.

6. Standardize revision ID/date generation.
   - Current state: mixed usage of global counter and local `Math.random()` IDs.
   - Target: one shared ID/date provider in core utilities used by track-change builders and services.

7. Pre-index hot-path operations/hints to avoid repeated scans.
   - Current repeated scans (`filter`/`find`) in patching/table/format loops.
   - Target: indexed maps by offsets and row/col keys for O(1) lookups in core loops.

8. Normalize XML query strategy through shared helpers.
   - Current mix of `getElementsByTagNameNS` and prefixed `w:*` lookups.
   - Target: `xml-query` helpers for common node discovery and parse error checks.

9. Normalize serialization options signature.
   - `pipeline/serialization.js` currently supports mixed option forms and inconsistent handling.
   - Target: single options object contract and consistent formatting/font handling.

10. Unify paragraph offset policy across extraction/reconstruction.
   - Current newline/paragraph boundary offset handling is not centralized.
   - Target: shared offset policy utility used by extraction + reconstruction paths.

11. Fix numbering wrapper boolean bug.
   - `engine/oxml-engine.js` currently uses `result.includeNumbering || true`, which always resolves to `true`.
   - Target: use `??` or explicit boolean handling.

## Execution Strategy

### Phase 1 (Minimal Risk, Start Here)

Focus: non-invasive structural cleanup, duplication reduction, and correctness fixes with low regression risk.

Included items:
1. Package builder consolidation (item 1)
2. Integration import cycle removal (item 2)
3. Shared list marker parser utility (item 3)
4. Numbering boolean bug fix in `oxml-engine` (item 11)

Out of scope for Phase 1:
- Large internal decomposition (`format-application` split)
- Performance indexing rewrites in core loops
- Deep API removals and offset model changes

### Phase 2 (Medium Risk)

Focus: internal cleanup and API surface tightening.

Included items:
1. `format-application` split (item 4)
2. dead/legacy cleanup and export pruning (item 5)
3. serialization options normalization (item 9)

### Phase 3 (Higher Risk / Behavior-Sensitive)

Focus: correctness/performance internals with higher chance of edge-case drift.

Included items:
1. ID/date provider unification (item 6)
2. hot-path indexing (item 7)
3. XML query normalization (item 8)
4. paragraph offset unification (item 10)

## Definition of Done (Per Phase)

1. All affected tests pass (existing suites + targeted new tests).
2. No new circular dependencies in `reconciliation`.
3. `standalone_smoke` and `no_word_api_standalone_check` pass.
4. No behavior drift in list generation, table reconciliation, comment injection, and format-only workflows.
5. Architecture docs updated to reflect moved/shared modules.

## Validation Requirements

1. Run targeted tests for:
   - list generation/parsing
   - table reconciliation
   - comment injection/package wiring
   - format-only and format-removal paths
2. Add focused regression cases for:
   - package builder reuse across all call sites
   - list marker equivalence across router/pipeline/patching
   - `includeNumbering` false/undefined behavior
3. Re-run standalone checks:
   - `tests/standalone_smoke.mjs`
   - `tests/no_word_api_standalone_check.mjs`
