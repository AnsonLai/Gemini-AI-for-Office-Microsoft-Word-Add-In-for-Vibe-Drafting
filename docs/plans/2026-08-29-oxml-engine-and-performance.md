# OXML Engine and Performance Plan

**Date:** 2026-08-29  
**Status:** Active — consolidated plan; no work has been executed from this document.

## Objective

Finish the remaining OXML engine verification and performance work without
reopening the refactors that are already complete.

## Completed baseline

- XML and logger adapters, focused engine modules, deduplication, JSDoc/section
  cleanup, standalone entrypoint support, and the Phase 4 complexity work are
  complete.
- Diff optimization and the `splitSpans` single-pass algorithm are complete for
  their planned scope.

## Remaining work

1. **Close engine verification (former Phase 7).** Update the current test
   layout, add smoke checks, and verify that no Word API leaks across the
   standalone/package boundary. Adapt checks to the extracted package rather
   than recreating the removed in-repo reconciliation tree.
2. **Reduce parse/serialize churn (P5.1).** Reuse serializers and parsed DOMs
   across non-pipeline hot paths, and confirm that lazy paragraph-property
   serialization remains safe.
3. **Profile surgical allocations (P5.2).** Run a targeted large-table-cell
   profile and change allocation patterns only where the profile shows a gain.
4. **Benchmark string operations (P5.3).** Consolidate namespace/string rewrite
   work only when the performance harness demonstrates a measurable benefit.
5. **Collapse Word sync clusters (P5.4).** Address remaining `modify_text`,
   `replace_range`, and fallback-heavy paths with scoped prefetch/cache and
   explicit invalidation.
6. **Reduce memory churn (P5.5).** Review `RunModel` copying and lazy allocation
   of reconstruction maps/hint structures.
7. **Extend web runtime tuning (P5.7).** Expand lazy loading and tune browser
   thresholds using browser-side measurements.
8. **Build the shared `DocumentIndex` (P5.8).** Remove duplicate paragraph/span
   traversals from formatting flows.
9. **Optional, benchmark-gated follow-ups (P5.6/P5.9).** Add diff-result
   caching or deferred DOM mutation only if profiling justifies it.

## Verification gates

- Existing golden/performance harnesses remain green after each focused change.
- `npm run build:dev` and the relevant Node suites pass.
- No standalone/package boundary regression is introduced.
- Performance claims include before/after measurements.

## Migrated source plans

- `docs/plans/migrated/2026-02-05-oxml-engine-refactor-plan.md`
- `docs/plans/migrated/2026-02-05-oxml-engine-refactor-task-list.md`
- `docs/plans/migrated/2026-02-07-oxml-engine-phase1-refactor-plan.md`
- `docs/plans/migrated/web-performance-optimization-plan.md`

