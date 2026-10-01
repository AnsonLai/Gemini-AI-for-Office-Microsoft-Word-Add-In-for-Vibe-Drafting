# OOXML Engine and Performance Plan

> **Retrospective ownership review (2026-09-30):** See the [library offload review](../../library-offload-review.md) for measured code changes, responsibilities delegated to the library, and remaining consumer/native routes. Version pins and validation counts in this closure record describe their dated checkpoints; current consumers pin exact 0.8.3. All five original plans are closed; remaining canonical list work is in the [fresh follow-up](../2026-09-30-canonical-list-migration-follow-up.md).

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** Complete (2026-09-30) — current-version profiles, measured startup optimization and Word verification recorded.

**Recommended order:** 4 of 4 August plans.

## Baseline and dependencies

Both consumers use exact `@ansonlai/docx-redline-js@0.8.2`. The [September upgrade](2026-09-09-docx-redline-js-v0.5.4-upgrade.md) and reliability plan are complete, including actual Office.js transport validation.

Follow [reliability](2026-08-29-reliability-and-quality-gates.md), [agentic tools](2026-08-29-agentic-tools-and-list-reliability.md) and [package boundaries](2026-08-29-package-boundaries-and-integrations.md). Correctness and portability take priority over throughput.

## Completed / inherited work

| Original work | Current disposition |
| --- | --- |
| Implement `executePureOoxmlBatch` | Done in upgrade WP3 |
| Remove iterative redline proxy loop and table recovery | Done |
| Create formatting/structural fidelity suite | Done; mandatory v0.8.2 boundary regressions included |
| Create benchmark harness | Done: `scripts/benchmark-ooxml-pipeline.mjs` |
| Native Word insertion fidelity | COM lane and actual Office.js body/comment insertion pass |
| Document-level binary transport | Proposed, not implemented; see package-boundaries plan |

The redline path reads scope OOXML once, resolves operations against an immutable source, applies one atomic batch and inserts once on successful change. Tracking/transport synchronization may require additional calls. No-op/preparation failure perform no insertion.

## Observed performance and limits

The upgrade's **v0.8.1** benchmark measured ten localized edits:

| Paragraphs | Core median | Core p95 |
| --- | --- | --- |
| 100 | 35.610 ms | 39.126 ms |
| 1,000 | 296.499 ms | 393.545 ms |

Historical raw samples: `scripts/ooxml-benchmark-latest.json`. Current v0.8.2 ten-edit core medians/p95 are 28.094/30.266 ms at 100 paragraphs and 228.094/313.311 ms at 1,000. Current reports include machine metadata and separate phases; the historical run is not a controlled comparison. The user accepts the slower 1,000-paragraph workload. No universal sub-100ms, sub-10ms diff, sub-5ms no-op or 30-second suite gate is imposed.

The pre-extraction taskpane entry was approximately 795 KiB. Optional tool/engine loading now reduces initial JavaScript by about 59%; the full set is about 10% larger after deferred modules load. Builds pass with remaining webpack size warnings. This is an initial-loading improvement, without a universal bundle target.

## Work packages and closure record

### Execution record (2026-09-30)

- Baseline commit: `b8f23b9`; exact library 0.8.2. Package boundaries are complete and archived.
- Profiling ownership is split across workload/core/package measurements, production bundle/startup analysis, and real Word transport instrumentation.
- No new timing gate is imposed. Consumer changes require measured benefit and preserved fidelity; library internals and defects remain separate work items.
- No model/provider calls are required for these measurements.
- WP1 current-version profiles are recorded for six supported workloads with ten samples/two warmups, including 100/1,000 paragraph runs. Ten-edit core medians are 28.094/228.094 ms; raw reports and phase definitions are in the [execution record](../../validation-reports/2026-09-30-ooxml-engine-and-performance.md).
- Workload/host profiling checkpoint committed as `6b44f51`.
- Actual Word timing instrumentation passes 8 Office.js and 12 independent Word checks. Baseline and final startup profiles pass in Word, with all three deferred content-hashed chunks fetched and initialized.
- Final profiled initial scripts fall from 1,041,664 to 430,950 bytes (58.6%). Usable UI observations were 217.4 ms baseline and 172.2 ms final; final deferred load was 46.1 ms. These are local observations, not startup percentiles.
- Final verification: 49 offline suites passed, zero failed, four excluded; development and production builds passed. The provider transport guard remains intact. No provider calls, library changes or deployment occurred. Temporary collectors and registrations were removed.

### WP1 — Profile current consumer workloads

**Status: Complete.** Current-version mapping/core/package/browser projection and actual Word transport costs are separated, with metadata/raw samples. No new timing budget is justified by the accepted workload.

- Reuse `npm run benchmark:ooxml`; record version, machine, workload, sample count, median/p95 and raw observations.
- Cover representative no-op, localized multi-edit, full paragraph, nested list/table and commented-document workloads as their paths settle.
- Separate core application, package open/save, operation mapping and actual Word read/write latency.
- Treat heap deltas as observations, not reliable peak-memory measurements. Investigate sustained retention if evidence suggests it.
- Set workload-specific budgets after reproducible measurements and a user-visible need.

### WP2 — Improve taskpane startup and consumer overhead

**Status: Complete.** Optional editing/engine modules load through cached, retryable loaders; platform and tool initialization order is preserved. Real Word validates the deferred chunks. The measured payload benefit and full-size/first-use tradeoff are recorded. No additional inspection cache or speculative micro-optimization was justified.

- Profile bundle composition/time to usable UI; identify heavy modules loaded before needed.
- Consider lazy imports for optional editing/UI paths where measurements justify them.
- Reduce duplicate consumer inspection/serialization only when profiles identify it; preserve immutable targeting and atomic batches.
- Verify development/production builds, offline contracts and relevant actual Word paths.
- Stop when the measured problem is resolved; keep upstream internals out of consumer changes.

### WP3 — Route engine findings upstream

**Status: Complete for this profiling pass.** No new library regression was established; existing list fidelity defects remain separate in the agentic plan. No library internals were modified.

- Library parse/serialize churn, diff allocation, run reconstruction and numbering algorithms belong to `Docx Redline JS`.
- Reproduce engine defects/performance regressions with a minimal source, exact operations, version and measurements.
- Raise a separate library issue/work item before changes there; do not add consumer fidelity workarounds.
- After an upstream release, update pins and run relevant compatibility/fidelity checks.

## Decisions carried forward

- Do not recreate a consumer DocumentIndex, diff cache or deferred DOM layer without demonstrated need.
- Retired proxy-loop optimizations are superseded by the implemented batch path.
- Upstream ownership does not mean every optimization/verification is complete; use evidence rather than assumed timings.
- Real Word no-repair/revision/package checks remain separate from offline tests.
- PDF export is optional and outside required performance scope.

## Acceptance

- [x] Batch runner, fidelity suite and benchmark harness exist.
- [x] Observed 1,000-paragraph throughput accepted for the upgrade.
- [x] Current-version profiles separate consumer/core/package/host costs.
- [x] Selected optimizations show before/after benefit without fidelity regression.
- [x] Library findings recorded separately; consumer changes respect the boundary.

The performance plan is closed. Separately tracked agentic list fidelity and capability follow-ups remain open in their own plan.
