# OOXML Engine and Performance Plan

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** Original batching, fidelity-suite and benchmark setup completed in the upgrade; measured consumer performance work remains.

**Recommended order:** 4 of 4 August plans.

## Baseline and dependencies

Both consumers use exact `@ansonlai/docx-redline-js@0.8.2`. The [September upgrade](2026-09-09-docx-redline-js-v0.5.4-upgrade.md) is near complete with actual Office.js transport pending.

Follow [reliability](2026-08-29-reliability-and-quality-gates.md), [agentic tools](2026-08-29-agentic-tools-and-list-reliability.md) and [package boundaries](2026-08-29-package-boundaries-and-integrations.md). Correctness and portability take priority over throughput.

## Completed / inherited work

| Original work | Current disposition |
| --- | --- |
| Implement `executePureOoxmlBatch` | Done in upgrade WP3 |
| Remove iterative redline proxy loop and table recovery | Done |
| Create formatting/structural fidelity suite | Done; mandatory v0.8.2 boundary regressions included |
| Create benchmark harness | Done: `scripts/benchmark-ooxml-pipeline.mjs` |
| Native Word insertion fidelity | Passing COM lane; actual Office.js pending |
| Document-level binary transport | Proposed, not implemented; see package-boundaries plan |

The redline path reads scope OOXML once, resolves operations against an immutable source, applies one atomic batch and inserts once on successful change. Tracking/transport synchronization may require additional calls. No-op/preparation failure perform no insertion.

## Observed performance and limits

The upgrade's **v0.8.1** benchmark measured ten localized edits:

| Paragraphs | Core median | Core p95 |
| --- | --- | --- |
| 100 | 35.610 ms | 39.126 ms |
| 1,000 | 296.499 ms | 393.545 ms |

Raw samples: `scripts/ooxml-benchmark-latest.json`. Measurements exclude actual Office.js latency and have not been rerun on v0.8.2. The user accepts the slower 1,000-paragraph workload. No universal sub-100ms, sub-10ms diff, sub-5ms no-op or 30-second suite gate is imposed.

Production builds pass with bundle-size warnings (taskpane approximately 747 KiB). Investigate startup impact before choosing a bundle target.

## Remaining work packages

### WP1 — Profile current consumer workloads

- Reuse `npm run benchmark:ooxml`; record version, machine, workload, sample count, median/p95 and raw observations.
- Cover representative no-op, localized multi-edit, full paragraph, nested list/table and commented-document workloads as their paths settle.
- Separate core application, package open/save, operation mapping and actual Word read/write latency.
- Treat heap deltas as observations, not reliable peak-memory measurements. Investigate sustained retention if evidence suggests it.
- Set workload-specific budgets after reproducible measurements and a user-visible need.

### WP2 — Improve taskpane startup and consumer overhead

- Profile bundle composition/time to usable UI; identify heavy modules loaded before needed.
- Consider lazy imports for optional editing/UI paths where measurements justify them.
- Reduce duplicate consumer inspection/serialization only when profiles identify it; preserve immutable targeting and atomic batches.
- Verify development/production builds, offline contracts and relevant actual Word paths.
- Stop when the measured problem is resolved; keep upstream internals out of consumer changes.

### WP3 — Route engine findings upstream

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
- [ ] Current-version profiles separate consumer/core/package/host costs.
- [ ] Selected optimizations show before/after benefit without fidelity regression.
- [ ] Library findings recorded separately; consumer changes respect the boundary.
