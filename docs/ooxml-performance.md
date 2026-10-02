# OOXML performance measurements

Performance observations are separate from fidelity gates. Both consumers pin
exact `@ansonlai/docx-redline-js@0.8.4`. The dated performance reports below
measured 0.8.2; they have not been rerun as 0.8.4 timing evidence. A slower large document does not justify
changing its tracked-change semantics.

The [library offload review](library-offload-review.md) separates delegated
execution from local optimization. Atomic library batches replace iterative
Word redline execution; extracting portable preparation remains local code.
Loading editing modules on first use reduces initial transfer, while retaining
those modules in the complete build. It is not an engine-code deletion.

## Offline workloads

Run the current consumer benchmark from the repository root:

```powershell
npm run benchmark:ooxml -- --output=.cache/ooxml-performance.json
```

Create the output directory first when using a new directory. The report records
the machine, library version, workload, sample count and raw observations. Compare
the same workload and environment; do not subtract independently measured medians
to invent a phase duration. Heap deltas are observations, not peak memory or proof
of retained objects.

## Actual Word transport

Build the validation entry and run the local collector:

```powershell
npx webpack --mode development --env WORD_HOST_VALIDATION=1
node scripts/run-officejs-validation.mjs --artifacts-dir .cache/performance/officejs --launch
```

The report records source read/sync, portable preparation, insertion/sync and
adapter total separately. Insertion covers the insertion helper and sync;
adapter total also includes tracking management requested inside the adapter.
External tool tracking changes are outside that timer; the production-list
harness reports them only in its separate full tool duration. Baseline capture and
DOCX export outside the adapter are not
part of those phase timings. These are host observations, not portable benchmark
percentiles.

After stopping the collector, remove its temporary registration:

```powershell
node scripts/run-officejs-validation.mjs --artifacts-dir .cache/performance/officejs --cleanup
```

Run the existing independent Word oracle against the exported fixture manifest
when changing the adapter. Timing does not replace Accept All/Reject All and
package checks.

## Taskpane startup

Startup profiling is enabled only in a dedicated build:

```powershell
npx webpack --mode production --env TASKPANE_STARTUP_PROFILE=1 --env urlProd=https://localhost:3000/
node scripts/run-officejs-validation.mjs --startup --artifacts-dir .cache/performance/startup --launch
```

That build records module evaluation, Office readiness and usable UI milestones
to the local collector, and suppresses automatic Glance provider calls. Ordinary
builds keep the normal startup behavior and do not post startup reports.
Use `--dist-dir` to serve an isolated output directory for before/after builds.
Stop the collector and use the same `--cleanup` command with its artifacts path.

Library parse/diff/reconstruction/numbering findings belong in separate upstream
work items with exact sources, operations, versions and raw samples. Consumer
profiling does not authorize patching library internals.
