# OOXML engine and performance — execution record

## Scope and baseline

Consumer baseline: `b8f23b9`, exact `@ansonlai/docx-redline-js@0.8.2` in
the add-in and MCP. The package-boundary plan is complete. The work checkpoint
is `31f70e0`.

This lane measures workload mapping, core application, package open/save,
Word host transport and taskpane startup separately. No provider calls or
universal timing threshold are required. The accepted 1,000-paragraph workload
does not justify a library rewrite by itself.

## WP1 current-version workload profiles

The expanded benchmark verifies six workloads: no-op, ten localized edits,
full-paragraph rewrite, localized text inside a Word-authored nested list,
localized text inside an existing table, and a new comment in a Word-authored
threaded-comment document. Existing structures are retained; these cases do not
certify the known unsupported list transformations.

Both runs use ten measured samples per phase and two warmups on an AMD Ryzen 7
9800X3D, Windows x64, Node 24.11.1, library 0.8.2. Raw samples, fixture hashes,
machine/version metadata and methodology are in the
[100-paragraph report](2026-09-30-ooxml-performance.json) and
[1,000-paragraph report](2026-09-30-ooxml-performance-1000.json).

Ten-edit workload; values are median / p95 in milliseconds:

| Phase | 100 paragraphs | 1,000 paragraphs |
| --- | ---: | ---: |
| Operation mapping | 0.005 / 0.005 | 0.009 / 0.010 |
| Portable baseline capture | 1.893 / 2.264 | 22.862 / 25.375 |
| Package open | 0.226 / 0.394 | 0.322 / 0.594 |
| Package inspection | 1.427 / 2.000 | 18.962 / 23.063 |
| Browser prompt projection | 3.645 / 4.298 | 44.690 / 50.513 |
| Core application | 28.094 / 30.266 | 228.094 / 313.311 |
| Package apply | 32.801 / 37.805 | 266.937 / 288.315 |
| Package serialization | 0.233 / 0.351 | 1.442 / 2.325 |
| Open + apply + save | 31.761 / 33.169 | 276.331 / 292.532 |

These phases are measured independently. Their medians are not additive.
Preparation and output checks run outside each timed action. Heap deltas are
observational, not a peak-memory measurement.

The current core figures are below the historical 0.8.1 observations at both
sizes. The old report lacks machine metadata, so this does not establish a
controlled version improvement. Application dominates this 1,000-paragraph
workload; no engine defect or regression is inferred from that result.

### Actual Word host phases

The opt-in adapter observer measures source `getOoxml`/sync, portable preparation,
actual insertion/sync and adapter total. Tracking requested inside the adapter
is included in its total; outer tool tracking is measured only by full tool time
when the production-list harness runs. Invalid clocks and observer exceptions do not affect edits;
timing-only options are removed before portable preparation. No-op/refused
batches have no insertion phase and perform zero writes.

Desktop Word 16.0.20430.20092 (PC) passed all eight transport checks. Two changed
workloads produced these single samples, in milliseconds:

| Workload | Read/sync | Preparation | Insert/sync | Adapter total |
| --- | ---: | ---: | ---: | ---: |
| Plain replacement | 32.2 | 9.8 | 39.6 | 81.7 |
| Threaded-comment reply | 95.3 | 13.8 | 49.7 | 159.4 |

These are actual host observations, not percentiles. The preliminary source
baseline and DOCX export are outside this adapter timer. Independent Word
verification passed 12 source/tracked/accepted/rejected and thread/package
checks. Evidence: [Office.js phases](2026-09-30-performance-officejs.json),
[Word oracle](2026-09-30-performance-word.json).
The collector was stopped and its temporary registration removed.

## WP2 taskpane startup

Pre-change production taskpane: 814,240 bytes (about 795 KiB), with editing
tools in the initial source graph. A dedicated profiling build records module
evaluation, Office readiness and usable UI. It suppresses automatic Glance
provider calls and posts only to the local collector; ordinary builds retain
their normal behavior.

Actual desktop Word baseline: module evaluation 183.2 ms, Office ready 216.8 ms,
usable UI 217.4 ms. The instrumented taskpane plus polyfill transfers 1,041,664
encoded bytes. This is one local observation, not a cold-start percentile or a
network-wide guarantee. The first launch returned no evidence; a relaunch
completed the measurement. [Baseline report](2026-09-30-taskpane-startup-baseline.json).

The consumer now defers agentic tools, Word preparation and the DOCX engine
until needed. Platform setup precedes baseline parsing; tool dependencies
initialize once. Concurrent first requests share one import, and failed loads
can be retried. Tool import failure occurs before the mutation loop. Navigation
response assertions still pass with the module namespace.

The final profiled initial scripts are 430,950 bytes, **58.6% below the baseline**.
The ordinary production entry is smaller because profiling code is excluded.
The final local Word observation reaches usable UI at 172.2 ms and loads the
deferred tool module in 46.1 ms after that milestone. All three content-hashed
chunks were fetched and initialization succeeded.
[Final Word startup report](2026-09-30-taskpane-startup-optimized.json).
These samples establish working paths and initial payload reduction; they do
not establish a startup percentile or promise a particular network latency.

The full set of JavaScript is about 10% larger once all deferred modules load.
The selected benefit is reduced initial loading; first use pays the deferred
load. Deploy the entire `dist/` output, including content-hashed chunks. The
stable entry filenames preserve the existing HTML/manifest paths.

Ordinary production initial scripts: **1,038,203 → 425,270 bytes (59.04% lower)**.
All emitted JavaScript including commands: **1,038,570 → 1,143,898 bytes
(10.15% larger)**. Raw build/Word comparisons and asset sizes are retained in
the [bundle report](2026-09-30-taskpane-startup-bundle.json).

Startup POST handling is isolated in a local-only diagnostic helper, preserving
the taskpane's no-direct-provider-fetch guard. Ordinary builds do not post
startup evidence or suppress automatic Glance. The profiling build performs
no model request or mutation. Resource/JavaScript errors are reported by the
local collector's startup diagnostic hook.

Final validation: **49 offline suites passed, zero failed, four excluded
entrypoints**; existing historical skips and known list defects remain explicit.
Development and production builds pass. Eight real Office.js and twelve
independent Word fidelity checks pass. All temporary collectors were stopped
and registrations removed.

## WP3 library ownership

No new library performance regression was established. Existing list fidelity
defects stay in the separate [agentic list plan](../plans/completed/2026-08-29-agentic-tools-and-list-reliability.md).
No library internals or consumer fidelity workaround were changed. No diff
cache, alternate document index or arbitrary timing budget was introduced.

Reproduction and profiling instructions: [OOXML performance](../ooxml-performance.md).

The plan is complete and archived in `docs/plans/completed/`. The workload/host
checkpoint is `6b44f51`; startup optimization and plan closure follow it.
