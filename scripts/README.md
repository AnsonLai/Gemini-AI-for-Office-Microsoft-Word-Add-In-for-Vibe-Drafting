# Verification and profiling commands

## Testing the library boundary

See the [testing overview](../docs/TESTING.md) for lane responsibilities and limits.

The [library offload review](../docs/library-offload-review.md) records which
execution and packaging responsibilities moved into the shared engine. Consumer
tests still own operation mapping, exact source targets, failure propagation,
and host outcomes. Package success alone cannot confirm a Word insertion.

Use the offline compatibility, batch, MCP workflow and cross-host parity suites
for library upgrades; use actual Office.js plus independent Word save/reopen and
Accept All/Reject All checks for affected Word routes. Native list routes need
their own expectations because they do not execute a canonical library batch.
Keep unsupported canonical conversions in the active follow-up, and report
library defects upstream rather than adding a second implementation here.

## Offline tests

`npm test` runs each consumer JavaScript suite in its own Node process,
including the MCP stdio workflow and the golden guardrail with `--verify`.
Suites are discovered recursively under `tests/` and
`mcp/docx-server/tests/`; new suites enter the offline lane automatically.
Failures, timeouts and known printed failure markers fail the command. Use
`node scripts/run-all-tests.mjs --list` to inspect the inventory.

`DOCX_TEST_CONCURRENCY` sets parallel processes (default 4), and
`DOCX_TEST_TIMEOUT` sets each suite timeout in milliseconds (default 180000).
Excluded entrypoints and suite-internal skips are reported. The historical
v0.5.4 suite runs its consumer guards but skips version-specific behavior on
the current 0.8.3 package.

The offline lane excludes live provider/model evaluation, desktop Word fixture
generation, XML provider setup and observational performance scripts. Passing
it does not certify native Word behavior.

## Desktop Word lane

On Windows with desktop Word installed, run:

```powershell
npm run test:word
```

This lane opens fixtures without repair, exercises Word Accept All/Reject All,
comment-thread identities, existing footer text, formatting, tab stops and
section layout. Two insertion checks read real `Content.WordOpenXML`, replay
the exact operations through the production add-in bridge, insert with native
`Range.InsertXML`, save, reopen and inspect revisions. This tests desktop Word
and the production bridge; COM transport does not certify Office.js
`insertOoxml` itself. PDF export is optional (`-Render`) and is not an upgrade
gate; successful export still requires visual inspection.

The runner has a bounded timeout and records stalled calls and partial results.
It stops only a Word process newly created by its worker with a matching
creation time. Useful options include `-SkipNativeInsert`, `-CaseName`,
`-VisibleWord`, `-ArtifactsDir` and `-TimeoutSeconds`; see the script help or
the active validation report for a reproducible invocation. The current
v0.8.3 Word run includes the two former Reject All findings as passing
public-facade cases: 12 independent Word checks passed. The dated v0.8.2
failures remain historical evidence. See the [v0.8.3 report](../docs/validation-reports/2026-09-30-docx-redline-v083.md)
for the complete current matrix and its artifacts.

Golden XML may be exported for review:

```powershell
node tests/phase4/golden-guardrail.mjs --export-dir .cache/wp6/golden-review --export-only
node tests/phase4/golden-guardrail.mjs --verify
```

`--verify` compares without rewriting tracked output. Use `--update` only after
reviewing raw differences and their provenance.

## OOXML benchmark

`npm run benchmark:ooxml` observes six deterministic workloads: no-op, ten
localized edits, full-paragraph rewrite, localized text in a Word-authored
nested list, localized text in an existing table, and a comment in a
Word-authored threaded-comment document. Paragraph fixture size defaults to
100; operation count defaults to 10. The multi-edit fixture has at least as
many paragraphs as requested operations. The paragraph setting is capped at
1,000 before this minimum is applied, so a deliberately larger operation count
can create a larger multi-edit fixture.

The benchmark uses two warmups and ten measured samples per phase by default.
It reports raw samples, median and nearest-rank p95 with machine, runtime,
package and fixture metadata. Assertions run after each timed action. Mapping,
baseline capture, package open/inspection, browser prompt projection, core
application, package application, serialization and open/apply/save are timed
as separate observations. Their durations are not additive. Fixture creation,
Word calls, disk I/O, model calls and transport are excluded. Heap deltas are
not peak-memory measurements. There is no default pass/fail timing threshold.

Override workloads with `DOCX_BENCH_PARAGRAPHS`,
`DOCX_BENCH_OPERATIONS`, `DOCX_BENCH_WARMUPS` (zero allowed) and
`DOCX_BENCH_ITERATIONS` (positive integers). Save a full report:

```powershell
npm run benchmark:ooxml -- --output=.cache/ooxml-performance/report.json
```

The runner creates the parent directory. Compare runs only when workload and environment
match. The measurements do not establish Word round-trip latency or a universal
performance guarantee.

## Actual Office.js transport lane

This development-only lane requires desktop Word on Windows and already trusted
development certificates. It does not install certificates or change trust
settings. The collector uses a separate validation add-in identity and local
loopback evidence server; it makes no provider calls. Port 3000 must be
available.

```powershell
npx webpack --mode development --env WORD_HOST_VALIDATION=1
node scripts/run-officejs-validation.mjs --launch
```

The validation entry is absent from ordinary builds. It forwards calls to real
Word proxies through the production atomic batch bridge, checks reads/writes,
no-op and refusal behavior, and exports DOCX bytes using Office's compressed
file API. Seeding fixtures and exporting bytes are harness operations; the
production edit path is `getOoxml` → portable preparation → `insertOoxml`.
Inspect `.cache/reliability/officejs/officejs-report.json`; status must be
`passed` and timestamps current. Collector checks alone do not prove Word's
accepted/rejected result.

Reopen the collector's export manifest through the independent desktop Word
oracle, then clean up the temporary registration:

```powershell
npm run test:word -- -FixtureManifest .cache/reliability/officejs/officejs-fixtures.json -SkipNativeInsert -ArtifactsDir .cache/reliability/officejs-oracle -TimeoutSeconds 120
node scripts/run-officejs-validation.mjs --cleanup
```

The oracle inspects the source, actual tracked/accepted/rejected views and
engine-resolved references independently. Keep mocked invocation tests,
Office.js proxy checks and the desktop Word oracle distinct in summaries.

## Golden scenario (desktop Word)

`scripts/golden/nda-scenario.mjs` holds the owner's 16-prompt editing session
on `tests/fixtures/golden/sample-nda.docx`: each prompt, the model responses to
replay, and objective checks. The runner builds a dedicated taskpane bundle
(`--env GOLDEN_SCENARIO=1`, into `.cache/golden/dist`), serves it on port 3001
under a separate add-in identity with its own settings storage, and opens it in
a **separate** Word instance (`WINWORD /x`), so a running Word session and the
dev server on 3000 are untouched. A driver types each prompt into the real chat
UI and waits for the production chat loop; after every step the document is
exported and scored on its accepted view (paragraphs, effective formatting
including character styles, lists and number formats, tables, comments,
highlights) plus a check that new redlines carry the configured author.
Earlier checks are regressed at later steps.

```powershell
# Deterministic lane: model responses replayed from the scenario (no provider calls)
node scripts/run-golden-scenario.mjs --launch
# Live lane: real Gemini, scored by the same checks (incurs provider usage)
$env:GEMINI_API_KEY = '...'
node scripts/run-golden-scenario.mjs --mode live --model gemini-flash-latest --launch
# Remove the sideload registration
node scripts/run-golden-scenario.mjs --cleanup
```

Useful options: `--until <n>` (first n steps), `--skip-build`, `--keep-word`
(leave the golden Word instance open), `--step-timeout <seconds>`,
`--artifacts-dir`. Results land in `.cache/golden/<timestamp>/`: `report.md`,
`report.json`, per-step `.docx` exports with transcript/console JSON, and
`model-log.jsonl` (every model request and the response served or received;
never the API key). Replay references such as `{ $p: { startsWith: '...' } }`
resolve against the `[P#]` context inside each intercepted request, so scripted
paragraph numbers follow earlier edits. A replay desync (an unexpected or
missing model call, e.g. a corrective retry) fails the step.

Checks linked to a `knownIssue` report as known library defects
(`docs/library-issues/`) instead of failures. `tests/golden_scenario_tests.mjs`
covers the harness offline (reference resolution, replay queue, checks) and runs
in `npm test`; the scenario itself needs desktop Word.

## Taskpane startup profile

Startup profiling is enabled only in a dedicated build and is not part of
ordinary production behavior:

```powershell
npx webpack --mode production --env TASKPANE_STARTUP_PROFILE=1 --env urlProd=https://localhost:3000/
node scripts/run-officejs-validation.mjs --startup --artifacts-dir .cache/performance/startup --launch
```

The profile records module evaluation, Office readiness and usable UI to the
local collector. It suppresses automatic Glance calls only in that profile
build. Ordinary builds do not post startup reports. Use `--dist-dir` for an
isolated build output when comparing before and after; stop the collector and
remove its validation registration with `--cleanup` after collection. See
[OOXML performance methodology](../docs/ooxml-performance.md) for phase
definitions and reported sample limits.

## Provider and change verification

`node tests/gemini_client_tests.mjs` verifies bounded transient retries,
non-retryable HTTP failures, cancellation, timeout and cleanup using stubs. It
is offline and consumes no provider credits. All provider calls share
`src/taskpane/modules/chat/gemini-client.js`; transport retries do not rerun
document tools.

Optional evals in `tests/evals/run-evals.mjs` call a configured live model and
incur provider usage:

```powershell
$env:GEMINI_API_KEY = '...'
node tests/evals/run-evals.mjs --model <configured-model> --case <case-name>
```

For consumer changes run `npm test`, `npm run build:dev` and `npm run build`.
For host transport, package or revision behavior, run the relevant Word lane.
Performance is observational; PDF export is optional.
