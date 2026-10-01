# Verification commands

`npm test` runs each consumer JavaScript suite in its own Node process, including
the MCP stdio workflow and the golden guardrail with `--verify`. Suites are
discovered recursively under `tests/` and `mcp/docx-server/tests/`; new suites
automatically enter the offline lane. Exit failures, timeouts, and known printed
failure markers make the command fail. Output is ordered by filename.

`node scripts/run-all-tests.mjs --list` shows the inventory without executing it.
`DOCX_TEST_CONCURRENCY` controls parallel processes (default 4) and
`DOCX_TEST_TIMEOUT` controls the per-suite timeout in milliseconds (default
180000). Every excluded entrypoint and suite-internal skip is reported. The
historical v0.5.4 suite runs its consumer guards but reports its version-specific
package behavior as skipped when testing v0.8.2.

The offline lane explicitly excludes live model/API evaluation, Word Desktop
fixture generation, XML provider setup, and observational performance scripts.
Passing it does not certify the native host. `npm run test:word` runs the separate
live Word Desktop lane on Windows with Word installed. It checks no-repair
opening, Word Accept All/Reject All, comment-thread identities, existing footer
text, formatting, tab stops and section layout. Two insertion cases read real
`Content.WordOpenXML`, replay exact operations through the production add-in
bridge, insert with native `Range.InsertXML`, save, reopen and check revisions.
This exercises real desktop Word and the production bridge, but COM transport
does not certify the Office.js `insertOoxml` call itself.
It has a 120-second timeout, records the stalled call and partial results, and
stops only a Word process newly created by its worker with a matching creation
time. `npm run test:word -- -SkipNativeInsert` runs only the differential lane.
PDF export is optional (`-Render`) and is not an upgrade gate. Its success would
still require visual inspection. `-SkipRender` remains accepted for older commands;
rendering is disabled by default. `-ArtifactsDir .cache/wp6/word-run` selects an
output directory; `-TimeoutSeconds 180` changes the bounded run time. A timeout
fails the requested lane.

`-CaseName word-addin-plain-replacement` selects one case; `-VisibleWord` exposes
the disposable Word window for diagnosis. The four boundary regressions (localized
and full-paragraph forms of both defects) are mandatory in the offline and live
Word lanes with v0.8.2. `-IncludeKnownDefects` remains accepted for compatibility;
there are no longer excluded defect cases. The original defects are tracked as
fixed library issues
[#3](https://github.com/AnsonLai/docx-redline-js/issues/3) and
[#4](https://github.com/AnsonLai/docx-redline-js/issues/4). A portable reproducer
and issue drafts live under `docs/library-issues/`. No library workaround is
applied by the host tests.

Golden XML can be exported for review with
`node tests/phase4/golden-guardrail.mjs --export-dir .cache/wp6/golden-review --export-only`.
`--verify` compares without rewriting tracked latest output. Update hashes with
`--update` only after reviewing raw differences and their provenance.

`npm run benchmark:ooxml` observes independent localized tracked replacements
in 100- and 1,000-paragraph legal-style documents, with 10 operations per batch.
It reports raw samples, median and p95 after 3 warmups and 15 measured iterations.
The core batch starts from XML; separate measurements cover DOCX open, serialization
of an edited document, and the combined open/apply/save lifecycle. Their timings
are separate observations and should not be subtracted or summed into an estimate.
Fixture construction, Word calls, disk I/O, model calls, and transport are excluded.
The correctness assertions run inside the measured actions.

Use `DOCX_BENCH_PARAGRAPHS`, `DOCX_BENCH_OPERATIONS`, `DOCX_BENCH_WARMUPS`, and
`DOCX_BENCH_ITERATIONS` to reproduce a workload. All must be positive integers;
operations must not exceed paragraphs. Save the full report with
`npm run benchmark:ooxml -- --output=<existing-directory>/report.json`.
`--require-sub100ms` optionally fails when any workload's core batch median
reaches 100 ms. The ordinary benchmark is observational: the 100 ms goal is not
a claim about every document size or Word round-trip latency.

This lane structure follows the upstream project's testing methodology:
exact XML and return contracts, independent Word accept/reject expectations,
and measured performance provide different evidence. Visual review remains a
separate judgment even when Word opens and exports a document successfully.

## Actual Office.js transport lane

Use desktop Word on Windows and existing trusted development certificates.
The collector does not install certificates or change trust settings. Close any
previous collector before starting a new run; port 3000 must be available.

```powershell
npx webpack --mode development --env WORD_HOST_VALIDATION=1
node scripts/run-officejs-validation.mjs --launch
```

The second command registers a separate local validation add-in and launches a
disposable Word document. It uses no API key and makes no model calls. The
validation entry point is absent from ordinary development/production builds.
The page seeds synthetic/Word-authored fixtures, forwards calls to genuine Word
proxies through the production batch bridge, counts reads/inserts, checks an
unchanged paragraph and an empty batch, and refuses an invalid target without
insertion. It exports the resulting DOCX bytes through Office's compressed-file
API to a collector bound to 127.0.0.1. See [Microsoft's whole-document API guide](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word?tabs=powerpoint).
Seeding/export are harness operations; the production edit path remains Flat-OPC.

Inspect `.cache/reliability/officejs/officejs-report.json`: it must have status
`passed` and current timestamps. A pending/failed report is not completion. The
collector accepts one claim per run so reopened exported packages do not rerun
edits. It expires after ten minutes; stop it with Ctrl+C after verification.

Then independently reopen and resolve the actual Office.js output in Word:

```powershell
npm run test:word -- -FixtureManifest .cache/reliability/officejs/officejs-fixtures.json -SkipNativeInsert -ArtifactsDir .cache/reliability/officejs-oracle -TimeoutSeconds 120
node scripts/run-officejs-validation.mjs --cleanup
```

The external-manifest option reuses the supervised Word oracle without generating
replacement fixtures. It checks source, actual Office.js tracked/accepted/rejected
views and engine-resolved reference packages. The companion oracle report must
also pass; the collector alone cannot certify Word's accepted/rejected result.
Cleanup removes only this validation add-in's registration. Test documents and
reports remain available for review; close the disposable document when finished.

## Provider and change verification

`node tests/gemini_client_tests.mjs` uses deterministic stub responses to verify
bounded transient retries, non-retryable HTTP failures, cancellation, timeout
and cleanup. It is included in `npm test` and consumes no provider credits.
All provider calls share `src/taskpane/modules/chat/gemini-client.js`; transport
retries do not run document tools.

The usage comment in `tests/evals/run-evals.mjs` describes the optional live model
lane: `node tests/evals/run-evals.mjs --model <configured-model> --case <case-name>`,
with `GEMINI_API_KEY` supplied through the environment.
Run it with explicitly configured credentials/model when evaluating model output;
it incurs provider usage and is not part of the offline correctness gate.
Do not print credentials or raw provider bodies as diagnostic evidence.

For consumer changes, run `npm test`, `npm run build:dev` and `npm run build`.
For transport/package/revision changes, also run the relevant live Word lane.
Run the actual Office.js lane when Word transport or host outcome handling changes.
Benchmarks remain observational; PDF export remains optional.
