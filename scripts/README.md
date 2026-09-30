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
