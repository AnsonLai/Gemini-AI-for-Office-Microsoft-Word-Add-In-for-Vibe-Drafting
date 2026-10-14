# Verification and profiling commands

See [consumer testing](../docs/TESTING.md), the
[tool contract](../docs/agentic-tool-contracts.md) and
[current plan](../docs/plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md).
Offline, genuine Office.js, independent Word and provider evidence cover
different boundaries. Unsupported native routes are retired.

## Offline and build gates

```powershell
npm test
npm run build:dev
npm run build
npm run validate
node scripts/run-all-tests.mjs --list
```

The runner discovers tests recursively and runs each in a separate Node process.
`DOCX_TEST_CONCURRENCY` defaults to 4; `DOCX_TEST_TIMEOUT` defaults to 180000 ms.
Manual performance, desktop fixture generation and XML setup entrypoints are
excluded. Historical package-version skips are printed separately. The former
paid redline-schema eval is retired; the direct taskpane live lane below replaces
it. Public engine issue reproducers remain under `docs/library-issues/`.

## Genuine Office.js and independent Word

Requires Windows desktop Word and existing trusted development certificates.
The validation entry is absent from normal builds. No provider calls are made.
The runner starts its own `/x` Word, uses a temporary add-in registration and
cleans up only that instance/registration. Its loopback port defaults to **3002**,
leaving development 3000 and golden 3001 separate.

```powershell
npx webpack --mode development --env WORD_HOST_VALIDATION=1 --output-path .cache/reliability/direct-dist
node scripts/run-officejs-validation.mjs --launch --dist-dir .cache/reliability/direct-dist --artifacts-dir .cache/reliability/direct-officejs
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 -FixtureManifest .cache/reliability/direct-officejs/officejs-fixtures.json -ArtifactsDir .cache/reliability/direct-oracle -SkipNativeInsert -TimeoutSeconds 600
```

The default generator is `tests/direct_document_fidelity_manifest_tests.mjs`.
It creates nineteen required capability cases (fourteen original, five list-action) and one
no-write refusal control from Word-authored sources with independent expected text, list
identity/continuation/restart, labels, formatting and table cells.
The collector runs production inspect/apply, records real transport counters,
verified/no-op/stale/refusal outcomes, and saves/reads the exact body checkpoint
through the production IndexedDB store. A separate restored export tests the
production checkpoint helper, tracking restoration and expiry of a handle
issued after the edit. Its restoration write is separate from edit counts.
Capability cases include sequential new comments on an existing root/reply/
resolved-thread source, a combined rewrite/comment batch after fresh inspection,
and untracked wording/formatting changes alongside another reviewer's pending
revision. These assert expected comment counts and original thread identities,
rather than treating refusal as a passing preservation check. The sequential
case uses two checkpoints/two edit writes and restores the latest checkpoint;
single-stage edit cases retain one checkpoint/one write.

The list-action cases use sources from `scripts/create-agentic-list-action-fixtures.ps1`
(`tests/fixtures/agentic-lists/list-*-source.docx` and
`list-actions-source-observations.json`, authored, saved and reopened in Word):
`list-convert-noncontiguous-typed-headers` (`1. Alpha`/`2. Beta`/`3. Gamma` around body paragraphs,
one `convert_to_list` each), `list-convert-letter-headers-start-at-c` (`C.`/`D.`, `upperLetter`),
`list-change-format-bullets-to-decimal`, `list-continue-second-list-after-first` (labels 3./4.) and
`list-rewrite-and-convert-changed-header` (`rewrite_paragraph` plus `convert_to_list` on one paragraph).
The oracle asserts the labels from Word's own `ListFormat.ListString` after `Revisions.AcceptAll()`/`RejectAll()`.
`list-convert-ambiguous-roman-marker-refusal` is a safety control, not a capability pass: the collector requires
the library's `LIST_MARKER_AMBIGUOUS` (in `error.failedOperations`), zero content writes, zero checkpoints and an
unchanged source; the oracle excludes it like every `expectedPrewriteRefusal` case.

The default success gate for untracked pure insertion on a revised target was
blocked by the 0.8.4 surgical-insertion defect and passes on 0.9.0 (verified
2026-10-12: all 14 default cases, 20/20 host checks, 3 checkpoint restores). Its
original applied-text, Reject All and exact native-revision expectations remain
required;
a refusal is not a capability pass. `--fixture-manifest PATH` permits an explicit
diagnostic subset. A separate manifest can request the prewrite tracking-policy
safety control, which is excluded from capability/COM totals. Always record
excluded capability gates; a passing subset is not a full-manifest pass. See the
[correction report](../docs/validation-reports/2026-10-02-lightweight-harness-wp5-correction.md).

Read `officejs-report.json`, then verify `officejs-fixtures.json` exports with
the independent Word oracle. Each applicable source/tracked/accepted/rejected
check must pass. Inapplicable engine-reference-package checks remain excluded;
the legacy oracle's generic reason mentioning native operations is not a claim
that the current tool uses native edits. Save/reopen, recognized revisions,
original list labels/bindings, unaffected formatting/comments and table cells
must be checked independently of prepared engine output.

Since 2026-10-12 the oracle (`scripts/verify-wp6-word.ps1`) resolves revisions
with `Document.Revisions.AcceptAll()/RejectAll()`, which is what the ribbon does,
instead of `Document.AcceptAllRevisions()/RejectAllRevisions()`. The latter leave
a body-final table append's tracked mark behind, a Word API quirk (KB-3 in
[`KNOWN_BUGS.md`](../docs/plans/KNOWN_BUGS.md)). Oracle reports from before that
date used the `Document.*` methods. See the
[0.9.0 validation](../docs/validation-reports/2026-10-12-docx-redline-v090.md).

For explicit cleanup after an interrupted launch:

```powershell
node scripts/run-officejs-validation.mjs --artifacts-dir .cache/reliability/direct-officejs --cleanup
```

`npm run test:word` also exercises the existing public-library fixture lane.
COM insertion checks do not establish Office.js transport. PDF export (`-Render`)
is an optional diagnostic and still needs visual review. The bounded worker
records stalls and kills only its identified owned Word process; normal user
Word sessions are preserved.

## Original sixteen-prompt NDA taskpane gate

`nda-scenario.mjs` preserves the original prompts/legacy replay. The current
`nda-direct-scenario.mjs` uses the same prompts with opaque inspection handles,
retained edits and explicit unchanged refusals. The runner defaults to direct
protocol and deterministic replay. Use `--protocol legacy` only with the frozen
legacy bundle for baseline comparison.

```powershell
node scripts/run-golden-scenario.mjs --launch --artifacts-dir .cache/golden/direct
```

The runner builds its development-only driver, serves on loopback 3001, sets
isolated origin settings and opens only its own `/x` Word. It exports each step,
scores intended content and earlier regressions, author names, model turns,
paired receipts and content writes. It cleans its registration, owned Word and
connections at completion. `--keep-word` retains the disposable instance when
explicitly requested; `--cleanup` removes its registration. `--until N` limits
the run; `--skip-build --dist-dir PATH` uses a previously frozen driver build.
Artifacts include per-step DOCX/JSON, transcript, model log and report.
`--source-diagnostics` is restricted to replay of the repository NDA fixture.

Live drafting validation requires `GEMINI_API_KEY` in the environment and incurs
provider usage. Run the complete scenario for both configured profiles before
shipping the changed tool prompt:

```powershell
# With GEMINI_API_KEY supplied securely in this shell:
node scripts/run-golden-scenario.mjs --mode live --profile fast --model gemini-flash-latest --launch --artifacts-dir .cache/golden/live-fast
node scripts/run-golden-scenario.mjs --mode live --profile slow --slow-model gemini-2.5-pro --launch --artifacts-dir .cache/golden/live-slow
```

`--profile fast` or `slow` sends every prompt (including the think-button step)
through that selected model. These names are existing defaults, not new model
recommendations; substitute the actual supported configured choices when needed.
`--profile configured` (default) uses the fast/slow buttons with `--model` and
`--slow-model` overrides. Both settings are initialized each run to prevent old
localStorage state contaminating evidence. Live receipts come only from actual
current-turn provider history; missing/truncated history stays incomplete and
fails evidence. Logging is awaited before scoring. Request URLs/headers/API
keys are excluded from model logs. Refused/uncertain batches never replay.

No live model or Word Online result may be inferred from deterministic replay.
Desktop is the initial tested host. The final plan/report names remaining
shipping gates when credentials or Online validation are unavailable.

## Observational benchmark and startup

```powershell
npm run benchmark:ooxml -- --output=.cache/ooxml-performance/report.json
npx webpack --mode development --env TASKPANE_STARTUP_PROFILE=1
node scripts/run-officejs-validation.mjs --startup --launch --artifacts-dir .cache/reliability/startup
```

The benchmark uses supported deterministic no-op, literal/rewrite, existing
list/table text and comment workloads. Direct intent compile, canonical
preparation and package facade phases remain separate observations, not
additive timings. `DOCX_BENCH_PARAGRAPHS`, `DOCX_BENCH_OPERATIONS`,
`DOCX_BENCH_WARMUPS` and `DOCX_BENCH_ITERATIONS` control workload/samples.
No model, Word or disk calls occur inside timed phases. Startup profiling
suppresses automatic Glance, measures usable UI and first deferred tool-module
load, and uses the same isolated host runner. There is no universal latency
gate. Historical measurements remain in [performance notes](../docs/ooxml-performance.md).

Reproduce the fixed physical-line production surface used in the milestones:

```powershell
node scripts/count-document-agent-code.mjs --ref '2c179a1^'
node scripts/count-document-agent-code.mjs --ref 86b6ddc
node scripts/count-document-agent-code.mjs
```

All new/replacement taskpane, command, integration and chat JS is included;
blank/comment lines and final unterminated lines count. Other utility/config
changes are identified separately, rather than inflating this net reduction.
