# Published 0.8.3 consumer upgrade and list validation

Both root and MCP package manifests/lockfiles pin exact
`@ansonlai/docx-redline-js@0.8.3`. Upgrade and regression checkpoint: `148c96b`.
API/schema and CLI contract remain unchanged (contract version 8).

## Results

| Lane | Result | Evidence |
| --- | --- | --- |
| Offline aggregate | 51 suites passed, zero failed; four excluded entrypoints | `npm test` |
| Supported list matrix | 12 cases pass exact accepted/rejected text, numbering identities and untouched bold formatting | `tests/agentic_list_fidelity_tests.mjs` |
| Actual Office.js list matrix | 48 checks passed | [Collector](2026-09-30-v083-list-officejs.json) |
| Actual Office.js native matrix | 20 checks passed across five cases | [Native collector](2026-09-30-v083-native-officejs.json) |
| Combined Office.js exports, independent Word | 92 checks passed, zero failed; five engine-reference entries not applicable | [Word oracle](2026-09-30-v083-officejs-word.json) |
| Complete DOCX facade, original diagnostic cases | 12 independent Word checks passed | [Facade oracle](2026-09-30-v083-facade-word.json) |
| Build | Development validation entry and ordinary production build pass | webpack |

Office.js host: Word `16.0.20430.20092` (PC). COM oracle reports build
`16.0.20430`. The two former library-defect cases are now passing matrix cases,
not expected failures. Single-empty/all-empty source paragraph restoration and
public-facade same-kind numbering are permanent focused regressions in
`tests/docx_redline_v083_compat_tests.mjs`. The historical-inspection cutover test
now requires historical numbering to be absent and verifies the native plain
paragraph path. The historical 0.5.4 suite still reports its version-specific
skip separately from its passing consumer guards.

The first Office.js attempt exposed a harness expectation mismatch: the newly
included plain-anchor case still uses production native insertion, while the
default assertion expected a canonical body write. Declaring its native route
fixed the test. Production routing and library source were not modified.

The twelve-case host matrix includes eight canonical production insertions,
one native plain-anchor insertion and three candidate range/header operations.
The five additional native cases cover deep insertion/outdent, upper/lower
Roman numbering and tracking disabled with prior tracking restoration. Word
verifies source/tracked/Accept All/Reject All and, for canonical fixtures,
separately resolved engine packages. Five native engine-reference entries are
explicitly not applicable and are not counted as passes.

The facade oracle reruns both original defects through installed `openDocx`.
It verifies restored paragraph boundaries, source list types and untouched
continuation values. These checks resolve the separate public-facade numbering
report, which previously failed four Word checks.

Both local collectors stopped and their temporary add-in registrations were
removed. No paid provider calls or PDF exports ran. Prior 0.8.2 performance and
startup samples remain historical measurements; they were not relabeled as
0.8.3 measurements. Optional telemetry network failure did not fail either
build; production retains its deferred-chunk size warning.

## Reproduction

```powershell
npm test
node tests/agentic_list_fidelity_tests.mjs --export-host-dir .cache/v083/lists
node tests/agentic_list_native_fallback_manifest_tests.mjs --export-dir .cache/v083/native
npx webpack --mode development --env WORD_HOST_VALIDATION=1
node scripts/run-officejs-validation.mjs --fixture-manifest .cache/v083/lists/manifest.json --artifacts-dir .cache/v083/officejs-final --launch
# After success, stop the collector and remove the registration:
node scripts/run-officejs-validation.mjs --artifacts-dir .cache/v083/officejs-final --cleanup
node scripts/run-officejs-validation.mjs --fixture-manifest .cache/v083/native/manifest.json --artifacts-dir .cache/v083/native-officejs --launch
# After success, stop this collector and remove its registration:
node scripts/run-officejs-validation.mjs --artifacts-dir .cache/v083/native-officejs --cleanup
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 -FixtureManifest .cache/v083/officejs-final/officejs-fixtures.json -ArtifactsDir .cache/v083/list-word -SkipNativeInsert -TimeoutSeconds 600
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 -FixtureManifest .cache/v083/native-officejs/officejs-fixtures.json -ArtifactsDir .cache/v083/native-word -SkipNativeInsert -TimeoutSeconds 360
npm run build
```

This run combined the two collected manifests for one oracle invocation; the
separate commands above reproduce the same 72 supported and 20 native checks.
Keep the prepared deep/Roman source exports when verifying native cases.

## Closure and transferred work

The release's Known Limitations explicitly retain unsupported canonical
plain-to-list conversion and list-format changes. The bare-XML unchanged-header
receipt error also reproduces; packaged equivalents checked during review
succeed. The library's canonical-list plan proposes those missing features,
but 0.8.3 does not implement them.

At the user's explicit request, the [August agentic plan](../plans/completed/2026-08-29-agentic-tools-and-list-reliability.md)
is closed and archived with a deferred-scope note. WP1 and current-route WP3
are complete; the remaining WP2 work and full canonical migration acceptance
criterion transfer to the fresh [Canonical List Migration Follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md).
Release fidelity fixes do not automatically migrate production command paths.
The fresh plan owns capability implementation/release, remaining consumer
migrations and corresponding host checks, and helper retirement after all
callers migrate. This closes the old plan without claiming full migration.
