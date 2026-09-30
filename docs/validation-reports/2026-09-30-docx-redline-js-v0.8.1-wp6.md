# v0.8.1 upgrade: WP6 validation

Date: 2026-09-30. Root add-in and MCP server pin exact `0.8.1`.

**Historical record:** The dependency pins and two open defect findings below are
superseded by the [v0.8.2 validation report](2026-09-30-docx-redline-js-v0.8.2.md).
Both defects now pass required offline and live Word regressions on v0.8.2.

**Status: offline and live native Word gates implemented and passing for supported
cases; two library defects filed separately; actual Office.js transport remains
unverified. PDF export is optional and the measured larger workload is accepted.**

## Results

| Lane | Result | Scope |
| --- | --- | --- |
| `npm test` | 33 suites passed, 0 failed | Offline consumer suites, MCP stdio workflow and reviewed golden guardrail |
| `npm run build:dev` | Passed | Webpack; no Node polyfill errors |
| `npm run build:prod` | Passed with three bundle-size warnings | `taskpane.js` approximately 745 KiB |
| Live Word no-repair/revision/native lane | 45 checks passed | Six fixtures × six checks, two insertion roundtrips × four checks, plus invalid-content-type negative control |
| Word PDF export (optional) | Earlier attempt timed out | Not an upgrade gate; no PDF or page-appearance certification |
| Native Word package insertion | Passed | Real Word scope export → production bridge → native insertion → save/reopen; actual Office.js transport remains untested |
| Strict hyperlink boundary probe | Confirmed by live Word | Plain punctuation becomes clickable after Accept All; filed as library issue #3 |
| Strict tab/break lifecycle probe | Confirmed by live Word | Word's own Reject All reorders text around the break; filed as library issue #4 |

The offline command reports four excluded entrypoints: live Gemini evaluation,
XML provider setup, the old observational performance harness, and Word Desktop
fixture generation. The historical v0.5.4 suite passes its consumer guards and
reports its version-specific package behavior as skipped. The two strict probes
are optional arguments to the fidelity suite; their failures are open gates and
are **not** represented by the successful default invocation.

Builds print blocked Application Insights telemetry errors but exit successfully.

## Methodology and fixtures

Followed the local `C:\Users\Phara\Desktop\Projects\Docx Redline JS` project's
`docs/TESTING.md`, visual regression tests, comment-thread fixture tests and Word
COM oracle scripts. XML contracts, Word differential checks, rendering review and
performance measurements are separate evidence.

`tests/fixtures/wp6/README.md` records three byte-identical Word-authored source
DOCX fixtures, their generators and SHA-256 hashes. A small synthetic fixture
provides bold/plain/italic runs, tabs, a break, a hyperlink, portrait/landscape
sections and columns. The default suite verifies:

- Localized bold/italic inheritance without plain-run formatting bleed; unchanged
  tab/break/hyperlink/section structures and unaffected styles/relationships.
- Exact accepted/rejected body text; atomic refusal with zero writes for the
  unsupported combined structural shape.
- Multi-paragraph root identity, reply markers, resolve/reopen and cascade delete.
- Optional `commentsIds`/`commentsExtensible` identities and synchronization.
- Actual add-in Flat OPC sibling synchronization and repair of the old
  `commentsExtended` content type during a successful body edit.
- Localized first-page footer editing in a Word-authored package; exact
  accept/reject part text, all six parts discovered, PAGE field and all other
  package parts preserved byte-for-byte.
- A plain-body bridge insertion case with independent literal expectations,
  isolating native import from comment-thread handling.
- Word COM assertions for actual bold/italic state, no plain-run formatting
  bleed, Calibri/12-point inheritance, explicit tab position/alignment, section
  orientation and column counts. These complement XML assertions without claiming
  page appearance. Word exposes default tab stops as well as the explicit stop.

The bridge now uses the package's public sibling reconciler and part registry.
The consumer avoids maintaining a second implementation of comment identities,
relationship types and content types.

## Word Desktop evidence

Word version `16.0`, build `16.0.20326`. Final command:

```powershell
npm run test:word -- -ArtifactsDir .cache/wp6/live-complete
```

Each of six supported fixtures was opened as source and tracked output, resolved with
Word's own Accept All and Reject All, and opened again as independently
engine-resolved accepted/rejected packages. Exact body text and zero remaining
revisions were checked for resolved states. Engine-resolved files must have no
residual body/part revisions **before** Word performs any resolution, avoiding a
second resolver hiding incomplete engine output. Body-edit cases must contain
Word-recognized revisions. The footer case additionally checks
first-page footer text through Word's COM model. Comment cases compare counts,
unique author/text identities, parent identities and resolved state; enumeration
order differs between package inspection and Word and is not the oracle.

Two native insertion cases cover a plain replacement and a threaded-comment
reply. The harness exports the source's actual `Content.WordOpenXML`, invokes
`scripts/apply-live-word-ooxml.mjs` to call the production `executePureOoxmlBatch`
with the exact fixture operations, requires one successful payload write, inserts
with Word tracking disabled, saves to a new file, and reopens tracked/accepted/
rejected views. The reply remains attached to its root, the resolved payment
thread remains resolved, all four comments remain present, and the plain edit
retains recognized revisions and exact Accept All/Reject All text.

Word rejected a deliberately corrupted `commentsExtended` content type without
repair. This negative control verifies that the no-repair check can detect the
release's relevant package defect.

Reports and generated outputs are ignored under `.cache/wp6/`. The original
native timeout was attributed incorrectly to `InsertXML`: the progress marker
covered the following save too. Finer instrumentation showed insertion returned
and **SaveAs2 stalled**. Matching upstream's correctly boxed by-reference
`SaveAs2` arguments fixes the harness and both native roundtrips pass. An unchanged
Word-generated package also imported successfully as an independent control.

The earlier `ExportAsFixedFormat` attempt and a plain new document's PDF export
timed out. At the user's direction this is optional diagnostic evidence, not a
product feature or upgrade gate. The harness defaults to no PDF export; `-Render`
requests it explicitly.

The harness supervises a hidden worker, records partial results, fails on timeout
and cleans up only a newly created Word PID with a matching creation time.
Native insertion can be skipped explicitly for the differential lane, and the
report records skips. COM `InsertXML` and Office.js `insertOoxml` are distinct
transports. The live production bridge/native parser check does not certify
Office.js dispatch or page appearance. See [Microsoft's insertion guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml).

## Separate library issues

Both defects reproduce against installed 0.8.1 and clean reference source HEAD
`c4db17a8a622852c0596ec8725573ac0101d9a09`, version 0.8.1. The closest localized
replacement, structural tab/field and insertion affinity suites pass, showing
coverage gaps. No library-side fix or consumer XML workaround was introduced.

- [Issue #3: hyperlink punctuation boundary](https://github.com/AnsonLai/docx-redline-js/issues/3):
  Word accepts correct body text `example.net.` but reports hyperlink Range.Text
  as `example.net.` instead of `example.net`. The period changes clickable scope.
- [Issue #4: line-break rejection order](https://github.com/AnsonLai/docx-redline-js/issues/4):
  replacing `tabbed` with `aligned` accepts correctly, but Word's own Reject All
  returns `\ttabbedLine\n with ` instead of source `\ttabbed\nLine with `. Word
  confirms this is an emitted-revision problem, not only a resolver discrepancy.

The issue bodies include the portable
`docs/library-issues/2026-09-30-fidelity-reproducer.mjs`. Its synthetic one-paragraph
cases require independent outcomes and exit 1 until corrected; set
`DOCX_REDLINE_SOURCE_ROOT` to test source instead of the installed package.
The exported known-defect cases retain expected and observed engine states
separately. `npm run test:word -- -IncludeKnownDefects -SkipNativeInsert` checks
them in real Word and currently exits 1: the two defective resolved states each
fail in Word's own resolution and the independently engine-resolved package.

## Golden baseline review

The previously stored baseline combined old in-repository engine output with a
partial v0.5.4 refresh. Its original provenance is commits `b76167a`/`da00b7f`;
commit `0454172` refreshed only format-add, mixed-edit and text-to-table hashes.
It was not a complete v0.5.4 baseline.

Replayed the exact nine cases against upstream tag `v0.5.4`, commit
`2871bca73d590bd626518d9c419aff81254a4219`, extracted read-only into `.cache/wp6/`.
**All nine raw XML outputs and hashes match installed v0.8.1 exactly.** Also
replayed the historical in-repository engine to inspect raw changes:

| Case | Reviewed difference from stored/historical output |
| --- | --- |
| List generation | Source paragraph retained with deletion of its paragraph mark; inserted marks on three list paragraphs; removed redundant blank paragraph; nesting remains 0/1/0 |
| Table reconciliation | Repeating header (`w:tblHeader`) and `w:cantSplit` on three rows; +80 characters |
| Text to table | Same +80 table-property characters relative to the partially refreshed stored baseline; older historical output also lacked paragraph-mark tracking |
| Comments | `w14:paraId` and namespace on each comment paragraph; document anchors unchanged; +87 characters per comment |

Refreshed baseline/latest after this review. Hash equality establishes unchanged
output across the package upgrade for these cases; it does not certify every
pre-existing engine behavior. In particular, the engine's text-to-table Reject
All helper retains empty table-cell paragraphs, which has not received Word
rendering validation here. `--verify` now leaves tracked latest output untouched;
`--export-dir ... --export-only` supports future raw XML review.

## Performance

Node `v24.11.1`, Windows x64, three warmups and 15 measured iterations; ten
independent localized tracked edits per batch. Full samples and workload metadata
are in `scripts/ooxml-benchmark-latest.json`.

| Paragraphs | Core batch median / p95 | Open median | Save median | Open/apply/save median |
| ---: | ---: | ---: | ---: | ---: |
| 100 | 35.610 / 39.126 ms | 0.258 ms | 0.285 ms | 42.471 ms |
| 1,000 | 296.499 / 393.545 ms | 0.493 ms | 1.309 ms | 357.329 ms |

These measurements exclude Word, disk, model and transport latency. The user
accepts the slower 1,000-paragraph workload; it does not block the upgrade.
The benchmark remains observational; `--require-sub100ms` is an explicit optional
budget gate, not the current acceptance requirement.

## Remaining gates

1. Resolve the two separately filed library issues; rerun their strict and live
   Word probes against the corrected library release.
2. Validate the actual Office.js single-hop write in Word Desktop, especially
   commented documents. Native package insertion now passes its own lane.

PDF export is optional diagnostic evidence. The observed performance is accepted.

WP6 must not be marked fully complete while these fidelity/host gates remain open.
