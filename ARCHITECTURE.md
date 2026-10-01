# System Architecture

## Runtime map

The repository has three document consumers and one shared reconciliation
package. The Word add-in also has a local portable preparation layer and a
Word-only adapter.

```mermaid
graph TD
    Package[@ansonlai/docx-redline-js 0.8.3] --> Portable[Portable consumer core]
    Portable --> WordAdapter[Word adapter / Office.js]
    WordAdapter --> Addin[Word taskpane]
    Package --> Browser[Browser document session]
    Package --> MCP[Local DOCX MCP server]
    Gemini[Gemini client] --> Addin
    Gemini --> Browser
```

| Component | Responsibility |
| --- | --- |
| `@ansonlai/docx-redline-js` | Public DOCX facade, OOXML operations, package inspection and reconciliation. Both package manifests pin exact version `0.8.3`. |
| `src/taskpane/modules/docx-redline-js-integration/consumer-core.js` | Portable source inspection, canonical operation preparation, result handling and Flat OPC package construction. It has no Word, UI or filesystem globals. |
| `word-operation-runner.js`, `word-redline-runner.js`, `word-ooxml.js` | Word proxy reads/inserts, synchronization, native tracking management and the add-in's compatibility bridge. |
| `browser-demo/document-session.js` | Browser open/inspect/atomic edit/serialize lifecycle through the package facade. JSZip is retained for preview. |
| `mcp/docx-server/src/services/docx-document-service.mjs` | Node file I/O and session operations through the package facade; persistence occurs only on `docx_save_as`. |
| `src/taskpane/modules/chat/gemini-client.js` | Shared Gemini HTTP transport used by the taskpane, browser demo and eval harness. |

The add-in integration `index.js` reexports the package surface and existing
Word-facing exports. `word-operation-runner.js` preserves its previous named
exports while delegating pure preparation to `consumer-core.js`.

## Word batch lifecycle

`executePureOoxmlBatch` is the Word adapter for the migrated atomic operation
path:

1. Resolve the Word body/range/paragraph target and request OOXML once.
2. Synchronize the read, then call `prepareCanonicalBatch` with that immutable
   source and the operation array or source-aware operation factory.
3. Return `noop` or `refused` without an insertion, or submit the prepared
   package once and synchronize.
4. Return the engine result/receipts together with host observations such as
   `written`, `writeAttempted` and `mutationOutcome`.

The portable preparation result can be `ready`, `noop`, `refused` or `error`;
`ready` means package preparation succeeded and does not mean Word applied it.
The adapter reports an unconfirmed attempted insertion as indeterminate, never
claims rollback without host confirmation, and does not retry a mutation after
an uncertain write. Tracking management requested inside this adapter remains
Word-specific. Some older tool paths have not yet moved to this one-batch
adapter, so the single-read/single-write contract applies only to the migrated
batch path.

The redline planner in `redline-plan.js` maps AI changes against inspected
paragraphs and asserts the immutable source baseline before preparation. The
Word host separately captures that baseline for prompt context. Stale or
unavailable supplied baselines refuse before native or canonical list writes.

## Taskpane startup and tool dispatch

`taskpane.js` uses `createLazyModuleLoader` for editing support. The Word
operation module and reconciliation package load on first use, rather than in
the initial entry graph. A concurrent first request shares the same pending
promise. Import or initialization failure clears that promise so a later
request can retry.

Before capturing a Word source baseline, the chat path awaits the Word support
module after applying `Office.context.platform`. Function-call dispatch awaits
agentic tool loading and dependency initialization before entering the mutation
loop. Successful tool dependencies initialize once. The taskpane's initial
payload is smaller; first use incurs the deferred module load. Deploy the full
`dist/` output, including hashed chunks.

Dedicated startup profiling is isolated from ordinary builds. It records module
evaluation, Office readiness and usable UI with a local collector, and suppresses
automatic Glance calls only in that profiling build. See
[OOXML performance instructions](docs/ooxml-performance.md).

## Model and provider flow

The main chat call selects tools. Tools such as `apply_redlines` may then make a
separate structured-output call through `callGeminiForDiffs`, followed by
sanitization, anchor verification and document preparation. Corrective diff
generation is distinct from transport retry and is only attempted when the
document was not mutated. A tool failure is returned to the chat loop as
structured feedback.

All Gemini HTTP requests use `gemini-client.js`. It caps attempts at three and
retries transient network/timeouts and selected HTTP responses (408, 429, 500,
502, 503 and 504). Authorization/invalid-request responses and malformed JSON
are not retried. Abort handling covers fetch, JSON parsing and retry backoff.
Transport retries never rerun a document tool. Model-specific generation
settings remain in `modules/config/model-profiles.js`.

## List tools and current boundaries

Request validation checks types, ranges and supported enums before Word
execution, then checks target indexes against the live paragraph count.
`edit_list` does not clamp an invalid target, and header conversion keeps each
replacement paired with its original target when sorting. Table content is
validated against declared nested-array dimensions before the first content
write.

Canonical `insert_list_item` migration is intentionally limited to the tested
active bullet/decimal subset and its required anchors, levels and tracking mode.
Other levels or numbering formats, tracking-off requests and unsupported
anchors stay on native Word paths. The earlier 2026-09-30 v0.8.2 baseline recorded
five native fallback cases with 20 actual Office.js checks and 20 independent
Word checks on Word 16.0.20430.20092. The five cases were deep `+1` insertion,
source-level-2 outdent, UpperRoman insertion, lowerRoman insertion, and
insertion with redlining disabled while restoring prior `TrackAll`; the five
engine-reference views were marked not applicable, not counted as passes.

Current v0.8.3 verification covers 12 supported list cases with 48 actual
Office.js checks, plus those five native routes with 20 more checks (68 total).
The independent Word oracle checked 17 exports: 92 applicable checks passed,
none failed, and five were not applicable. The two original public-facade
Reject All cases also pass 12 Word checks. `npm test` passes 51 suites with
zero failures and four exclusions; validation and production builds pass. See
the [v0.8.3 report](docs/validation-reports/2026-09-30-docx-redline-v083.md),
[Office.js list reports](docs/validation-reports/2026-09-30-v083-list-officejs.json),
[native Office.js report](docs/validation-reports/2026-09-30-v083-native-officejs.json),
[Word report](docs/validation-reports/2026-09-30-v083-officejs-word.json), and
[public-facade Word report](docs/validation-reports/2026-09-30-v083-facade-word.json).
These results verify the reported cases; they do not cover every list shape or
complete canonical migration.

The v0.8.2 reports recorded Reject All paragraph-boundary failures after plain
insertion and list-range replacement, plus inspection of historical paragraph
properties as current. The v0.8.3 package passes 12 independent Word checks on
the two former public-facade Reject All failures; the current list matrix also
passes its reported Office.js and Word checks. The release notes describe the
fixes, including all-empty source ranges, historical properties, `openDocx`
list-numbering reuse, and explicit list-start parsing. The dated v0.8.2
reports retain their earlier results.

The current canonical limitations are narrower but still block full migration:
plain or text-changing header-to-list conversion and list-format changes do
not yet have supported canonical operations. On bare `document.xml`, a
marker-prefixed `1. Header` to `1. Header` operation can fail with
`RECEIPT_RECONCILIATION_FAILED`. These capabilities and the migration status are
tracked in the [canonical list migration follow-up](docs/plans/2026-09-30-canonical-list-migration-follow-up.md).
Successful native fallbacks do not complete canonical migration.

## Validation and evidence

Offline tests verify deterministic mapping, package output and mocked host
contracts. Browser checks exercise the browser session and real Chromium
workflow. The actual Office.js collector forwards calls to genuine Word proxies;
its exported DOCX is then checked by the independent desktop Word oracle.
Native COM Word checks are a separate lane. Performance reports are
observational and are not correctness gates. A report's sample counts describe
that specific host/build and do not guarantee other Word versions or documents.

Use [scripts/README.md](scripts/README.md) for commands and
[validation reports](docs/validation-reports/) for dated results. No live model
call is required by the offline suites.
