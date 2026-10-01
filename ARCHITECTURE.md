# System Architecture

## Runtime map

The repository has three document consumers and one shared reconciliation
package. The Word add-in also has a local portable preparation layer and a
Word-only adapter.

```mermaid
graph TD
    Package[@ansonlai/docx-redline-js 0.8.2] --> Portable[Portable consumer core]
    Portable --> WordAdapter[Word adapter / Office.js]
    WordAdapter --> Addin[Word taskpane]
    Package --> Browser[Browser document session]
    Package --> MCP[Local DOCX MCP server]
    Gemini[Gemini client] --> Addin
    Gemini --> Browser
```

| Component | Responsibility |
| --- | --- |
| `@ansonlai/docx-redline-js` | Public DOCX facade, OOXML operations, package inspection and reconciliation. Both package manifests pin exact version `0.8.2`. |
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
anchors stay on native Word paths. Five native fallback cases pass actual
Office.js (20 checks) and independent desktop Word (20 checks) on Word
16.0.20430.20092: deep `+1` insertion, source-level-2 outdent, UpperRoman
insertion, lowerRoman insertion, and insertion with redlining disabled while
restoring prior `TrackAll`. This is evidence for those five cases on that host;
the 20 independent checks cover source, tracked, accepted and rejected views.
Five engine-reference views are marked not applicable for these native routes,
not counted as passes. This does not migrate them to canonical OOXML or cover
every list shape. The active migration plan remains open until the remaining
canonical paths are enabled by library capability and fidelity fixes. See the
[Office.js report](docs/validation-reports/2026-09-30-agentic-native-list-officejs.json)
and [independent Word report](docs/validation-reports/2026-09-30-agentic-native-list-word.json).

The pinned 0.8.2 reports document these open constraints:

- Reject All after plain-anchor insertion can leave an extra empty paragraph.
- Reject All after list-range replacement can merge original paragraphs.
- Some unmarked/text-changing header or list-format conversions lack a
  verified canonical mapping.
- Historical `numPr` in revision history can be mistaken for active numbering;
  the observed route refuses before writing.

These are tracked in [`docs/library-issues/`](docs/library-issues/README.md) and
the [active list reliability plan](docs/plans/2026-08-29-agentic-tools-and-list-reliability.md).
Do not treat successful native fallbacks as completion of canonical migration.

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
