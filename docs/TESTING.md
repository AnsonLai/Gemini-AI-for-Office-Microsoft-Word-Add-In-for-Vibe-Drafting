# Consumer testing and library integration

The shared library owns supported document operations and package semantics.
This repository verifies how its consumers map requests, preserve source
context, transport edits and report results. See the
[library offload review](library-offload-review.md) for the measured ownership
changes and [verification commands](../scripts/README.md) for runnable commands.

| Lane | What it verifies | Limit |
| --- | --- | --- |
| `npm test` | Library compatibility, localized replacements, immutable-source mapping, no-write refusal, package preservation, MCP workflows, browser session behavior, mocked Word outcomes and golden semantics | Offline mocks do not certify Word transport. Exclusions and version-specific skips are reported separately. |
| Actual Office.js collector | Production tools and adapters call genuine Word proxies; counts reads/writes and records tracking and outcomes | Prepared engine output alone does not prove host acceptance; export the document for independent checks. |
| Independent desktop Word worker | Save/reopen without repair, recognized revisions, literal source/Accept All/Reject All text, numbering and preserved formatting/comments | Assertions cover specific fixtures and Word builds, not every document shape. |
| Golden scenario (`scripts/run-golden-scenario.mjs`) | The owner's 16-prompt NDA session through the real taskpane UI, chat loop and tools in desktop Word; replayed (deterministic) or live model; per-step accepted-view checks, regressions and redline author attribution | One fixture and scenario; replay pins model output, live runs vary. Exports are scored offline, not by Word's own Accept/Reject. |
| Browser validation | Real open/edit/download/reopen and cross-host package semantics | Preview is not a Word fidelity oracle. |
| Benchmarks/startup profiling | Library/core/package phases, Word transport and deferred module loading | Dated observations are not correctness gates or universal latency promises. |

Canonical library list insertion and retained native/legacy routes need distinct
expectations. A passing native fallback does not validate an unsupported
canonical conversion. Keep exact expected text and source list observations
independent of the engine being tested. Use frozen Word-authored fixtures and
preserve evidence when expanding a supported route.

The [0.8.3 validation record](validation-reports/2026-09-30-docx-redline-v083.md)
records 51 passing offline suites, four exclusions, 68 actual Office.js checks,
92 independent Word checks on their exports, and 12 separate facade Word
checks. Five inapplicable engine-reference checks are excluded from passes.
These are existing dated results; the documentation review did not rerun them.
The [0.8.4 validation](validation-reports/2026-10-02-docx-redline-v084.md) report records the later 0.8.4 upgrade.

## 2026-10-01 table-append incident

The 2026-10-01 incident's focused regressions and offline aggregate pass 54 suites;
the production build passes with the existing asset-size warning. These are
separate, newer results and do not revise the dated 51-suite v0.8.3 record
above. A deterministic installed-package reproduction confirms that the
coalesced operation can construct a table with nine cells. The consumer refuses
the two reproduced tracked inline-formatting combinations before writing with
`UNSUPPORTED_TABLE_FORMATTING`; see the separate
[library report](library-issues/completed/2026-10-01-table-append-tracked-formatting.md).

Desktop Word validation on Word 16.0 build 16.0.20430 completed 20 checks: 13
passed and 7 failed. Plain accepted DOCX and COM native insertion preserve the
three-by-three table, nine nouns, paragraph text, and required trailing blank.
Tracked Reject All leaves an eighth empty paragraph after the seven source
paragraphs. The formatting fixture's Accept All loses underline in both
engine-accepted and native Word views, and its Reject All also leaves an extra
empty paragraph. Both native COM `InsertXML` insertion checks pass. Office.js
`insertOoxml` transport and the actual Office.js collector were not exercised.
Full results and evidence links are in the
[table incident validation report](validation-reports/2026-10-01-table-creation-reliability.md).
The separate library findings are recorded in the [tracked-formatting report](library-issues/completed/2026-10-01-table-append-tracked-formatting.md)
and [Reject All paragraph report](library-issues/2026-10-01-word-reject-table-append-paragraph.md).
No full live chat/model reproduction or live provider call was performed.

The subsequent Undo/stale-context regression proves refusal before any write,
refresh without batch replay, and a newly planned edit against the restored
document. It also retains refusal on a paragraph-ID-only fingerprint mismatch
and verifies safe mismatch reasons. These deterministic tests do not establish
that Word changed paragraph IDs in the user's failure.

Full canonical list migration is in the
[active follow-up](plans/2026-09-30-canonical-list-migration-follow-up.md).
Report engine defects to the library and retain consumer regressions. Provider
evaluation is opt-in; PDF export is optional diagnostic evidence.
