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

Full canonical list migration is in the
[active follow-up](plans/2026-09-30-canonical-list-migration-follow-up.md).
Report engine defects to the library and retain consumer regressions. Provider
evaluation is opt-in; PDF export is optional diagnostic evidence.
