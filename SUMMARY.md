# Change Summary

> Role: **Session summary log** of what happened and what changed.

## [2026-10-01] Table incident checkpoint

The user reported a stalled underline-then-table request. Consumer fixes now
refresh document context after confirmed mutation exchanges, preserve real
requests and complete tool pairs in bounded history, retain failed-mutation
counts across read-only tools, and show a terminal status when the model-turn
budget expires. Redline planning narrowly coalesces a compatible last-paragraph
edit and one table append, keeps the append anchor literal, maps both source
changes to one receipt, and refuses tracked inline-formatting combinations
before writing. The 54-suite offline aggregate and production build pass. The
final desktop Word run passed 13 of 20 checks: plain accepted table content and
both native insertions pass; tracked Reject All leaves an extra paragraph,
while Accept All loses underline in the formatting fixture. See the
[incident plan](docs/plans/2026-09-30-table-creation-reliability.md),
[host report](docs/validation-reports/2026-10-01-table-creation-reliability.md),
[tracked-formatting report](docs/library-issues/2026-10-01-table-append-tracked-formatting.md),
and [Reject All paragraph report](docs/library-issues/2026-10-01-word-reject-table-append-paragraph.md).
No full live chat/model reproduction is claimed.

## Current package status: v0.8.3

The add-in, MCP server and browser demo use the public
`@ansonlai/docx-redline-js@0.8.3` package. Its release notes describe fixes for
tracked paragraph-boundary behavior, historical paragraph-property
inspection, list-range numbering through `openDocx`, and explicit list-start
parsing. Current consumer and Word validation is recorded in the [v0.8.3
report](docs/validation-reports/2026-09-30-docx-redline-v083.md); the dated
v0.8.2 reports below remain historical evidence.

Canonical list migration remains open. Plain or text-changing header-to-list
conversion and list-format changes still lack supported canonical operations.
A marker-prefixed `1. Header` to `1. Header` operation on bare `document.xml`
can fail with `RECEIPT_RECONCILIATION_FAILED`. The original list plan is closed
with these remaining capabilities transferred to the [canonical migration
follow-up](docs/plans/2026-09-30-canonical-list-migration-follow-up.md), which
remains open. All five original plans are closed; this follow-up carries the
uncompleted canonical migration work forward without marking it complete.

The current ownership split is documented in the [library offload
review](docs/library-offload-review.md): the library owns supported DOCX
operations, reconciliation and package lifecycle; consumers retain mapping,
host transport, reliability and UI/session code. The add-in's redline route is
now a single atomic library batch with a local Flat OPC/Office.js bridge. MCP
uses `DocxDocument` sessions and no longer carries its four redundant package
services or direct JSZip/xmldom dependencies. Browser editing uses the package
facade; JSZip remains for preview. Canonical list insertion is verified only
for tracking-on bullet/decimal levels 0–1; broader list paths remain local.

## [2026-09-30] v0.8.2 Consumer Boundaries, Reliability, and Performance

### What Changed

- The add-in and MCP server now pin exact `@ansonlai/docx-redline-js@0.8.2`.
- Four dated plans closed: the September package upgrade, reliability gates,
  package boundaries/browser/MCP integration, and OOXML performance/startup
  work. Their detailed closure evidence is under `docs/plans/completed/` and
  `docs/validation-reports/`.
- Portable operation mapping/preparation is kept separate from Word host I/O.
  The browser demo edits through the public document-session facade; MCP uses
  public document sessions. JSZip remains only for the browser preview.
- Current-version benchmarks cover no-op, localized multi-edit, full-paragraph
  rewrite, existing nested-list/table text edits, and comments. They separately
  report mapping, baseline capture, inspection, browser projection, core,
  package open/apply/save, and actual Word transport. The 1,000-paragraph
  workload is accepted; no universal latency threshold was added.
- Taskpane optional module loading reduced the measured initial JavaScript
  payload; actual Word validation and the full-size/first-use tradeoff are
  recorded in the performance report.
- Native list validation now covers deeper insertion/outdent, upper/lower Roman
  numbering and tracking-off insertion with prior-mode restoration: 20 actual
  Office.js checks and 20 independent Word checks pass. The final offline
  aggregate passes 50 suites, zero failures, with four exclusions.

### Open Work at This v0.8.2 Checkpoint

The [original agentic tools and list reliability plan](docs/plans/completed/2026-08-29-agentic-tools-and-list-reliability.md)
remained open. The supported insertion subset and its validation evidence did
not close the full canonical list migration. Four upstream reports covered
insertion Reject All leaving an empty paragraph, list-range Reject All merging
paragraphs, missing canonical header-conversion mapping, and historical list
properties treated as active numbering. At that dated checkpoint, no library
fix release was available; see the current v0.8.3 status above.

Current posture and verification records are summarized in
[STATE.md](STATE.md), [ROADMAP.md](ROADMAP.md), and the linked dated plans.

## [2026-02-13] Documentation Consolidation (GSD Transition)

### What Happened
Consolidated fragmented architecture and usage documentation into a structured GSD (Get Shit Done) methodology.

### Major Changes
- **SPEC.md**: Updated with host-agnostic vision, standalone usage examples, and the **multi-project splitting vision**.
- **ARCHITECTURE.md**: Created a single source of truth for the reconciliation engine, including a **Project Map** that defines the relationships between `Core`, `Word Add-in`, `Browser Demo`, and `MCP Server`.
- **ROADMAP.md**: Now tracks clear migration phases (1, 2, 3) and a final **Repository Split** milestone.
- **Sub-Project Alignment**: Updated READMEs for `browser-demo`, `mcp`, and `src/taskpane` to reference the central GSD structure.
- **Project Refinement**: Integrated the "Thin Runtime" and "Logic Inward" principles into the core documentation to enforce the host-agnostic vision.
- **New GSD Files**: Introduced `STATE.md`, `PLAN.md`, and `SUMMARY.md` for better task tracking and project memory.
- **Cleanup**: Deprecated 4 redundant documentation files.

### Why This Matters
Reduces context overhead for AI agents and provides a clear "Mission Control" center for the project, ensuring the portability vision is never lost during feature development.
