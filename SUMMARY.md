# Change Summary

> Role: **Session summary log** of what happened and what changed.

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

### Current Open Work

The [agentic tools and list reliability plan](docs/plans/2026-08-29-agentic-tools-and-list-reliability.md)
remains open. The supported insertion subset and its validation evidence do
not close the full canonical list migration. Four upstream reports cover
insertion Reject All leaving an empty paragraph, list-range Reject All merging
paragraphs, missing canonical header-conversion mapping, and historical list
properties treated as active numbering. No library patch or unreleased fix is
claimed; complete migration remains pending release and downstream validation.

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
