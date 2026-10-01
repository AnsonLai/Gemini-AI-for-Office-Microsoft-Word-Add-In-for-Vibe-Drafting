# Project Roadmap

> Current status: 2026-09-30. This roadmap summarizes shipped and verified work,
> the active agentic-list plan, and longer-term portability goals. The dated
> plans and validation reports linked below are the detailed source of truth.

## Current package and consumers

The Word add-in and MCP server pin the exact
`@ansonlai/docx-redline-js@0.8.2` release. The package owns supported OOXML
document operations; consumer adapters own Word host I/O, browser session
lifecycle, and MCP tool/session contracts. A package upgrade or a consumer
feature migration is not complete until the corresponding compatibility and
host evidence is recorded.

## Completed plan milestones

- [September v0.8.2 upgrade](docs/plans/completed/2026-09-09-docx-redline-js-v0.5.4-upgrade.md):
  the redline path uses one atomic batch and one confirmed insertion; actual
  Office.js and independent Word checks are recorded.
- [Reliability and quality gates](docs/plans/completed/2026-08-29-reliability-and-quality-gates.md):
  Gemini transport/retry behavior and structured mutation outcomes are
  centralized; host writes are not replayed after uncertain outcomes.
- [Package boundaries and integrations](docs/plans/completed/2026-08-29-package-boundaries-and-integrations.md):
  portable preparation, the browser document-session facade, MCP sessions,
  and cross-host semantic parity are validated. JSZip remains for preview.
- [OOXML performance](docs/plans/completed/2026-08-29-oxml-engine-and-performance.md):
  current v0.8.2 consumer, browser, package, Word, and startup measurements are
  recorded. The observed 1,000-paragraph workload is accepted; no universal
  timing threshold or speculative engine rewrite was selected.

These milestones close their stated scopes. They do not imply that every
agentic list command has a supported canonical operation.

## Active work: canonical list migration and fidelity

The [agentic tools and list reliability plan](docs/plans/2026-08-29-agentic-tools-and-list-reliability.md)
remains open. Contract inventory and stale-source validation are complete, and
a supported bullet/decimal insertion subset is verified. Remaining list
commands and unsupported formats/levels still use existing native paths or
refuse before mutation. The full canonical migration is pending upstream
capability/fidelity work and its subsequent consumer and Word validation. This
plan stays open even when the current live validation gates pass.

Four separately tracked library follow-ups currently constrain that broader
migration:

- [Plain insertion Reject All leaves an empty paragraph](docs/library-issues/2026-09-30-list-insertion-rejection.md).
- [List range Reject All merges source paragraphs](docs/library-issues/2026-09-30-list-range-rejection.md).
- [Unmarked or text-changing header conversion lacks a canonical mapping](docs/library-issues/2026-09-30-canonical-list-operations.md).
- [Historical list properties are inspected as active numbering](docs/library-issues/2026-09-30-historical-list-inspection.md).

These reports describe upstream library defects or missing capabilities. No
unreleased library fix is assumed, and no consumer workaround is being treated
as a substitute for a supported canonical operation. The current run status
and any remaining live gates are in the
[validation report](docs/validation-reports/2026-09-30-agentic-tools-and-list-reliability.md).

## Longer-term direction

- Keep OOXML as the editing path where the package supports the operation;
  retain Word APIs for host transport and feature paths that lack a safe
  canonical operation.
- After upstream list fixes/capabilities are released, update the exact package
  pin, rerun diagnostic and supported matrices, and migrate remaining list
  paths only when accepted/rejected views and package preservation are
  verified.
- Treat deeper Word Online fidelity, broader host decoupling, and splitting the
  add-in, browser demo, and MCP server into separate repositories as future
  work, not completed milestones.
