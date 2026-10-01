# Project Roadmap

> Current status: 2026-09-30. This roadmap summarizes shipped and verified work,
> the active canonical-list follow-up, and longer-term portability goals. The dated
> plans and validation reports linked below are the detailed source of truth.

## Current package and consumers

The Word add-in and MCP server pin the exact
`@ansonlai/docx-redline-js@0.8.3` release. The package owns supported OOXML
document operations, reconciliation and DOCX package lifecycle; consumer code
owns tool mapping, host I/O, browser UI/session flow, and MCP tool contracts.
The add-in still has a local Flat OPC bridge/mapping layer, while the browser
and MCP editing paths use the public package facade. See the [library offload
review](docs/library-offload-review.md) for the file-level inventory. A package
upgrade or consumer feature migration is not complete until corresponding
compatibility and host evidence is recorded.

## Completed plan milestones

- [September v0.8.2 upgrade](docs/plans/completed/2026-09-09-docx-redline-js-v0.5.4-upgrade.md):
  the iterative Word redline loop, native insertion recovery and manual table
  synthesis were replaced on the migrated redline path by one atomic library
  batch and one confirmed insertion; actual Office.js and independent Word
  checks are recorded. MCP now uses `DocxDocument` sessions; four redundant
  services and its direct JSZip/xmldom dependencies were removed.
- [Reliability and quality gates](docs/plans/completed/2026-08-29-reliability-and-quality-gates.md):
  Gemini transport/retry behavior and structured mutation outcomes are
  centralized; host writes are not replayed after uncertain outcomes.
- [Package boundaries and integrations](docs/plans/completed/2026-08-29-package-boundaries-and-integrations.md):
  portable preparation, the browser document-session facade, MCP sessions,
  and cross-host semantic parity are validated. Browser editing and save use
  the package facade; JSZip remains only for `docx-preview`.
- [OOXML performance](docs/plans/completed/2026-08-29-oxml-engine-and-performance.md):
  dated v0.8.2 consumer, browser, package, Word, and startup measurements are
  recorded. The observed 1,000-paragraph workload is accepted; no universal
  timing threshold or speculative engine rewrite was selected. Deferred
  startup and the measurement lanes remain local consumer work.
- [Original agentic tools and list reliability plan](docs/plans/completed/2026-08-29-agentic-tools-and-list-reliability.md):
  its completed inventory, supported-route work, and v0.8.3 fidelity checks are
  recorded. Remaining canonical migration is deferred to the [new follow-up
  plan](docs/plans/2026-09-30-canonical-list-migration-follow-up.md); closing
  the original plan does not mark that migration complete.

These milestones close their stated scopes. They do not imply that every
agentic list command has a supported canonical operation.

## Active work: canonical list migration and fidelity

The [canonical list migration follow-up](docs/plans/2026-09-30-canonical-list-migration-follow-up.md)
remains open. The original agentic list plan is closed with its remaining
canonical work carried forward. Contract inventory and stale-source validation are complete, and
a supported bullet/decimal insertion subset with tracking enabled at
source/resolved levels 0–1 is verified. General `edit_list` and header
conversion remain on native/legacy paths; deeper or other-style routes and
tracking-off requests also remain native. Unsupported inputs refuse before
mutation. The full canonical migration is pending capability work and
validation of the remaining operations. This plan stays open even when current
live validation gates pass.

The v0.8.2 library follow-up reports recorded these earlier defects and gaps:

- [Plain insertion Reject All leaves an empty paragraph](docs/library-issues/2026-09-30-list-insertion-rejection.md).
- [List range Reject All merges source paragraphs](docs/library-issues/2026-09-30-list-range-rejection.md).
- [Unmarked or text-changing header conversion lacks a canonical mapping](docs/library-issues/2026-09-30-canonical-list-operations.md).
- [Historical list properties are inspected as active numbering](docs/library-issues/2026-09-30-historical-list-inspection.md).

The v0.8.3 release notes report fixes for the paragraph-boundary Reject All
failures, historical paragraph-property inspection, `openDocx` list-range
numbering, and explicit list-start parsing. Current validation covers 12
supported list cases (48 Office.js checks), five native routes (20 Office.js
checks), and 17 exports checked by the independent Word oracle (92 applicable
checks passed, zero failed, five not applicable). The original two
public-facade Reject All cases pass 12 Word checks. `npm test` passes 51 suites
with zero failures and four exclusions; validation and production builds
pass. The original dated reports retain their v0.8.2 evidence. Remaining
limits are that plain or text-changing header-to-list conversion and
list-format changes lack supported canonical operations, and a
marker-prefixed `1. Header` to `1. Header` operation on bare `document.xml` can
fail with `RECEIPT_RECONCILIATION_FAILED`. See the [v0.8.3 report](docs/validation-reports/2026-09-30-docx-redline-v083.md)
for the current run status and linked artifacts.

## Longer-term direction

- Keep OOXML as the editing path where the package supports the operation;
  retain Word APIs for host transport and feature paths that lack a safe
  canonical operation.
- Implement canonical operations for the remaining capabilities, then migrate
  list paths only when accepted/rejected views and package preservation are
  verified.
- Treat deeper Word Online fidelity, broader host decoupling, and splitting the
  add-in, browser demo, and MCP server into separate repositories as future
  work, not completed milestones.
