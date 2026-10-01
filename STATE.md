# Project State: Gemini Word Add-in

> Current technical posture: 2026-09-30. Detailed implementation and test
> evidence lives in the [dated plans](docs/plans/) and
> [validation reports](docs/validation-reports/).

## Current Posture

- **Table incident (2026-10-01):** confirmed writes now refresh the next chat
  iteration's prompt text and canonical baseline. Bounded history retains real
  requests and complete tool exchanges; read-only calls do not reset failed
  mutation budgets, and loop exhaustion reports a terminal status. Table append
  mapping and a separately reproduced library formatting limitation are still
  being verified in the [incident plan](docs/plans/2026-09-30-table-creation-reliability.md).

- **Library:** the Word add-in and MCP server pin exact
  `@ansonlai/docx-redline-js@0.8.3`. Browser editing uses the same package
  through its public document-session facade.
- **Host boundaries:** portable mapping and OOXML preparation are separated
  from Word transport. The library owns supported document operations,
  reconciliation and package lifecycle; add-in `consumer-core.js` retains
  tool mapping and Flat OPC preparation/assembly. The Word adapter owns
  `getOoxml`, `insertOoxml`, tracking-mode changes, and confirmed/indeterminate
  write outcomes. Browser editing uses a package document session; MCP uses
  `DocxDocument` sessions. JSZip remains in the browser demo for preview.
- **Redline path:** targets are checked against a canonical source snapshot;
  supported batches execute atomically and are inserted once on confirmed
  success. A failed or uncertain host write is not replayed.
- **Reliability:** Gemini request transport/retry behavior and structured
  mutation outcomes are centralized. Retries are bounded and never repeat a
  document mutation. These, the regression/host test lanes, portable consumer
  extraction and deferred taskpane startup remain local consumer concerns; they
  are not new library offloads. See the [library offload review](docs/library-offload-review.md).
- **Performance:** the dated v0.8.2 mapping, core, package, browser-projection,
  Word-host, and startup measurements are recorded. The measured 1,000-
  paragraph workload is accepted; there is no universal latency gate.

## Completed Milestones

Five original plans are complete:

1. [September v0.8.2 upgrade](docs/plans/completed/2026-09-09-docx-redline-js-v0.5.4-upgrade.md)
2. [Reliability and quality gates](docs/plans/completed/2026-08-29-reliability-and-quality-gates.md)
3. [Package boundaries and integrations](docs/plans/completed/2026-08-29-package-boundaries-and-integrations.md)
4. [OOXML performance](docs/plans/completed/2026-08-29-oxml-engine-and-performance.md)
5. [Original agentic tools and list reliability plan](docs/plans/completed/2026-08-29-agentic-tools-and-list-reliability.md), closed with its remaining canonical migration transferred to the [new follow-up](docs/plans/2026-09-30-canonical-list-migration-follow-up.md).

Their closure reports record actual Office.js and independent Word checks,
browser open/edit/download/reopen validation, cross-host semantic parity,
current-version performance samples, and taskpane startup measurements. These
closures are scoped. The original list plan's remaining capability work and
full canonical migration were transferred to the new follow-up and are not
marked complete.

## Active Work and Upstream Dependencies

The [canonical list migration follow-up](docs/plans/2026-09-30-canonical-list-migration-follow-up.md)
remains open. The original list plan is closed with deferred scope transferred
here. Contract inventory and stale-source checks are complete. A
bullet/decimal insertion subset with tracking enabled at source/resolved levels
0–1 is verified, but full canonical list migration still requires
canonical-operation capability work and later consumer/Word validation.
Existing native or legacy paths remain for general `edit_list` and header
conversion, deeper/other-style insertion, tracking-off requests and unsupported
outdent contexts.
Current v0.8.3 validation is recorded in the
[release report](docs/validation-reports/2026-09-30-docx-redline-v083.md);
passing these gates alone will not close the full migration.

The current list matrix covers 12 supported cases with 48 Office.js checks and
five native routes with 20 additional checks (68 total). The independent Word
oracle reports 92 applicable checks passed, zero failed, and five not
applicable across 17 exports. The two original public-facade Reject All cases
pass 12 Word checks. `npm test` passes 51 suites with zero failures and four
exclusions; validation and production builds pass. See the [v0.8.3 report](docs/validation-reports/2026-09-30-docx-redline-v083.md)
and its linked JSON artifacts for the current evidence.

The dated 2026-09-30 v0.8.2 consumer baseline records five native cases, 20 actual
Office.js checks, 20 independent Word source/tracked/Accept All/Reject All
checks, and an offline aggregate of 50 suites with zero failures and four
excluded entrypoints. These are historical results, not new v0.8.3 validation
counts.

The v0.8.2 library reports documented Reject All paragraph-boundary failures
and historical paragraph-property inspection. The v0.8.3 release notes report
fixes for those behaviors, including all-empty list ranges, and for
`openDocx` list numbering and explicit list starts; current validation confirms
the original two public-facade Reject All cases. The dated reports retain
their original evidence. Current known limits are canonical conversion of
plain or text-changing headers to lists, canonical list-format changes, and possible
`RECEIPT_RECONCILIATION_FAILED` for a marker-prefixed `1. Header` no-op on bare
`document.xml`. The active plan remains open until canonical migration and its
downstream validation are complete.

## Operating Decisions

1. Prefer OOXML operations when the library supports the requested change and
   its accepted/rejected views and untouched package parts are verified.
2. Keep host reads/writes and Word-specific state in adapters; keep browser and
   MCP consumers on public package APIs.
3. Resolve a batch against one immutable source. Refuse ambiguous, stale, or
   unsupported operations before writing.
4. Preserve actual host outcomes. Do not claim rollback after an unconfirmed
   insertion and do not retry a mutation after an uncertain write.
5. Keep upstream library changes separate from this consumer repository.
   Update the exact version pin only after a release and compatibility checks.
6. Desktop Word evidence does not by itself certify Word Online behavior or
   every document structure; use the feature-specific validation records.

## Longer-Term Work

Broader list migration, additional Word host decoupling, Word Online fidelity,
and splitting the add-in, browser demo, and MCP server into separate
repositories remain future work. See [ROADMAP.md](ROADMAP.md).
