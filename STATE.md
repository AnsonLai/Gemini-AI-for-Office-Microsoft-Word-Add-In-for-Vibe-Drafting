# Project State: Gemini Word Add-in

> Current technical posture: 2026-09-30. Detailed implementation and test
> evidence lives in the [dated plans](docs/plans/) and
> [validation reports](docs/validation-reports/).

## Current Posture

- **Library:** the Word add-in and MCP server pin exact
  `@ansonlai/docx-redline-js@0.8.2`. Browser editing uses the same package
  through its public document-session facade.
- **Host boundaries:** portable mapping and OOXML preparation are separated
  from Word transport. The Word adapter owns `getOoxml`, `insertOoxml`,
  tracking-mode changes, and confirmed/indeterminate write outcomes. Browser
  editing uses a document session; MCP uses package document sessions. JSZip
  remains in the browser demo for preview.
- **Redline path:** targets are checked against a canonical source snapshot;
  supported batches execute atomically and are inserted once on confirmed
  success. A failed or uncertain host write is not replayed.
- **Reliability:** Gemini request transport/retry behavior and structured
  mutation outcomes are centralized. Retries are bounded and never repeat a
  document mutation.
- **Performance:** current v0.8.2 mapping, core, package, browser-projection,
  Word-host, and startup measurements are recorded. The measured 1,000-
  paragraph workload is accepted; there is no universal latency gate.

## Completed Milestones

Four dated plans are complete:

1. [September v0.8.2 upgrade](docs/plans/completed/2026-09-09-docx-redline-js-v0.5.4-upgrade.md)
2. [Reliability and quality gates](docs/plans/completed/2026-08-29-reliability-and-quality-gates.md)
3. [Package boundaries and integrations](docs/plans/completed/2026-08-29-package-boundaries-and-integrations.md)
4. [OOXML performance](docs/plans/completed/2026-08-29-oxml-engine-and-performance.md)

Their closure reports record actual Office.js and independent Word checks,
browser open/edit/download/reopen validation, cross-host semantic parity,
current-version performance samples, and taskpane startup measurements. These
closures are scoped; they do not certify the unimplemented list operations
below.

## Active Work and Upstream Dependencies

The [agentic tools and list reliability plan](docs/plans/2026-08-29-agentic-tools-and-list-reliability.md)
remains open. Contract inventory and stale-source checks are complete. A
supported bullet/decimal insertion subset is verified, but full canonical list
migration is pending upstream library capabilities/fidelity fixes and later
consumer/Word validation. Existing native paths remain for unsupported
formats, levels, tracking modes, general list edits, and header conversion.
The current live-validation state is recorded in the
[agentic validation report](docs/validation-reports/2026-09-30-agentic-tools-and-list-reliability.md);
passing those live gates alone will not close the full migration.

The latest native matrix passes five cases: 20 actual Office.js checks and 20
independent Word source/tracked/Accept All/Reject All checks. The offline
aggregate passes 50 suites, zero failures, with four excluded entrypoints.

Four separately reported library follow-ups remain:

- [Reject All after plain insertion leaves an empty paragraph](docs/library-issues/2026-09-30-list-insertion-rejection.md) — independently confirmed in Word.
- [Reject All after a list-range edit merges source paragraphs](docs/library-issues/2026-09-30-list-range-rejection.md) — independently confirmed in Word.
- [Unmarked/text-changing header conversion has no verified canonical mapping](docs/library-issues/2026-09-30-canonical-list-operations.md).
- [Historical list properties are treated as active numbering](docs/library-issues/2026-09-30-historical-list-inspection.md) — reproduced offline; the current operation refuses without a write.

These are upstream-owned defects or capability gaps. No unreleased fix or
consumer workaround is assumed. Keep the package pin at 0.8.2 until a released
version passes the relevant contract, fidelity, and host checks.

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
