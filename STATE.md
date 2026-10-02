# Project state: Gemini Word add-in

**Updated:** 2026-10-02. The [current implementation plan](docs/plans/2026-10-02-library-first-lightweight-agent-reliability.md)
records milestone status, commits and continuation contracts. Dated evidence is
in [validation reports](docs/validation-reports/).

The Word plugin uses two tools: focused canonical inspection and one typed
atomic edit batch. Opaque handles bind to the document, accepted revision view
and captured parts. Every requested intent validates before preparation;
unsupported mixed requests refuse completely. The library performs document
transformations. Office.js reads/inserts body Flat OPC and preserves host state.
Native paragraph/list/table editing algorithms and tool-internal AI calls are
retired. The local Flat OPC assembly boundary remains explicit.

Retained actions are literal corrections, clause rewrites, plain paragraph
insertion/deletion, comments, basic inline emphasis, certified same-level
bullet/decimal insertion at levels 0–1, and table-cell text. Broader list/table
structure, font/layout, research and revision acceptance are refused. The
[tool contract](docs/agentic-tool-contracts.md) gives exact restrictions.
Adding comments requires a source without existing comments; supported wording
edits preserve existing comments and threads. The narrow refusal prevents an
unsupported merge-delta package from replacing original comment definitions.
Untracked changes on a paragraph with pending revisions also refuse before
writing; they must not silently flatten prior revision history.

A new write requires a persisted exact-source body checkpoint and a final
freshness check. One insertion is followed by fresh host verification. Verified
no-ops create no write/checkpoint. Cancellation preserves settled evidence;
uncertain writes stop for inspection and never replay. The serial harness keeps
six model cycles, two failed mutation cycles, one proven no-write stale refresh
and complete call/response history. The owner persona cannot override execution
rules. There is one compact editing guide and no general skill framework.

The Word add-in and MCP pin exact `@ansonlai/docx-redline-js@0.8.4`. Browser and
MCP consumers retain their complete-DOCX facade; the Word plugin keeps its Flat
OPC transport. A complete-DOCX Word facade/binary experiment is excluded.

The milestone reports distinguish offline suites, actual Office.js exports,
independent desktop Word Accept All/Reject All checks and the original 16-prompt
NDA replay. [Final WP5 validation](docs/validation-reports/2026-10-02-lightweight-harness-wp5.md)
passes 46 offline suites, 18 actual Office.js checks, 52 independent Word checks
and 49 replay checks across the original 16 prompts. Eleven replay batches use
ten writes and one no-op, with zero nested AI calls, zero regressions and correct
authors. The fixed production surface is 4,355 lines, 49.5% below WP0. Replay does not
certify model drafting quality. Live fast/slow provider checks remain required
before shipping; Word Online is not certified by desktop evidence.

Open [library issues](docs/library-issues/README.md) remain upstream dependencies:
unsupported canonical header conversion/list styles, separate header operations
restarting numbering, final table append leaving an extra paragraph under Reject
All, and weak-target soft-break deletion. Those shapes are refused or avoided
through strong supported targeting. Retirement of unsupported consumer commands
is a scope decision, rather than proof that upstream defects were fixed.

Historical incidents and prior milestones remain in dated plans/reports. The
[September list follow-up](docs/plans/2026-09-30-canonical-list-migration-follow-up.md)
and [table incident](docs/plans/2026-09-30-table-creation-reliability.md) retain
upstream evidence and cleanup conditions; their former native-command migration
queue is superseded by the approved smaller product scope.
