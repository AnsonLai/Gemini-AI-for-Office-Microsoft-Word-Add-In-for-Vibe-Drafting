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
WP5 was reopened after owner review rejected blanket restrictions on adding
comments to comment-bearing documents and untracked edits on revised paragraphs.
The repair restores comment delta merging through public helpers and delegates
revision-aware edits to the library's supported carrier-preserving policy.
Capability-preservation tests, rather than expected refusals, govern completion.
The expanded Word gate found an in-scope 0.8.4 surgical-insertion defect: new text
can become tracked despite tracking being disabled. A precise prepared-output
check blocks that write; ordinary revised-target replacement and formatting
remain supported. WP5 stays open until a published fix passes the original
insertion gate. This limitation is not an approved capability retirement.

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
NDA replay. The [WP5 correction report](docs/validation-reports/2026-10-02-lightweight-harness-wp5-correction.md)
records 46 passing offline suites, passing builds and 49 replay checks across
the original 16 prompts. Eleven replay batches use ten writes and one no-op,
with zero nested AI calls, zero regressions and correct authors. Actual Office.js
passes 13 supported edits with 19 checks and three checkpoint restores;
independent desktop Word passes 64 applicable checks with 16 not applicable.
A separate zero-write safety control does not satisfy the remaining insertion
capability in the default 14-case manifest. WP5 remains open. The fixed production surface
is 4,402 lines, 49.0% below WP0. The original WP5 report remains historical
evidence for the rejected blanket guards. Replay does not certify model drafting
quality. Live fast/slow provider checks remain required before shipping; Word
Online is not certified by desktop evidence.

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
