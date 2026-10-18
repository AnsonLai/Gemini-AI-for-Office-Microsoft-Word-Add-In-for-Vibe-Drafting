# Project state: Gemini Word add-in

**Updated:** 2026-10-12. The [implementation plan (closed 2026-10-12)](./docs/plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md)
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
The 0.8.4 surgical-insertion defect (new text tracked despite tracking being
disabled) is fixed in 0.9.0, and the original untracked-insertion gate passes
(verified 2026-10-12, [0.9.0 validation](docs/validation-reports/2026-10-12-docx-redline-v090.md)).
The prepared-output `TRACKING_POLICY_MISMATCH` check stays as a backstop (KB-5).

A new write requires a persisted exact-source body checkpoint and a final
freshness check. One insertion is followed by fresh host verification. Verified
no-ops create no write/checkpoint. Cancellation preserves settled evidence;
uncertain writes stop for inspection and never replay. The serial harness keeps
six model cycles, two failed mutation cycles, one proven no-write stale refresh
and complete call/response history. The owner persona cannot override execution
rules. There is one compact editing guide and no general skill framework.

The Word add-in and MCP pin exact `@ansonlai/docx-redline-js@0.9.0`. Browser and
MCP consumers retain their complete-DOCX facade; the Word plugin keeps its Flat
OPC transport. A complete-DOCX Word facade/binary experiment is excluded.

The milestone reports distinguish offline suites, actual Office.js exports,
independent desktop Word Accept All/Reject All checks and the original 16-prompt
NDA replay. The [0.9.0 validation](docs/validation-reports/2026-10-12-docx-redline-v090.md)
records 46 of 46 offline suites (three entrypoints excluded), passing builds, and
the 16-prompt direct replay at 16/16 steps and 49/49 checks with zero regressions
and zero auxiliary model calls. Actual Office.js passes all 14 default capability
cases, 20 of 20 host checks and three checkpoint restores, including untracked
insertion on a revised target; the independent Word oracle gives 68 PASS and
0 FAIL. WP5 is complete. The fixed production surface is 4,402 lines, 49.0%
below WP0 (as of the [WP5 correction report](docs/validation-reports/2026-10-02-lightweight-harness-wp5-correction.md)).
The original WP5 report remains historical evidence for the rejected blanket
guards. Replay does not certify model drafting quality. Live fast/slow provider checks remain required before shipping; Word
Online is not certified by desktop evidence.

No [library issues](docs/validation-reports/2026-10-17-library-issues-closeout.md) are open as of 0.10.0. Remaining
consumer-side bugs and the Word JavaScript API inventory are in
[KNOWN_BUGS.md](docs/plans/KNOWN_BUGS.md), notably KB-1 (a body-final table append
via live insertion leaves an extra paragraph on Reject All, so table append stays
refused) and KB-4 (released `list-format` capabilities are not yet exposed as
intents, so header conversion and numbering-style changes stay refused).
Retirement of unsupported consumer commands is a scope decision, rather than
proof that a capability is certified.

Historical incidents and prior milestones remain in dated plans/reports. The
[September list follow-up](docs/plans/completed/2026-09-30-canonical-list-migration-follow-up.md)
stays open (browser/MCP facade parity and retired-helper checks remain); its
library-issues cleanup was done on 2026-10-17. The
[table incident](docs/plans/completed/2026-09-30-table-creation-reliability.md) is
closed (2026-10-12) and retains its upstream evidence; the former native-command
migration queue is superseded by the approved smaller product scope.
