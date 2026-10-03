# WP5: final bridge thinning and checkpoint reliability

**Date:** 2026-10-02 (America/Vancouver). Evidence for commit `2aa6aca`; WP5 reopened.
**Scope:** Final package of the [approved reliability plan](../plans/2026-10-02-library-first-lightweight-agent-reliability.md).
**Library:** Both consumers remain exact-pinned to `@ansonlai/docx-redline-js@0.8.4`.

**Owner review correction:** The two blanket restrictions below are unacceptable
day-to-day capability regressions. Passing refusal checks does not establish
feature preservation. This report records the committed behavior and its tests;
it does not certify WP5 completion. The [plan](../plans/2026-10-02-library-first-lightweight-agent-reliability.md)
now tracks repairs for ordinary comments and revision-aware editing. The separate
[correction report](2026-10-02-lightweight-harness-wp5-correction.md) supersedes
these blanket guards and retains the original failed artifacts and conclusions
here as historical evidence.

The Word plugin retains two document tools, a bounded local harness and Flat OPC
transport. The library owns transformations. All orphan native command modules,
compatibility planner/result/barrel exports and paragraph/range runner paths are
deleted. `consumer-core.js` exposes only canonical inspection and preparation;
the Word runner performs one body insertion with tracking preservation. This is
local deletion and simplification, rather than a claimed new upstream API.

Unused baseline formatting, context/range bookkeeping, generic loop checkpoint
branches, model diff settings and taskpane dependency scaffolding are removed.
The compiler, strong handles, whole-batch validation, bounded history and truthful
host evidence remain. Current documentation describes the retained actions and
refusals. Browser/MCP behavior and library pins are unchanged.

## Checkpoint capture and restore

`saveCheckpoint` now resolves only on the IndexedDB transaction's completion.
Request success followed by transaction abort/error cannot authorize a new
document write. Regression tests exercise that ordering and failure behavior.

The wired restore helper validates the saved package, disables native tracking,
replaces the body once, restores the prior tracking mode and reads the resulting
parts. It never clears the body first or replays an uncertain write. Cancellation
after a settled insertion still permits verification; host errors retain observed
write evidence. A failed restore preserves history. A verified restore clears
model history and visible messages together, including stale retry/revert controls.

Actual Word initially rejected the strict comparison because Replace regenerated
body paragraph `w14:paraId` and `w14:textId`, alongside known volatile session
metadata. The saved package and restored package covered the same nine parts;
diagnostics isolated the remaining differences to these IDs. A separate restore
initial comparison omitted those body IDs and prior confirmed session metadata.
Normal freshness tokens still include real paragraph/text IDs, and original
checkpoint XML is untouched. Revision/comment IDs, authors, text, formatting,
numbering, relationships and all other transported part content still count.
Tests verify both preservation failures and strict ordinary source identity.

A subsequent three-comment checkpoint exposed consistent regeneration of
comment paragraph IDs and their `commentsExtended`/`commentsIds` references.
Public canonical inspection matched every remaining comment field, including
reply parent ID, resolved state, author/date/text and paragraph anchors; durable
IDs and body comment-range/reference IDs were retained. Restore now omits those
proven metadata IDs from its derived token only, requires resolvable metadata
references, and separately requires identical canonical comment identity. It
validates the saved identity before writing. It never reconstructs comment
metadata or changes the original checkpoint. Unknown mappings refuse.

## Scope boundary discovered by Word probes

Two isolated new-comment probes on an existing threaded-comment source failed
Word insertion with `GeneralException`. Each recorded one attempt, an uncertain
outcome and no replay. Investigation compared prior and thinned adapters: both
assembled four body anchors with only the new comment definition. The library
correctly labels new-comment output as a merge delta; the adapter had treated
it as a complete replacement. The public injection helper appends a duplicate
comments part and is unsuitable for merging. No supported complete merge helper
for the retained Word Flat OPC seam was found. The public ZIP helper merges
ordinary comments on a different transport boundary and does not alone assemble
all thread siblings. No private helper or consumer XML merge was copied.

The retained Word path now refuses adding comments whenever existing comments
are present, including whole mixed batches, before checkpoint/write. Text edits
preserving existing threads remain supported. The versioned guide and source
summary explain this restriction. The [dated transport report and public repro](../library-issues/2026-10-02-word-existing-comments-delta-transport.md)
preserve the evidence and API dependency. This is a consumer transport limitation,
rather than a claimed library mutation defect or an upstream fix.

A source-preparation probe also showed that an untracked rewrite of the same
paragraph as an existing own revision flattened that revision, making Reject
All retain the new wording. This was preserved as a failed fidelity result.
The Word path now refuses any untracked non-comment change targeting a paragraph
with pending revisions, with `UNTRACKED_EDIT_ON_REVISED_TARGET`, before
checkpoint/write. The whole mixed batch refuses. A separate pending-revision
paragraph is used to verify preservation during an otherwise supported edit and
restoration of the actual checkpoint with its three comments and pending revision.

## Verification

The final collector and independent oracle used desktop Word 16.0.20430.

- Actual Office.js: **11/11 main cases and 18 checks pass**, including two
  prewrite refusal controls and two saved IndexedDB checkpoint restorations.
  Both restricted batches leave checkpoint and write counts unchanged.
- Independent Word: **52/52 applicable checks pass, zero failures**, across
  11 main cases and two restored checkpoints. Thirteen engine-reference checks
  are explicitly not applicable. Tracked, accepted and rejected views preserve
  expected text, comments, list identity, structure and author attribution.
  The threaded checkpoint restores all three comments and exactly one native
  Word revision. Its two OOXML revision elements are grouped by Word as one
  COM revision before and after restoration; these counts are recorded separately.

- Offline aggregate: **46 suites pass, zero failures; three entrypoints
  excluded** (XML-provider setup, manual Word list fixture lane, observational
  performance lane). The 0.5.4 compatibility suite checks the production atomic
  guard but explicitly skips package-version behavior under installed 0.8.4.
  Legacy planner/executor-only suites and old paid eval fixtures were retired;
  supported compiler/session/batch/checkpoint/list/public-package coverage remains.
- Final direct NDA taskpane replay: **16/16 original prompts, 49/49 checks**, no
  regressions, replay/model-evidence failures or author failures. Eleven verified
  batches contain ten writes and one no-op; 38 mocked chat calls, zero auxiliary
  or other calls. This validates harness/host behavior, rather than model quality.
- Development/golden build and final production build pass. Production has one
  existing deferred library asset-size warning (529 KiB); blocked Office build
  telemetry is nonfatal. Manifest validation passes against Microsoft's service.
- Direct benchmark smoke passes with one warmup and one measured iteration per
  metric. This establishes runnable current paths, rather than a latency claim.
- The completed table/formatting public-library repro passes its 0.8.4 regression
  guard. The open Word Reject All incident has a new public-only fixture exporter;
  export smoke passes. Structural tables remain refused by the agent, and no
  upstream defect is closed by deletion of its consumer command.

The live runner defaults to the direct protocol, initializes actual configured
fast/slow profiles and records tool calls/receipts from real provider exchanges.
It flags missing, truncated or invalid current-turn evidence rather than
fabricating success from a planned response. Profile-specific commands are in
[scripts/README.md](../../scripts/README.md). Keys/headers/URLs are excluded from
model logs. No live provider calls were performed; that shipping gate remains.
Word Online is also unverified. Checkpoints cover the transported body package,
rather than a complete-document backup.

## Net production measurement

The checked-in `scripts/count-document-agent-code.mjs` counts physical lines,
including comments/blanks and final unterminated lines, over the same taskpane,
command, integration and chat roots at every milestone. Missing deleted roots
count as zero. All new replacement production modules are included.

| Fixed production surface | WP0 | WP4 | Final WP5 |
| --- | ---: | ---: | ---: |
| Taskpane | 2,658 | 1,328 | 1,242 |
| Commands | 3,452 | 1,346 | 0 |
| Integration | 1,683 | 1,871 | 739 |
| Chat | 804 | 2,424 | 2,339 |
| Ribbon command entrypoint | 35 | 35 | 35 |
| **Total lines** | **8,632** | **7,004** | **4,355** |
| **Files** | **24** | **29** | **20** |

Final reduction is **4,277 lines (49.5%)** from WP0 and **2,649 lines** from WP4.
WP4 additionally deleted 613 utility lines outside these fixed roots, which are
not added to this comparison. Model-profile/checkpoint-store changes outside
the roots are also excluded. Counts include replacement modules rather than
reporting isolated taskpane shrinkage.

## Reproduction and artifacts

Use `npm test`, `npm run build`, `npm run validate`, the Office.js/Word commands
and the direct golden/live commands in [scripts/README.md](../../scripts/README.md).
Count a milestone with `node scripts/count-document-agent-code.mjs --ref REF`.
Tests and fixture generation are checked in; machine artifacts are ignored.

Local evidence is under `.cache/reliability/wp5/`:

- `completed-golden-final/report.json`: final rebuilt direct replay, clean runner exit.
- `final-main-exact-restore/officejs-report.json`: actual host and checkpoint controls.
- `final-main-exact-restore/word-oracle/word-report.json`: independent Word assertions.
- `final-dist`, `benchmark-smoke.json`, `public-table-append-repro`: bundle and smoke artifacts.

Earlier failed restore/comment/revision probes remain in
`direct-officejs-diagnostic`, `direct-officejs-final`,
`exploratory-comment-add` and intermediate collector outputs. They precede the
documented identity comparison and scope restrictions; they are not counted
as passing validation. No paid provider run or Word Online run was performed.
