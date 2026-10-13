# WP4 — retire native editing routes

**Date:** 2026-10-02. **Library:** exact 0.8.4, unchanged in add-in/MCP.
**Status:** complete locally; live model and Word Online shipping limits remain.
See the [approved plan](../plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md).

The supported Word path now has no native paragraph/list/table edit executor or
fallback. Removed `agentic-tools.js`, mutation-outcome/converter/table validators,
orphan markdown/font and selection-statistics helpers, and obsolete executor
suites. Guards now assert two typed tools, no nested generation and no legacy
executor dispatch. Unsupported actions refuse before preparation; plain
paragraph deletion also refuses list/table targets. Every batch is capped at
40 intents.

Retained list insertion is deliberately narrow: one plain item at the same level
as an inspected bullet/decimal level 0–1 item outside tables/headings. It uses the
library's documented `list-change` with the complete marked current/new items.
Fixed Markdown markers describe structure; source numbering determines visible
labels. Public-library tests cover both bullet roots, nested items, decimal
continuation/restart, source/new/sibling list identity and rejected boundaries.

The initial ordinary `redline` recipe preserved the new item's numbering but
removed the existing anchor's list properties. Independent Word checks caught
this; it was not accepted as a passing capability. The documented public
`list-change` resolved it without modifying the engine or reconstructing OOXML.
The final retained scope below passed fresh Word checks. Deeper levels, other
styles, conversion, outdent and structural table changes remain refused.

## Verification

- Final offline aggregate: 56 suites pass, zero fail, four entrypoints excluded.
  This includes one concurrently prepared WP5 checkpoint-restoration suite;
  the WP4 milestone itself contains 55 suites. The historical v0.5.4 behavior
  skip remains explicit on installed 0.8.4.
- Development golden/collector bundles and production build pass. The deferred
  library asset remains 530 KiB with its size warning. Nonfatal Office tooling
  telemetry access errors did not affect compilation.
- Actual production Office.js collector: 11/11 cases pass. Each accounts for
  every edit, persists one exact-source checkpoint, commits once, verifies
  fresh host state, and makes zero native editing calls. Four tracked list
  contexts cover root bullet, nested decimal, continuation and separate restart;
  a fifth list case disables redlining with prior `TrackAll` restored and no
  new revisions. Core cases cover literal correction, rewrite plus comment,
  plain insertion/deletion, partial underline and table-cell text.
- Independent desktop Word: 44/44 applicable source/tracked/Accept All/Reject
  All checks pass; 11 engine-reference-package checks are not applicable.
  They are not counted as passes. Word build 16.0.20430; Office.js diagnostics
  report 16.0.20430.20092. Exact text, original/inserted/sibling list identity,
  continuation/restart labels, formatting sentinels, comments and table cells
  are checked independently of prepared output.
- Full real-taskpane direct NDA replay: all 16 original prompts complete,
  49/49 checks pass, zero regressions/desyncs, configured author throughout.
  Eleven fully accounted verified batches produce ten writes and one no-op.
  Batch counters: 21 source reads, ten content writes, ten verification reads.
  Thirty-eight mocked chat calls; zero auxiliary/other requests. Ten prompts
  exercise retained changes and six unsupported requests explicitly refuse
  without changing document semantics. Step 7 now verifies same-level insertion.

The collector uses genuine Word proxies and isolated `/x` Word on port 3002;
its temporary registration and owned PID are cleaned on result. Golden runs
on separate port 3001. Fixture membership is read in the same sync as OOXML;
only a certified export-only final paragraph is excluded from handles/oracle
projection. NDA host membership is 64 versus 65 canonical paragraphs, yielding
56 outside-table paragraphs in the Word oracle. No source or payload is changed
by this projection.

Ignored reproducible artifacts: `.cache/reliability/direct-word-document-tools/`
(`officejs-report.json`, `officejs-fixtures.json`, `oracle/word-report.json`);
`.cache/reliability/wp4/final-dist` and `completed-golden/report.json`.
Checked-in fixture generation: `tests/direct_document_fidelity_manifest_tests.mjs`.
Commands and retained fixtures do not depend on ignored artifacts.

## Net size and continuation

The frozen WP0 physical-line surface is 8,632 lines / 24 files. WP3 was 9,083 /
33. WP4 is **7,004 / 29**, 2,079 fewer than WP3 and **1,628 fewer than WP0**.
The count includes taskpane, all commands/integration/chat JS and the commands
entrypoint, with blank/comment lines and any final unterminated line. The
unwired 41-line WP5 checkpoint helper is excluded from this milestone. Another
613 orphan utility lines were deleted outside this fixed measurement surface;
they are not added to its net reduction.

WP5 removes unused compatibility/planning/context exports, preserves the minimal
Flat OPC body bridge, verifies actual checkpoint restoration and updates all
current docs. Browser/MCP facades and the exact package pins remain intact.
Offline/replay/desktop results do not certify live drafting quality or Word
Online. Live fast/slow provider checks remain mandatory before shipping.
