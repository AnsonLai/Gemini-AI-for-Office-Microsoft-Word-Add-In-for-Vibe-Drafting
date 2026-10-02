# WP2 — Canonical source and inspection handles

Date: 2026-10-02. Exact library: 0.8.4. Parent milestone: `2c179a1`.

The taskpane now builds its accepted text, list/table cues, comments, revision-author cues, and strong baseline from the same library inspection of body Flat OPC. Capture uses one body OOXML read, selection text, and one Office.js sync. Removed the duplicate Word metadata walk and silent canonical-to-Word text fallback. Missing canonical context stops editing.

The portable inspection session binds opaque paragraph/end handles and pagination cursors to document identity, revision view, full canonical part token and complete target descriptors. One fresh capture resolves every handle in a batch. External edits, Undo/source changes, styles, numbering, comment/thread/sibling changes and explicit invalidation expire handles. Truncation is visible and full-text coverage is retained privately. Output text and stored handles/cursors are bounded. Freshness is a check, not a host document lock. The direct model protocol is WP3; the old declarations remain in this milestone.

Validation:

- 63 milestone offline suites pass, 0 fail; four entrypoints excluded and existing suite-internal skips retained. The concurrent, uncommitted WP3 compiler suite also passed in the aggregate run (64 total), and is outside this milestone. Reader tests use actual library inspection and minimal host mocks; session tests cover source/view/document/structural freshness, thread cues, pagination, truncation, and no model-facing physical indexes.
- Development and production builds pass. Production retains the existing deferred 530 KiB asset-size warning. Golden build passes; blocked Application Insights telemetry emitted a non-fatal warning.
- `node scripts/run-golden-scenario.mjs --launch --skip-build --dist-dir .cache/reliability/wp2/dist --artifacts-dir .cache/reliability/wp2/golden`: 16/16 prompts, 29/29 content checks, same three author failures as WP0/WP1 (steps 6/9/12; Reviewer revisions 5/3/27), 0 regressions/desynchronizations. Exit 1 accurately records those failures. 43 mocked requests: 32 chat + 11 auxiliary; no live model calls. Disposable Word PID 59692 closed after completion; Word/PC executable version 16.0.20430.20092.
- No live Gemini, Word Online, new Accept All/Reject All certification, or Office.js runtime read/write counters in the golden runner. The reader’s one-sync contract is an offline observation; do not present it as a golden runtime counter.

Artifacts: `.cache/reliability/wp2/dist`, `.cache/reliability/wp2/golden`; parent build logs `.cache/reliability/wp2-build-dev.log` and `wp2-build-prod.log`. Ignored artifacts are not required by checked-in tests.

Same physical-line counting roots/method as WP0: taskpane 1,389; commands 3,469; integration 1,700; chat 2,429; command entrypoint 35. Total **9,022 / 29 files**, versus WP1 8,633 / 27: **+389 lines / +2 files**. Concurrent WP3 compiler/guide are excluded because they are not wired or part of this milestone. Net deletion remains required in WP3–WP5.
