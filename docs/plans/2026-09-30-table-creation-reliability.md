# Table creation and multi-step chat reliability

**Date:** 2026-09-30  
**Status:** Open — consumer recovery is implemented; tracked Word fidelity remains blocked by separate upstream issues.

**Updated:** 2026-10-01. Consumer regressions and the production build pass.
Compatible edit-plus-append mapping, append-anchor preservation, receipt mapping,
formatting refusal, refreshed context, bounded history, failed-mutation budget,
and visible loop termination are implemented. Offline coverage passes 54 suites.
The final Word run passed 13 of 20 checks: plain accepted table content and
both native insertions pass; tracked Reject All paragraph structure and
underline fidelity fail. See the
[validation report](../validation-reports/2026-10-01-table-creation-reliability.md),
[tracked-formatting report](../library-issues/completed/2026-10-01-table-append-tracked-formatting.md),
and [Reject All paragraph report](../library-issues/2026-10-01-word-reject-table-append-paragraph.md).

**2026-10-01 follow-up:** A later pasted log shows two successive P4 proposals
without `anchorText`, each refused as `STALE_DOCUMENT_CONTEXT` with
`written: false` and `writeAttempted: false`. This proves two no-write refusals,
not their specific cause. Missing `anchorText` skips the separate anchor check;
the stale guard can also refuse an unavailable baseline or a text/fingerprint
mismatch. The fingerprint includes `w14:paraId`, so an ID-only change is a
plausible false-stale cause, but no host evidence confirms Word changed the ID.
The fingerprint guard remains unchanged. A bounded COM probe found stable
captures across repeated reads, tracking toggles and Reject All. Captured
paragraph IDs were absent; restart and Ctrl+Z were not exercised. This does
not establish Office.js `body.getOoxml()` behavior or the cause of the user's log.
See the [baseline probe report](../validation-reports/2026-10-01-word-com-baseline-stability.md).

The consumer recovery is implemented in `54dd082`: after a proven no-write
stale refusal, take one fresh document snapshot and baseline per request, then
ask the model to replan from that context. The refused change set is not
replayed, the failed-mutation budget is not reset, and uncertain writes do not
trigger recovery. Safe diagnostics distinguish baseline, text and fingerprint
mismatches without exposing document text or paragraph IDs. The post-Undo
refusal/refresh/replan regressions, all 54 offline suites and production build pass.

## Incident

The user requested underlining followed by a table at the document bottom with
nine nouns in three rows and three columns. The pasted log contains two proposed
changes, attempts to append at P8 after a seven-paragraph context, repeated
generic batch warnings, no-progress counters repeatedly returning to 1/2, and
history pruning that loses tool calls and responses.

The reported final instruction retained `Final paragraph stays unchanged.` and
appended a Markdown table containing Mountain/River/Forest,
Ocean/Valley/Canyon and Meadow/Desert/Island.

## Reproduced findings

- The deterministic package reproduction creates a table with exactly nine
  cells. In desktop Word, the plain accepted view preserves the three-by-three
  table, all nine nouns, paragraph text, and the required trailing paragraph.
  Both native COM `InsertXML` insertions pass. Tracked Reject All leaves an eighth
  empty paragraph after the seven source paragraphs. The formatting fixture's
  Accept All loses underline, and its Reject All also leaves an extra empty
  paragraph. See the [host report](../validation-reports/2026-10-01-table-creation-reliability.md)
  and [tracked-formatting report](../library-issues/completed/2026-10-01-table-append-tracked-formatting.md)
  and [Reject All paragraph report](../library-issues/2026-10-01-word-reject-table-append-paragraph.md).
- When one proposal edits P7 and another appends at P8, the consumer maps both
  to P7. The library refuses overlapping source targets atomically with
  `OVERLAPPING_SOURCE_TARGETS`. Compatible requests need one consumer operation.
- The chat now refreshes prompt text and its source baseline after a confirmed
  mutation exchange, before the next model turn. It does not refresh after
  refused or uncertain writes, replay writes, or reinterpret queued operations
  against a different snapshot. External-edit stale-source guards remain.
- Before the consumer fix, read-only tool iterations reset the failed-mutation
  counter, allowing repeated edit attempts to avoid the two-failure stop. The
  regression now verifies that read-only calls do not reset it.
- Before the history fix, the bounded history window could drop the real user
  request, then discard remaining call/response pairs when the selected window
  began with a model turn. The regression now retains the request and complete
  recent exchanges.
- Default batch logging now reports structured failure codes, mutation outcome,
  and write evidence. Compatible final-paragraph editing plus one table append
  is narrowly coalesced and its two original change indexes map to the shared
  operation receipt. Unverified tracked inline formatting is refused before
  the host write with `UNSUPPORTED_TABLE_FORMATTING`. As of 0.8.4 (2026-10-02)
  that refusal covers only changing the text of a paragraph that carries the
  author's own pending formatting while adding a table; new inline formatting
  with a table and appends after an underlined paragraph are allowed.

The pasted log alone does not identify which engine code was returned, and this
incident did not include a full live chat/model reproduction. The consumer
regressions explain reproduced failure modes but do not prove every cause in
the user's document. The separate 0.8.3 library issue concerned tracked inline
formatting preservation; it was not repaired by the consumer changes and is
fixed in 0.8.4 (see the [0.8.4 validation](../validation-reports/2026-10-02-docx-redline-v084.md)).
Word Reject All still leaves one extra empty paragraph after a final table
append (open library report).

## Work packages

- [x] WP1: coalesce only the compatible final-paragraph edit plus one table
  append; preserve the append anchor, reject incompatible overlaps and map both
  source changes to the operation receipt. Word fidelity remains a release gate.
- [x] WP2: refresh prompt context and canonical baseline after confirmed writes
  before another model iteration; stop if rereading fails; keep external-edit
  guards and no replay after uncertain writes.
- [x] WP3: retain the real request and complete recent tool exchanges within
  the bounded history window; read-only tools do not reset mutation failures.
- [x] WP4: log failure codes, operation indexes and write outcomes without
  logging source text, generated content or raw errors.
- [ ] WP5: offline regressions (54 suites) and production build pass. The final
  Word run executed 20 checks: 13 passed and 7 failed. Plain accepted table
  content and both native insertions pass; tracked Reject All paragraph
  structure and underline fidelity remain blocked. Full host scope is recorded
  in the validation report.
- [x] WP6: current contracts/state/summary and library-issue notes are updated.
  Source/history and consumer/harness changes are committed as `7012fed` and
  `d3e315f`; these documentation updates accompany the validation checkpoint.
- [x] WP7: implement and verify one bounded fresh-context replan after a
  proven no-write `STALE_DOCUMENT_CONTEXT`; preserve the failed-mutation count,
  never replay the old batch, and report safe refusal reasons. Record the COM
  probe's covered paths and its restart/Ctrl+Z/Office.js limitations. No exact
  reproduction of the user's source document is claimed.

## Acceptance

The deterministic consumer/package tests pass for the nine-cell table,
preserved source targeting, refreshed context, complete tool history, and a
two-failure mutation budget that read-only calls cannot reset. Offline tests
and the production build pass. Plain accepted table content passes in Word,
but tracked Reject All paragraph structure and tracked underline fidelity do
not. Stale-refusal recovery passes deterministic regressions; the specific
cause of the user's P4 refusal remains unconfirmed.
No full live chat/model reproduction or provider call is claimed here.

Undo probe (2026-10-01): a bounded Word COM run (tracked `Range.Text` edit and untracked `InsertXML` edit, each reverted with `Document.Undo()`) restored the original text and fingerprints on the five-paragraph fixture; COM output carried no `w14:paraId`, so paragraph-ID behavior after Undo is still untested and restart remains unrun. See `docs/validation-reports/2026-10-01-word-com-undo-probe.md`.
