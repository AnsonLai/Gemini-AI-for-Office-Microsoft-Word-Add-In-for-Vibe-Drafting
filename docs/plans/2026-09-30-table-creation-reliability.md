# Table creation and multi-step chat reliability

**Date:** 2026-09-30  
**Status:** Open — consumer fixes are complete; tracked Reject All and underline fidelity remain blocked by upstream behavior.

**Updated:** 2026-10-01. Consumer regressions and the production build pass.
Compatible edit-plus-append mapping, append-anchor preservation, receipt mapping,
formatting refusal, refreshed context, bounded history, failed-mutation budget,
and visible loop termination are implemented. Offline coverage passes 54 suites.
The final Word run passed 13 of 20 checks: plain accepted table content and
both native insertions pass; tracked Reject All paragraph structure and
underline fidelity fail. See the
[validation report](../validation-reports/2026-10-01-table-creation-reliability.md),
[tracked-formatting report](../library-issues/2026-10-01-table-append-tracked-formatting.md),
and [Reject All paragraph report](../library-issues/2026-10-01-word-reject-table-append-paragraph.md).

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
  and [tracked-formatting report](../library-issues/2026-10-01-table-append-tracked-formatting.md)
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
  the host write with `UNSUPPORTED_TABLE_FORMATTING`.

The pasted log alone does not identify which engine code was returned, and this
incident did not include a full live chat/model reproduction. The consumer
regressions explain reproduced failure modes but do not prove every cause in
the user's document. The separate 0.8.3 library issue concerns tracked inline
formatting preservation; it is not repaired by the consumer changes.

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

## Acceptance

The deterministic consumer/package tests pass for the nine-cell table,
preserved source targeting, refreshed context, complete tool history, and a
two-failure mutation budget that read-only calls cannot reset. Offline tests
and the production build pass. Plain accepted table content passes in Word,
but tracked Reject All paragraph structure and tracked underline fidelity do
not. No full live chat/model reproduction or provider call is claimed here.
