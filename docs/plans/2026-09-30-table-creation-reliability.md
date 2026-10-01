# Table creation and multi-step chat reliability

**Date:** 2026-09-30  
**Status:** In progress — consumer fixes and independent validation.

**Updated:** 2026-10-01. Context refresh, bounded history and loop termination
regressions pass; table mapping and separate upstream formatting evidence are
still being verified.

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

- Installed 0.8.3 supports the retained final paragraph followed by that table;
  the standalone accepted view contains one table with nine cells. The rejected
  view restores the seven source paragraphs.
- When one proposal edits P7 and another appends at P8, the consumer maps both
  to P7. The library refuses overlapping source targets atomically with
  `OVERLAPPING_SOURCE_TARGETS`. Compatible requests need one consumer operation.
- The chat captures prompt text and its source baseline once per request and
  reuses them after its own writes. Subsequent structural edits can therefore
  use obsolete paragraph indexes/counts. A formatting-only underline need not
  change the library fingerprint, so it does not by itself prove stale refusal.
- Read-only tool iterations reset the failed-mutation counter, allowing repeated
  edit attempts to avoid the two-failure stop.
- The history window can drop the real user request, then discard the remaining
  call/response pairs because the selected window starts with a model turn.
- Default batch logging hides error codes behind a generic warning.

The pasted log alone does not identify which engine code was returned. These
are reproduced consumer defects, not proof of every cause in the user's document.
No library patch is required by the supported table reproduction.

## Work packages

- [ ] WP1: coalesce compatible final-paragraph editing and append requests;
  preserve literal targeting, reject ambiguous overlaps and verify table/text
  and tracked revision fidelity.
- [x] WP2: refresh prompt context and canonical baseline after confirmed writes
  before another model iteration; stop if rereading fails; keep external-edit
  guards and no replay after uncertain writes.
- [x] WP3: retain the real request and complete recent tool exchanges within
  the bounded history window; read-only tools do not reset mutation failures.
- [x] WP4: log failure codes, operation indexes and write outcomes without
  logging source text, generated content or raw errors.
- [ ] WP5: run targeted regressions, the offline aggregate and builds; verify
  the actual table package through Word and record the scope of host evidence.
- [ ] WP6: update current contracts/state/summary and commit the fixes.

## Acceptance

The matching supported exercise produces one native table with exactly nine
cells while preserving the final source paragraph and requested underline.
Reject All restores source text and paragraph boundaries. Multi-step execution
uses refreshed context, long tool turns retain request/response history, and
two unsuccessful mutation rounds stop even with read-only rounds between them.
Required source/refusal/privacy tests and builds pass. No provider call is
needed for the deterministic reproductions.
