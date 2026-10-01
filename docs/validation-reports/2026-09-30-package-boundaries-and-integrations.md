# Package boundaries and integrations — execution record

## Archive checkpoint

`3daf159` moves the completed September upgrade and August reliability plans
into `docs/plans/completed/` and repairs their links. The agentic list plan
remains active with its separately recorded upstream dependencies.

## WP1 portable preparation

`consumer-core.js` is the explicit portable entry point. It contains source
inspection, baseline capture, mapping/validation, pure operation execution and
Flat OPC package preparation. `redline-plan.js` contains redline mapping and
stale-source checks. The Word modules retain host I/O and tracking; old named
exports are preserved through re-exports.

`prepareCanonicalBatch` returns ready, no-op, refused or package-error states.
It does not confirm a Word write. The Word adapter keeps the existing
confirmed/indeterminate write outcome contract and performs one source read
and one successful insertion for migrated batches.

Validation: **45 offline suites pass**, zero failures, four exclusions, with
the historical 0.5.4 skip and known library defects still reported explicitly.
`npm run build:dev` passes. The new core test imports/executes without Word/UI
globals and checks the portable source graph, rollback, mapping and preparation.
Existing paragraph/scope/batch adapter suites pass.

## WP2/WP3 verification underway

The browser session's local validation page passes seven checks in a real
Chromium browser: Word-authored input inspection; localized tracked edit and
serialize/reopen; unchanged style/numbering/relationship parts; exact
Accept All/Reject All; mixed-batch rollback; direct edits without new revisions;
direct mode preserving another author's revisions; and comment insertion with
existing thread preservation. Related assertions are grouped into seven
reported checks. [Browser evidence](2026-09-30-browser-document-session.json).

Browser UI migration, actual file/download behavior and cross-host parity are
still being verified. No provider calls or library-side fixes are part of this
checkpoint. Preview retains JSZip independently of editing. See
[consumer package boundaries](../package-boundaries.md).
