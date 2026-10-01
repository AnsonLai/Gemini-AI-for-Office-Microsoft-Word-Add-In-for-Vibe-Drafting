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

Extraction checkpoint `3018409`: **45 offline suites pass**, zero failures, four exclusions, with
the historical 0.5.4 skip and known library defects still reported explicitly.
`npm run build:dev` passes. The new core test imports/executes without Word/UI
globals and checks the portable source graph, rollback, mapping and preparation.
Existing paragraph/scope/batch adapter suites pass.

## WP2 browser document lifecycle — complete

The browser session's local validation page passes seven checks in a real
Chromium browser: Word-authored input inspection; localized tracked edit and
serialize/reopen; unchanged style/numbering/relationship parts; exact
Accept All/Reject All; mixed-batch rollback; direct edits without new revisions;
direct mode preserving another author's revisions; and comment insertion with
existing thread preservation. Related assertions are grouped into seven
reported checks. [Browser evidence](2026-09-30-browser-document-session.json).

The demo opens, inspects, atomically edits and serializes through
`document-session.js`, using the public library facade. Prompt context retains
source list labels/depth and run formatting. Direct mode preserves another
author's revisions. JSZip remains a preview dependency. The browser import map
uses the published bundle so bare XML dependencies no longer prevent startup.

Actual UI checks cover source upload/preview, reference-file ingestion,
deterministic kitchen-sink edits, DOCX downloads, reupload and XML download.
Kitchen-sink marker seeding uses supported `structuredContent: true`, creates
four separate paragraphs and verifies every source paragraph remains. Plain
marker names avoid Markdown underscore interpretation; existing underscore
markers remain accepted aliases. No library patch was required.

The downloaded package validates and retains all 24 source paragraphs in both
accepted and rejected views. It contains the edited text, list, table and
comment; both DOCX download controls produce identical bytes. Reupload displays
42 prompt paragraphs (44 inspected paragraphs, including empty revision-view
paragraphs), and the XML download equals its document part. Download event waits
timed out in browser automation; the files were verified on disk and reopened
through the actual UI. [UI evidence](2026-09-30-browser-demo-ui.json).

## WP3 cross-host parity and Word transport — complete

The same fixture/operations pass through portable preparation, the production
Word adapter with mocked transport, the browser session and MCP. Tests compare
exact accepted/rejected text, hyperlink relationships, untouched styles and
numbering, comment content, foreign revisions and atomic failure. Missing
targets retain the generic batch error and nested cause. The Word adapter
performs one confirmed insertion on success and zero writes on refusal.
Existing MCP stdio create/edit/batch/comment/rollback/save/reopen tests pass.

Actual desktop Word validates the refactored transport: **8 Office.js checks**
and **12 independent Word checks**, including Word Accept All/Reject All and
thread persistence. Evidence:
[Office.js](2026-09-30-package-officejs.json),
[Word oracle](2026-09-30-package-officejs-word.json).
The temporary validation registration and collectors were removed; the local
browser server was stopped after verification.

Final `npm test`: **47 suites passed, zero failed, four excluded entrypoints**.
The historical package skip and known list defects remain explicit. Development,
production and Word-validation builds pass. Production reports three existing
webpack size warnings (taskpane about 795 KiB).

No provider calls or library-side fixes were used. Broader list fidelity gates
remain in the active agentic plan; preview is not a Word fidelity oracle.
Whole-document Word binary insertion remains deferred. See
[consumer package boundaries](../package-boundaries.md).
