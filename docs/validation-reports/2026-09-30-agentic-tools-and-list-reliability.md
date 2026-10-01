# Agentic tools and list reliability — execution record

## Consumer validation checkpoint

- Exact library version remains `@ansonlai/docx-redline-js@0.8.2`.
- All three list tools validate request types, indexes and supported enums before Word execution, then recheck indexes against the live paragraph count. Invalid targets produce `INVALID_LIST_REQUEST` with no content write.
- `edit_list` no longer silently clamps an invalid range to another paragraph.
- `convert_headers_to_list` binds replacement text to the original index before sorting; duplicate targets refuse.
- Deterministic production invocation tests cover malformed input, live range refusal, unsorted header pairing and duplicate refusal. The pure request validator and existing mutation-outcome suites pass.
- Table content is declared as nested arrays and normalized without silently discarding excess rows or cells. Live per-row dimensions are checked before the first content write. Production tests cover oversized overlays, strict coordinates, valid nested content, the declared schema, missing-key failures and navigation outcomes.
- Canonical source-baseline tests confirm accepted-view text, manual line breaks and stable fingerprints. The stale-context guard checks every targeted source paragraph before batch preparation. Passing regressions cover matching source, whole-paragraph edits, range interiors, insert-before anchors, append after a changed paragraph count, unavailable baselines and mixed valid/stale batches with zero writes.
- Ordinary development and development-only Word validation entry builds pass. Optional telemetry requests are blocked by the environment and do not fail compilation.

## Canonical list migration and live testing

The actual Office.js context lane passed **8 checks** on Word `16.0.20326.20158` (PC), including `STALE_DOCUMENT_CONTEXT` refusal with `written: false`, `writeAttempted: false` and zero insertions after an intervening host edit. [Collector evidence](2026-09-30-agentic-officejs-context.json).

Independent Word source/export/engine-reference reopening passed **12 checks**, including exact accepted/rejected text and comment-thread persistence. [Word oracle evidence](2026-09-30-agentic-officejs-context-word-oracle.json). The first oracle invocation exposed rooted-path handling in external manifests; after correcting the harness, all checks passed.

The verified `insert_list_item` subset is wired into production: redlining enabled, active bullet/decimal source and resolved levels 0–1, with a same-list root sibling when needed for outdent. It uses one body source read and one confirmed insertion. Other shapes retain native behavior; a failed engine/host batch never replays through native insertion. General list replacement and header conversion remain on established paths.

The final offline aggregate passes **44 suites**, zero failures, with four entrypoints excluded and the historical 0.5.4 package-behavior skip reported separately. Known library rejection defects remain explicitly reported and are not passing fidelity cases. The candidate planner suite passes against the Word-authored source, including exact accepted/rejected insertion text. Internal manual-header generation diagnostics persist, while final document/numbering XML parses strictly and the supported manual-marker case passes Word.

The frozen source fixture was authored in Word and verified after save/reopen. It contains nested bullets and decimal outline numbering, a continuation across a plain paragraph, a separate restart, plain transitions, marked and unmarked noncontiguous headers, and an untouched bold sentinel. Its [source observations](../../tests/fixtures/agentic-lists/source-observations.json) record COM properties and actual numbering XML.

The Word oracle checks list/plain state, levels, labels, values and logical continuation/restart identity. The Office.js collector supports external manifests and separate artifact directories; the passing live evidence is recorded below.

Seven supported cases pass **42 independent Word checks**. The candidate Office.js lane passes **28 checks**, and its exports pass **42 Word checks**. The subsequent production lane invokes the actual `executeInsertListItem` function for its five insertion cases and also passes **28 checks**, followed by **42 Word checks**. Word host is `16.0.20430.20092` (PC). [Production Office.js evidence](2026-09-30-agentic-list-production-officejs.json), [export oracle](2026-09-30-agentic-list-production-word-oracle.json), [static-package oracle](2026-09-30-agentic-list-word.json).

Two fidelity defects are independently confirmed in Word: plain insertion leaves an extra empty paragraph after Reject All, and list-range replacement merges the original paragraphs after Reject All. The diagnostic lane has **8 passing checks and 4 expected fidelity failures**; these failures remain open and are excluded from the supported matrix. [Diagnostic evidence](2026-09-30-agentic-list-known-defects-word.json), [plain insertion issue](../library-issues/2026-09-30-list-insertion-rejection.md), [range issue](../library-issues/2026-09-30-list-range-rejection.md). No library-side fix or consumer OOXML workaround was added.

Unmarked/text-changing header conversion can produce a literal marker rather than native list numbering; its separate capability report includes a runnable reproducer. Existing native header conversion remains active. The verified unchanged manual-marker candidate passes Word, but prints internal XML-provider diagnostics and is not routed into production.

A focused review found and fixed two main-chat wiring mistakes: source-baseline declaration scope and navigation response text shape. The new dispatch suite executes the actual extracted capture/dispatch and response-building code for success and failure. It passes.

The main chat also passes its canonical baseline to list insertion. Stale or unavailable supplied baselines refuse before either a canonical or native write; callers omitting the optional baseline retain their prior contract. Production invocation tests cover baseline refusal, canonical insertion receipts, source restoration, native capability selection, unknown host failure without replay, and malformed-source refusal.

A separate offline reproduction shows the library inspector treating historical `numPr` inside `pPrChange` as active list properties. The tested production route refuses without a write or native replay. Its nested library error is `EXISTING_REVISIONS`; the consumer retains the batch wrapper `BATCH_OPERATION_FAILED` and refused receipt. No live Word verification is claimed for this case. [Historical inspection issue](../library-issues/2026-09-30-historical-list-inspection.md).

Final ordinary development/production builds pass; the development-only validation entry also builds. Production reports the existing three webpack size warnings (`taskpane.js` approximately 790 KiB). No live provider calls were made, and temporary validation registration was removed.

## Resume points

1. Consumer validation and stale-context guard committed as `7e44b5d`; actual Office.js and independent Word evidence now pass.
2. Supported source/engine/Office.js Word validation and aggregate tests/builds pass; this checkpoint preserves the verified production insertion subset and evidence.
3. Resolve the separately recorded library defects/capability gaps, then rerun the diagnostic manifest expecting full source restoration.
4. Expand deep-level/other-style/tracking-off Word coverage before routing those paths, and finish the remaining list migrations.
