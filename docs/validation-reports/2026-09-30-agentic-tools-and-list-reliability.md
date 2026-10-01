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

Candidate mappings are being tested against Word-authored fixtures. They are not wired into production list commands until numbering, accepted/rejected text and host insertion fidelity pass.

The current offline aggregate passes **41 suites**, zero failures, with four entrypoints excluded and the historical 0.5.4 package-behavior skip reported separately. The candidate planner suite passes against the Word-authored source, including exact accepted/rejected insertion text; manual-header generation emits XML-provider diagnostics that require package and host investigation before migration.

The frozen source fixture was authored in Word and verified after save/reopen. It contains nested bullets and decimal outline numbering, a continuation across a plain paragraph, a separate restart, plain transitions, marked and unmarked noncontiguous headers, and an untouched bold sentinel. Its [source observations](../../tests/fixtures/agentic-lists/source-observations.json) record COM properties and actual numbering XML.

The Word oracle has been extended with list/plain state, levels, labels, values and logical continuation/restart identity checks. The Office.js collector supports external manifests and separate artifact directories. These harness changes do not establish a new passing live list result by themselves.

An initial library capability check found that an unmarked header conversion can produce a literal marker rather than native list numbering. This is being recorded separately as an upstream issue. Existing native header conversion remains active.

## Resume points

1. Consumer validation and stale-context guard committed as `7e44b5d`; actual Office.js and independent Word evidence now pass.
2. Complete Word-authored multilevel fixture structure assertions after save/reopen.
3. Verify candidate list operations offline and in Word. Record unsupported library semantics separately, retain proven native functionality, and migrate only verified paths.
4. Update the plan's acceptance checks from actual evidence.
