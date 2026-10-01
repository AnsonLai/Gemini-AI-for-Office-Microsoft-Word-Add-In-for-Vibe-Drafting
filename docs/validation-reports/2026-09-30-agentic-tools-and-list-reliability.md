# Agentic tools and list reliability — execution record

## Consumer validation checkpoint

- Exact library version remains `@ansonlai/docx-redline-js@0.8.2`.
- All three list tools validate request types, indexes and supported enums before Word execution, then recheck indexes against the live paragraph count. Invalid targets produce `INVALID_LIST_REQUEST` with no content write.
- `edit_list` no longer silently clamps an invalid range to another paragraph.
- `convert_headers_to_list` binds replacement text to the original index before sorting; duplicate targets refuse.
- Deterministic production invocation tests cover malformed input, live range refusal, unsorted header pairing and duplicate refusal. The pure request validator and existing mutation-outcome suites pass.
- Canonical source-baseline tests confirm accepted-view text, manual line breaks and stable fingerprints. The stale-context guard checks every targeted source paragraph before batch preparation. Passing regressions cover matching source, whole-paragraph edits, range interiors, insert-before anchors, append after a changed paragraph count, unavailable baselines and mixed valid/stale batches with zero writes.
- Ordinary development and development-only Word validation entry builds pass. Optional telemetry requests are blocked by the environment and do not fail compilation.

## Canonical list migration and live testing

Candidate mappings are being tested against Word-authored fixtures. They are not wired into production list commands until numbering, accepted/rejected text and host insertion fidelity pass.

The Word oracle has been extended with list/plain state, levels, labels, values and logical continuation/restart identity checks. The Office.js collector supports external manifests and separate artifact directories. These harness changes do not establish a new passing live list result by themselves.

An initial library capability check found that an unmarked header conversion can produce a literal marker rather than native list numbering. This is being recorded separately as an upstream issue. Existing native header conversion remains active.

## Resume points

1. Consumer validation and stale-context guard checkpoint is ready; live host guard evidence remains pending.
2. Complete Word-authored multilevel fixture structure assertions after save/reopen.
3. Verify candidate list operations offline and in Word. Record unsupported library semantics separately, retain proven native functionality, and migrate only verified paths.
4. Update the plan's acceptance checks from actual evidence.
