# Agentic Tools and List Reliability Plan

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** In progress — WP1 complete; WP2 verified insertion subset migrated; WP3 supported Word matrix passes, with two separately confirmed library fidelity defects blocking broader migration.

**Recommended order:** 2 of 4 August plans.

## Baseline and dependencies

Use exact `@ansonlai/docx-redline-js@0.8.2` in both consumers. The [September upgrade](completed/2026-09-09-docx-redline-js-v0.5.4-upgrade.md) and [reliability and quality gates](completed/2026-08-29-reliability-and-quality-gates.md) are complete, including actual Office.js validation, bounded provider recovery and structured mutation outcomes. This is the next plan to execute.

Accuracy comes first: target against an immutable source, fail closed on ambiguity, preserve formatting/package parts, and pass canonical operations to the library. Word proxies belong in transport adapters.

## Completed / inherited work

| Original proposal | Current disposition |
| --- | --- |
| Delete unused `word-structured-list.js` and export | Done in upgrade WP2 |
| Remove `detectRequestedContentKind`, direct reconciliation fallback and redline table heuristics | Done |
| Replace iterative redline application with one atomic batch | Done in upgrade WP3 |
| Prefer localized replacements in prompt/schema | Done in upgrade WP4; full-paragraph edits remain supported |
| Validate exact find text and repeated occurrences | Done, including index correction before replacement validation |
| Structured mutation outcomes and safe recovery | Done in reliability WP3; remaining tool mapping/atomic list migration still pending |
| Delete `normalizeListItemsWithLevels` / `buildListMarkdown` | Withdrawn: supported library imports with active callers |
| Delete `list-level-utils.js` | Withdrawn: active relative-indent/clamping adapter covered by tests |
| Delete all older Word operation helpers | Deferred until remaining callers migrate |

Evidence: `tests/agentic_list_generation_tests.mjs`, `tests/insert_list_item_level_tests.mjs`, `tests/word_list_binding_regression_tests.mjs`, `tests/change_validation_tests.mjs` and redline batch suites.

## Remaining work packages

### Execution record (2026-09-30)

- Reliability milestone committed as `e6a5df7` before starting this plan.
- Consumer validation and stale-context guard milestone committed as `7e44b5d`. [Execution and validation report](../validation-reports/2026-09-30-agentic-tools-and-list-reliability.md).
- Word source/host checkpoint committed as `ea07a9a`; table contracts and structured failures checkpoint committed as `3bbc7e4`.
- WP1 inventory: [actual tool contracts](../agentic-tool-contracts.md) now trace all eight content-mutating tools, argument/index conventions, library operations and host paths. Validation gaps are recorded explicitly; inventory does not certify unimplemented paths.
- WP1 request validation: all three list commands validate before entering Word and recheck indexes against the live paragraph count. Invalid requests refuse without content writes. `edit_list` no longer clamps stale indexes, and header conversion pairs supplied text with its original index before sorting. Deterministic production-invocation and pure-validator suites pass.
- WP1 targeting: a canonical prompt-time baseline is captured using the same accepted-view inspector as execution. A pre-write stale-context guard and mixed-batch no-write regressions pass; the unrelated Word-text projection is not used for comparison. Whole-paragraph, range, insertion anchor, append and unavailable-baseline cases are covered.
- WP2 implementation: verified bullet/decimal insertion uses canonical operations; broader list replacement and header conversion await the separately recorded library dependencies below.
- WP3 verification: frozen Word-authored fixtures pass seven supported cases and 42 independent Word checks, including numbering identity, indentation and restart/continuation.
- WP3 host harness: the Office.js collector supports external manifests and separate artifact directories. Its production insertion lane passes 28 checks; exported packages pass 42 independent Word checks. Ordinary builds omit the validation entry.
- WP1 actual Word evidence: Office.js context validation passes 8 checks, including zero-write stale-context refusal after an intervening edit; independent Word reopening passes 12 checks.
- Library capability gap: unmarked/text-changing header conversion has no verified canonical mapping in 0.8.2. [Separate library report](../library-issues/2026-09-30-canonical-list-operations.md); native header conversion remains active.
- Final checkpoint includes the verified insertion migration, 44 passing offline suites, successful builds and durable Word evidence. Remaining work is recorded below for resumption.
- WP3 follow-up: expanded the frozen-source matrix to ten cases, adding bullet root insertion and decimal insertion at continuation/restart anchors. Static packages pass 60 independent Word checks; actual Office.js passes 40 checks across eight production insertion routes and two candidate routes. Export verification is recorded in the execution report.
- The September upgrade closure was clarified in `1680349`; all of its acceptance criteria remain complete. Current agentic work does not reopen that plan.

### WP1 — Inventory and align actual tool contracts

**Status: Complete.** Inventory and strict list/table argument validation are covered by production invocation tests. Table schema now declares its canonical nested-array shape; compatibility inputs are normalized without truncation. Missing-key failures and navigation results have explicit success/error fields. Canonical prompt-time redline baselines refuse stale targets atomically; actual Office.js refusal and independent Word reopening pass. Remaining structural-tool freshness/atomicity limits are recorded for WP2 and future library capabilities.

- Trace every mutating tool in `commands/agentic-tools.js` to its mapper, library operation and transport adapter.
- Record actual names/argument shapes. The previous example `replace_paragraph` mapping was not the implemented localized redline schema; current prompting uses `edit_paragraph` changes with `replacements`.
- Preserve full `newContent` support and 1-based replacement occurrence semantics.
- Cover missing/repeated exact text, empty replacement strings, stale context and mixed valid/invalid batches with explicit no-write expectations.
- Preserve library receipts and generic errors under the reliability plan's outcome contract; withdraw the old illustrative response JSON.

### WP2 — Migrate remaining list commands to canonical operations

**Status: Partially implemented.** With redlining enabled, `insert_list_item` uses one canonical body batch for active bullet/decimal lists at source and resolved levels 0–1. Relative outdent needs a following root item in the same numbering instance. Plain anchors, other levels/formats, tracking-off requests and unsupported outdent contexts retain native execution. Capability selection occurs before engine preparation; an engine or host failure never triggers native replay. General `edit_list` and header conversion remain on established paths pending broader capability/fidelity work.

The main chat passes its canonical source baseline into list insertion; stale or unavailable supplied baselines refuse before canonical or native writes. The optional baseline remains backward compatible for callers that omit it. Final offline validation passes 44 suites, with the known library defects reported separately; ordinary development/production builds pass.

Separately reported library dependencies:

- [Plain insertion Reject All leaves an empty paragraph](../library-issues/2026-09-30-list-insertion-rejection.md), independently confirmed in Word.
- [List range Reject All merges source paragraphs](../library-issues/2026-09-30-list-range-rejection.md), independently confirmed in Word.
- [Unmarked/text-changing header conversion lacks a canonical mapping](../library-issues/2026-09-30-canonical-list-operations.md).
- [Historical list properties are inspected as active numbering](../library-issues/2026-09-30-historical-list-inspection.md), reproduced offline; the operation refuses without a write or native replay.

- Identify commands still using paragraph/scope adapters and native list metadata.
- Preserve relative indentation: -1 shallower, 0 same, +1 deeper, with Word levels clamped to 0–8.
- Use supported library normalization/Markdown generation where needed and structured-content operations for numbering creation/reconciliation.
- Resolve batch targets from one initial source; avoid per-operation proxy loops in migrated application paths.
- Preserve numbering identity, restart/continuation behavior and untouched list items.
- Remove helpers only when all callers have migrated and observable behavior has coverage. Maintain the no-legacy guard for actually retired modules.

### WP3 — Expand independent list and tool validation

**Status: Expanded supported matrix passes; broader acceptance remains open.** Ten supported cases pass 60 independent Word checks. Actual Office.js passes 40 checks, including eight production `insert_list_item` routes, no-op/refusal and mixed-batch zero-write behavior; its exports pass 60 independent Word checks. Two rejected-view defects have a separate diagnostic manifest and failed Word evidence; they are not counted as passing fidelity cases. Tests report these known defects explicitly through `npm test`. Deeper source/resolved levels, Roman-style fallback and tracking-off mode now have deterministic production invocation coverage, including tracking restoration; these additional tests use mocked Word proxies and do not close their live fidelity gates.

- Add Word-authored fixtures for nested bullets/numbering, insertion before/after, continuation/restart and list/plain-paragraph transitions.
- Assert exact accepted/rejected text, numbering relationships and untargeted formatting offline.
- Exercise migrated paths in real Word, including native insertion and actual Office.js transport where relevant; reuse the supervised harness and reference project's save/reopen methodology.
- Confirm ambiguous/unsupported edits refuse before a write and mixed batches are atomic.
- Keep live model evaluations separate from deterministic correctness checks. Measure prompt reduction on fixed cases if useful; withdraw the unsupported universal 60–90% claim.

## Acceptance

- [x] Localized redline prompting, validation and atomic batch execution implemented.
- [x] Actual tool-to-operation mapping/outcome contracts documented and covered, with unsupported/native structural paths explicit.
- [ ] Remaining list command paths use canonical operations with preserved indentation/numbering semantics.
- [ ] Nested-list Accept All/Reject All and real Word roundtrips pass.
- [x] No helper deleted while active callers rely on it.

Current resume order: fix the separately reported library fidelity defects/canonical header capability; upgrade the pinned package once released; rerun both supported and diagnostic manifests; expand deep-level/other-style/tracking-off coverage; then finish remaining list migrations and retire helpers only when unused.

The [OOXML performance](completed/2026-08-29-oxml-engine-and-performance.md) and [package boundaries](completed/2026-08-29-package-boundaries-and-integrations.md) plans are complete. Resume this plan with its separately tracked library fidelity and capability follow-ups, then the remaining native coverage and list migrations.
