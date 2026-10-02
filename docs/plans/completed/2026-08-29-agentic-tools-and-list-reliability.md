# Agentic Tools and List Reliability Plan

> **Retrospective ownership review (2026-09-30):** See the [library offload review](../../library-offload-review.md) for measured code changes, responsibilities delegated to the library, and remaining consumer/native routes. Version pins and validation counts in this closure record describe their dated checkpoints; current consumers pin exact 0.8.3. All five original plans are closed; remaining canonical list work is in the [fresh follow-up](../2026-09-30-canonical-list-migration-follow-up.md).

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** Closed with deferred scope — verified consumer work and 0.8.3 fidelity gates complete; full canonical migration transferred to a fresh follow-up at the user's request.

**Recommended order:** 2 of 4 August plans.

## Closure and transferred work

The user explicitly requested closure on 2026-09-30 despite the remaining
canonical capabilities. This supersedes the earlier keep-open decision below.
WP1 and current-route WP3 are complete; remaining WP2 migration and its
acceptance criterion are transferred to the [Canonical List Migration Follow-up](../2026-09-30-canonical-list-migration-follow-up.md).
This archive records a scoped closure, not completion of full canonical migration.

## Baseline and dependencies

Use exact `@ansonlai/docx-redline-js@0.8.3` in both consumers. The [September upgrade](2026-09-09-docx-redline-js-v0.5.4-upgrade.md) and [reliability and quality gates](2026-08-29-reliability-and-quality-gates.md) are complete, including actual Office.js validation, bounded provider recovery and structured mutation outcomes. All five original dated plans are now closed; remaining work is in the fresh follow-up above.

Accuracy comes first: target against an immutable source, fail closed on ambiguity, preserve formatting/package parts, and pass canonical operations to the library. Word proxies belong in transport adapters.

**Earlier closure decision (2026-09-30, superseded):** The user initially required this plan to remain open until upstream fixes enable full canonical migration. Release 0.8.3 resolves the fidelity defects but excludes canonical plain-to-list conversion and list-format changes. The later explicit request closes this plan with those capabilities transferred to the fresh follow-up. The library's `docs/plans/2026-09-30-canonical-list-operations.md` remains its upstream implementation plan.

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
- Consumer validation and stale-context guard milestone committed as `7e44b5d`. [Execution and validation report](../../validation-reports/2026-09-30-agentic-tools-and-list-reliability.md).
- Word source/host checkpoint committed as `ea07a9a`; table contracts and structured failures checkpoint committed as `3bbc7e4`.
- WP1 inventory: [actual tool contracts](../../agentic-tool-contracts.md) now trace all eight content-mutating tools, argument/index conventions, library operations and host paths. Validation gaps are recorded explicitly; inventory does not certify unimplemented paths.
- WP1 request validation: all three list commands validate before entering Word and recheck indexes against the live paragraph count. Invalid requests refuse without content writes. `edit_list` no longer clamps stale indexes, and header conversion pairs supplied text with its original index before sorting. Deterministic production-invocation and pure-validator suites pass.
- WP1 targeting: a canonical prompt-time baseline is captured using the same accepted-view inspector as execution. A pre-write stale-context guard and mixed-batch no-write regressions pass; the unrelated Word-text projection is not used for comparison. Whole-paragraph, range, insertion anchor, append and unavailable-baseline cases are covered.
- WP2 implementation: verified bullet/decimal insertion uses canonical operations; broader list replacement and header conversion await the separately recorded library dependencies below.
- WP3 verification: frozen Word-authored fixtures pass seven supported cases and 42 independent Word checks, including numbering identity, indentation and restart/continuation.
- WP3 host harness: the Office.js collector supports external manifests and separate artifact directories. Its production insertion lane passes 28 checks; exported packages pass 42 independent Word checks. Ordinary builds omit the validation entry.
- WP1 actual Word evidence: Office.js context validation passes 8 checks, including zero-write stale-context refusal after an intervening edit; independent Word reopening passes 12 checks.
- Library capability gap: unmarked/text-changing header conversion has no verified canonical mapping in 0.8.2. [Separate library report](../../library-issues/2026-09-30-canonical-list-operations.md); native header conversion remains active.
- Final checkpoint includes the verified insertion migration, 44 passing offline suites, successful builds and durable Word evidence. Remaining work is recorded below for resumption.
- WP3 follow-up: expanded the frozen-source matrix to ten cases, adding bullet root insertion and decimal insertion at continuation/restart anchors. Static packages pass 60 independent Word checks; actual Office.js passes 40 checks across eight production insertion routes and two candidate routes. Export verification is recorded in the execution report.
- The September upgrade closure was clarified in `1680349`; all of its acceptance criteria remain complete. Current agentic work does not reopen that plan.
- Package boundaries and performance are complete (`b8f23b9`, `42a62df`). The aggregate now has 49 passing offline suites; that does not certify the remaining native list fidelity gates.
- Final consumer follow-up passes five native Word cases: deeper insertion, outdent from a deeper source, upper/lower Roman numbering and tracking disabled with prior-mode restoration. Actual Office.js passes 20 checks; independent Word source/tracked/Accept All/Reject All passes 20 checks. Five engine-reference entries are explicitly not applicable and are not counted as passes. Documentation consolidation is complete; full canonical migration remains gated on the [separate upstream reports](../../library-issues/README.md). The offline aggregate passes 50 suites, zero failures, with four exclusions.

**Unpublished library review (2026-09-30):** the local tree fixes historical
inspection and both original paragraph-boundary cases; the existing consumer
standalone/merge output passes 12 independent Word checks. Focused regression
tests pass. Full canonical conversion/remapping remains a proposed library
plan, and review found an all-empty-range rejection defect plus pre-existing
public-facade numbering corruption. See the [review and resume gates](../../validation-reports/2026-09-30-unpublished-library-list-review.md).
No package pin or production route changed. At that checkpoint this plan remained open pending
released fixes, capability implementation and downstream host verification.

### WP1 — Inventory and align actual tool contracts

**Status: Complete.** Inventory and strict list/table argument validation are covered by production invocation tests. Table schema now declares its canonical nested-array shape; compatibility inputs are normalized without truncation. Missing-key failures and navigation results have explicit success/error fields. Canonical prompt-time redline baselines refuse stale targets atomically; actual Office.js refusal and independent Word reopening pass. Remaining structural-tool freshness/atomicity limits are recorded for WP2 and future library capabilities.

- Trace every mutating tool in `commands/agentic-tools.js` to its mapper, library operation and transport adapter.
- Record actual names/argument shapes. The previous example `replace_paragraph` mapping was not the implemented localized redline schema; current prompting uses `edit_paragraph` changes with `replacements`.
- Preserve full `newContent` support and 1-based replacement occurrence semantics.
- Cover missing/repeated exact text, empty replacement strings, stale context and mixed valid/invalid batches with explicit no-write expectations.
- Preserve library receipts and generic errors under the reliability plan's outcome contract; withdraw the old illustrative response JSON.

### WP2 — Migrate remaining list commands to canonical operations

**0.8.3 checkpoint:** All reported rejection/inspection/facade numbering bugs
are fixed and verified. The root and MCP pins upgraded together in `148c96b`.
Full migration remains incomplete because plain-to-list/header conversion and
list-format changes are explicitly unsupported in this release. Existing
native/legacy routes remain; no route was widened solely from release notes.
The defect bullets and earlier counts below record historical 0.8.2 work.

**Status: Partially implemented.** With redlining enabled, `insert_list_item` uses one canonical body batch for active bullet/decimal lists at source and resolved levels 0–1. Relative outdent needs a following root item in the same numbering instance. Plain anchors, other levels/formats, tracking-off requests and unsupported outdent contexts retain native execution. Capability selection occurs before engine preparation; an engine or host failure never triggers native replay. General `edit_list` and header conversion remain on established paths pending broader capability/fidelity work.

The main chat passes its canonical source baseline into list insertion; stale or unavailable supplied baselines refuse before canonical or native writes. The optional baseline remains backward compatible for callers that omit it. Final offline validation passes 44 suites, with the known library defects reported separately; ordinary development/production builds pass.

Separately reported library dependencies:

- [Plain insertion Reject All leaves an empty paragraph](../../library-issues/completed/2026-09-30-list-insertion-rejection.md), independently confirmed in Word.
- [List range Reject All merges source paragraphs](../../library-issues/completed/2026-09-30-list-range-rejection.md), independently confirmed in Word.
- [Unmarked/text-changing header conversion lacks a canonical mapping](../../library-issues/2026-09-30-canonical-list-operations.md).
- [Historical list properties are inspected as active numbering](../../library-issues/completed/2026-09-30-historical-list-inspection.md), reproduced offline; the operation refuses without a write or native replay.

- Identify commands still using paragraph/scope adapters and native list metadata.
- Preserve relative indentation: -1 shallower, 0 same, +1 deeper, with Word levels clamped to 0–8.
- Use supported library normalization/Markdown generation where needed and structured-content operations for numbering creation/reconciliation.
- Resolve batch targets from one initial source; avoid per-operation proxy loops in migrated application paths.
- Preserve numbering identity, restart/continuation behavior and untouched list items.
- Remove helpers only when all callers have migrated and observable behavior has coverage. Maintain the no-legacy guard for actually retired modules.

### WP3 — Expand independent list and tool validation

**0.8.3 checkpoint: Complete for the current routes.** All twelve supported
cases now pass, including the two former diagnostic defects. Actual Office.js
passes 48 checks for that matrix and 20 checks for five native cases. The
combined exports pass 92 independent Word checks, with five engine-reference
entries explicitly not applicable. The public facade passes twelve additional
Word checks. Single-empty/all-empty rejection and facade numbering have focused
regressions; the full offline aggregate passes 51 suites, zero failures, four
exclusions. Development validation and production builds pass. See the
[0.8.3 evidence](../../validation-reports/2026-09-30-docx-redline-v083.md).
The earlier 0.8.2 counts and failures below remain historical.

**Status: Expanded supported and native matrices pass; broader acceptance remains open.** Ten supported cases pass 60 independent Word checks. Actual Office.js passes 40 checks, including eight production `insert_list_item` routes, no-op/refusal and mixed-batch zero-write behavior; its exports pass 60 independent Word checks. Two rejected-view defects have a separate diagnostic manifest and failed Word evidence; they are not counted as passing fidelity cases. Tests report these known defects explicitly through `npm test`. Five additional native production cases cover deeper source/resolved levels, upper/lower Roman numbering and tracking-off insertion/restoration: 20 actual Office.js checks and 20 independent Word checks pass. Prepared sources are frozen before mutation; accepted/rejected text, labels, levels, logical list identity and untouched bold formatting are checked independently. These results certify this native subset, not full canonical migration.

- Add Word-authored fixtures for nested bullets/numbering, insertion before/after, continuation/restart and list/plain-paragraph transitions.
- Assert exact accepted/rejected text, numbering relationships and untargeted formatting offline.
- Exercise migrated paths in real Word, including native insertion and actual Office.js transport where relevant; reuse the supervised harness and reference project's save/reopen methodology.
- Confirm ambiguous/unsupported edits refuse before a write and mixed batches are atomic.
- Keep live model evaluations separate from deterministic correctness checks. Measure prompt reduction on fixed cases if useful; withdraw the unsupported universal 60–90% claim.

## Acceptance

- [x] Localized redline prompting, validation and atomic batch execution implemented.
- [x] Actual tool-to-operation mapping/outcome contracts documented and covered, with unsupported/native structural paths explicit.
- [ ] Transferred to the fresh follow-up: remaining list command paths use canonical operations with preserved indentation/numbering semantics.
- [x] Nested-list Accept All/Reject All and real Word roundtrips pass for the current verified canonical/candidate/native routes on 0.8.3.
- [x] No helper deleted while active callers rely on it.

Current resume order: implement the remaining canonical conversion and list-format capabilities in the library; upgrade after their release; migrate the remaining consumer command paths and add corresponding live cases; retire helpers only when unused. The 0.8.3 fidelity upgrade and supported/native validation are complete.

The [OOXML performance](2026-08-29-oxml-engine-and-performance.md) and [package boundaries](2026-08-29-package-boundaries-and-integrations.md) plans are complete. This plan is now closed with the remaining canonical migration transferred to the fresh follow-up above.
