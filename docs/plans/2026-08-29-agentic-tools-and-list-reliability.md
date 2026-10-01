# Agentic Tools and List Reliability Plan

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** In progress — contract inventory complete; exact request validation, canonical list capability checks and expanded Word fixtures underway.

**Recommended order:** 2 of 4 August plans.

## Baseline and dependencies

Use exact `@ansonlai/docx-redline-js@0.8.2` in both consumers. The [September upgrade](2026-09-09-docx-redline-js-v0.5.4-upgrade.md) and [reliability and quality gates](2026-08-29-reliability-and-quality-gates.md) are complete, including actual Office.js validation, bounded provider recovery and structured mutation outcomes. This is the next plan to execute.

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
- WP1 inventory: [actual tool contracts](../agentic-tool-contracts.md) now trace all eight content-mutating tools, argument/index conventions, library operations and host paths. Validation gaps are recorded explicitly; inventory does not certify unimplemented paths.
- WP1 request validation: all three list commands validate before entering Word and recheck indexes against the live paragraph count. Invalid requests refuse without content writes. `edit_list` no longer clamps stale indexes, and header conversion pairs supplied text with its original index before sorting. Deterministic production-invocation and pure-validator suites pass.
- WP1 targeting: a canonical prompt-time baseline is captured using the same accepted-view inspector as execution. A pre-write stale-context guard and mixed-batch no-write regressions pass; the unrelated Word-text projection is not used for comparison. Whole-paragraph, range, insertion anchor, append and unavailable-baseline cases are covered.
- WP2 investigation: v0.8.2 has no dedicated semantic list mutation API. Structured redline operations and public list helpers are being tested for numbering identity, indentation and restart/continuation. A declared operation name alone does not establish that the standalone runner supports its intended semantics.
- WP3 preparation: the Word oracle now accepts per-state list/plain expectations, levels, labels/values and logical continuation/restart identity groups. Word-authored fixtures and live verification remain pending.
- WP3 host harness: the Office.js collector can consume an external fixture manifest and a separate artifact directory. The validation entry maps `agenticRequest` through the candidate planner; ordinary builds still omit the validation entry. New list host evidence is pending.
- WP1 actual Word evidence: Office.js context validation passes 8 checks, including zero-write stale-context refusal after an intervening edit; independent Word reopening passes 12 checks. New list host evidence remains pending.
- Library capability gap: unmarked/text-changing header conversion has no verified canonical mapping in 0.8.2. [Separate library report](../library-issues/2026-09-30-canonical-list-operations.md); native header conversion remains active.
- Next checkpoint: commit the inventory and validated request/mapper work with test results. Preserve working native paths until a replacement has equivalent fidelity evidence; report any library limitation separately.

### WP1 — Inventory and align actual tool contracts

- Trace every mutating tool in `commands/agentic-tools.js` to its mapper, library operation and transport adapter.
- Record actual names/argument shapes. The previous example `replace_paragraph` mapping was not the implemented localized redline schema; current prompting uses `edit_paragraph` changes with `replacements`.
- Preserve full `newContent` support and 1-based replacement occurrence semantics.
- Cover missing/repeated exact text, empty replacement strings, stale context and mixed valid/invalid batches with explicit no-write expectations.
- Preserve library receipts and generic errors under the reliability plan's outcome contract; withdraw the old illustrative response JSON.

### WP2 — Migrate remaining list commands to canonical operations

- Identify commands still using paragraph/scope adapters and native list metadata.
- Preserve relative indentation: -1 shallower, 0 same, +1 deeper, with Word levels clamped to 0–8.
- Use supported library normalization/Markdown generation where needed and structured-content operations for numbering creation/reconciliation.
- Resolve batch targets from one initial source; avoid per-operation proxy loops in migrated application paths.
- Preserve numbering identity, restart/continuation behavior and untouched list items.
- Remove helpers only when all callers have migrated and observable behavior has coverage. Maintain the no-legacy guard for actually retired modules.

### WP3 — Expand independent list and tool validation

- Add Word-authored fixtures for nested bullets/numbering, insertion before/after, continuation/restart and list/plain-paragraph transitions.
- Assert exact accepted/rejected text, numbering relationships and untargeted formatting offline.
- Exercise migrated paths in real Word, including native insertion and actual Office.js transport where relevant; reuse the supervised harness and reference project's save/reopen methodology.
- Confirm ambiguous/unsupported edits refuse before a write and mixed batches are atomic.
- Keep live model evaluations separate from deterministic correctness checks. Measure prompt reduction on fixed cases if useful; withdraw the unsupported universal 60–90% claim.

## Acceptance

- [x] Localized redline prompting, validation and atomic batch execution implemented.
- [ ] Actual tool-to-operation mapping/outcome contracts documented and covered.
- [ ] Remaining list command paths use canonical operations with preserved indentation/numbering semantics.
- [ ] Nested-list Accept All/Reject All and real Word roundtrips pass.
- [ ] No helper deleted while active callers rely on it.

**Next:** [Package boundaries and integrations](2026-08-29-package-boundaries-and-integrations.md), using migrated command paths to define the portable consumer boundary.
