# Agentic Tools and List Reliability Plan

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** Partially implemented by upgrade WP2–WP4; remaining tool/list migration and host coverage pending.

**Recommended order:** 2 of 4 August plans.

## Baseline and dependencies

Use exact `@ansonlai/docx-redline-js@0.8.2` in both consumers. The [September upgrade](2026-09-09-docx-redline-js-v0.5.4-upgrade.md) is near complete; actual Office.js transport remains its final gate. Establish recovery/outcome rules in [reliability and quality gates](2026-08-29-reliability-and-quality-gates.md) first.

Accuracy comes first: target against an immutable source, fail closed on ambiguity, preserve formatting/package parts, and pass canonical operations to the library. Word proxies belong in transport adapters.

## Completed / inherited work

| Original proposal | Current disposition |
| --- | --- |
| Delete unused `word-structured-list.js` and export | Done in upgrade WP2 |
| Remove `detectRequestedContentKind`, direct reconciliation fallback and redline table heuristics | Done |
| Replace iterative redline application with one atomic batch | Done in upgrade WP3 |
| Prefer localized replacements in prompt/schema | Done in upgrade WP4; full-paragraph edits remain supported |
| Validate exact find text and repeated occurrences | Done, including index correction before replacement validation |
| Delete `normalizeListItemsWithLevels` / `buildListMarkdown` | Withdrawn: supported library imports with active callers |
| Delete `list-level-utils.js` | Withdrawn: active relative-indent/clamping adapter covered by tests |
| Delete all older Word operation helpers | Deferred until remaining callers migrate |

Evidence: `tests/agentic_list_generation_tests.mjs`, `tests/insert_list_item_level_tests.mjs`, `tests/word_list_binding_regression_tests.mjs`, `tests/change_validation_tests.mjs` and redline batch suites.

## Remaining work packages

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
