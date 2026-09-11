# `docx-redline-js` v0.5.4 automated validation

**Date:** 2026-09-09  
**Scope:** WP2–WP4 automated validation  
**Result:** Passed; interactive browser and Word Desktop checks remain manual.

## Dependency resolution

- Root add-in: one exact `@ansonlai/docx-redline-js@0.5.4`.
- MCP server: one exact `@ansonlai/docx-redline-js@0.5.4`.
- Both manifests and npm-generated lockfiles pin exact `0.5.4`.

## Automated results

- Serialized runnable Node inventory: **28 passed, 0 failed**.
- `npm run build:dev`: passed.
- `npm run build`: passed with the three existing webpack asset-size warnings
  for the task-pane bundle.
- `npm run validate`: passed; the Office manifest is valid.
- Relevant browser, MCP, and Word integration files passed `node --check`.

The serialized inventory excluded these non-test, live, fixture-rewriting, or
desktop-dependent entry points:

- `tests/setup-xml-provider.mjs`;
- `tests/evals/run-evals.mjs`;
- `tests/phase4/golden-guardrail.mjs`;
- `tests/phase4/perf-harness.mjs`;
- `tests/word-desktop/list-regression.mjs`.

## Safety and workflow coverage

- Package contracts cover canonical accepted-view text and fresh fingerprints,
  structured failures, commented-deletion refusal, atomic rollback, receipts,
  same-author merge, foreign-author refusal, revision-ID uniqueness,
  accept/reject symmetry, hyperlinks, structural text, and sanitization policy.
- Host adapters preserve stable codes and diagnostic metadata without treating
  failures as no-ops or selecting a local fallback algorithm.
- The MCP workflow smoke verifies that a stale edit returns
  `TARGET_NOT_FOUND` without changing XML or session dirty state, then verifies
  a successful tracked edit through save, reopen, package validation, and close.
- Numbering behavior is exercised through the v0.5.4 public package-builder API
  for included, excluded, and default numbering modes.

## Manual follow-ups

The validation environment exposed neither an in-app browser instance nor the
Word Desktop controller native pipe, so the following checks were not executed:

1. Browser demo UI flows for sequential edits, redlines on/off, comments,
   lists/tables, repeated anchors, and intentional non-mutating failure.
2. Word Desktop open-without-repair, Reviewing Pane grouping, Accept All,
   Reject All, save, and reopen against representative generated files with
   revisions, comments, hyperlinks, bookmarks, structural characters, lists,
   tables, and section properties.

These are environment-dependent acceptance checks. Their absence does not
change the passing automated package, adapter, MCP workflow, build, syntax, or
manifest results above.
