# Browser Demo Plumbing Extraction Implementation Plan

## Status

✅ **Completed in the current tree.** The browser demo now consumes the shared
standalone plumbing helpers from `@ansonlai/docx-redline-js`; the original
implementation and test locations were absorbed into the extracted package.

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Move duplicated OOXML plumbing/zip artifact/targeting support out of `browser-demo/demo.js` into shared reconciliation standalone support without changing behavior.

**Architecture:** Extract pure OOXML and package-manipulation helpers into reconciliation services, re-export through `standalone.js`, and replace local `demo.js` copies with imports. Keep runtime flow in `demo.js` intact and preserve existing logs and operation order.

**Tech Stack:** ES modules, reconciliation standalone exports, JSZip-compatible zip API, DOMParser/XMLSerializer via xml-adapter.

---

### Task 1: Add failing tests for extracted shared helpers

**Files:**
- Create: `tests/standalone_docx_plumbing_tests.mjs`

1. Add tests for:
- OOXML output extraction (`pkg:package`, `w:document`, fragment)
- Nested paragraph sanitization in table cells
- Section-properties normalization (`w:sectPr` ordering)
- Numbering/comments artifact wiring + validation
2. Run test file and confirm failure because helpers are not yet exported from standalone.

### Task 2: Implement shared extraction/plumbing package helpers

**Files:**
- Create: `src/taskpane/modules/reconciliation/services/standalone-docx-plumbing.js`
- Modify: `src/taskpane/modules/reconciliation/standalone.js`
- Modify: `src/taskpane/modules/reconciliation/index.js`

1. Implement helper exports:
- `normalizeBodySectionOrderStandalone`
- `sanitizeNestedParagraphsInTables`
- `extractReplacementNodesFromOoxml`
- `ensureNumberingArtifactsInZip`
- `ensureCommentsArtifactsInZip`
- `validateDocxPackage`
2. Re-export through `standalone.js` and `index.js`.
3. Run the new test file and confirm pass.

### Task 3: Switch browser demo to shared helpers and remove duplicates

**Files:**
- Modify: `browser-demo/demo.js`

1. Replace local implementations with standalone imports.
2. Keep existing call flow and behavior unchanged.
3. Remove now-unused constants/helpers.

### Task 4: Verify refactor safety

**Files:**
- Modify (if needed): `browser-demo/README.md`

1. Run targeted tests:
- `node tests/standalone_docx_plumbing_tests.mjs`
- `node tests/standalone_smoke.mjs`
- `node tests/include_numbering_behavior.mjs`
2. Run a lightweight lint/build sanity command if needed.
3. Report results with evidence.
