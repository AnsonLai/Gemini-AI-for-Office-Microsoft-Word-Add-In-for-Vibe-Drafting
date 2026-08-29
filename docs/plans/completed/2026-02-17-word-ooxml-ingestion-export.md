# Word OOXML Ingestion Export Implementation Plan

## Status

✅ **Completed in the current tree.** The browser demo consumes
`ingestWordOoxmlToMarkdown` from `@ansonlai/docx-redline-js`, whose public
package documentation includes the OOXML-to-Markdown ingestion surface. The
original implementation and test locations were absorbed into the extracted
package.

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add Word OOXML ingestion helpers that export readable plain text and basic markdown, with standalone/index exports and basic regression tests.

**Architecture:** Introduce a new ingestion export module that parses `w:p`/`w:r` structures and renders two outputs: plain text and markdown. Keep rendering conservative and deterministic by using only obvious paragraph/run signals (heading style, list properties, bold, italics). Re-export from public entry points for standalone and shared usage.

**Tech Stack:** JavaScript (ES modules), OOXML DOM parsing via existing reconciliation XML adapters, Node-based `.mjs` tests with `assert`.

---

### Task 1: Add Failing Tests For Plain Text + Markdown Export

**Files:**
- Create: `tests/standalone_ingestion_export_tests.mjs`
- Reference: `tests/setup-xml-provider.mjs`
- Reference: `src/taskpane/modules/reconciliation/standalone.js`

**Step 1: Write the failing test**

```js
import './setup-xml-provider.mjs';
import assert from 'assert';
import {
  ingestWordOoxmlToPlainText,
  ingestWordOoxmlToMarkdown
} from '../src/taskpane/modules/reconciliation/standalone.js';

// Add test OOXML samples with Heading1, numPr, and run-level w:b/w:i.
// Assert text/markdown shape and warning behavior.
```

**Step 2: Run test to verify it fails**

Run: `node tests/standalone_ingestion_export_tests.mjs`  
Expected: FAIL with missing export/function error.

**Step 3: Commit**

```bash
git add tests/standalone_ingestion_export_tests.mjs
git commit -m "test: add failing tests for standalone OOXML text and markdown ingestion"
```

### Task 2: Implement Word OOXML Plain Text + Markdown Export Module

**Files:**
- Create: `src/taskpane/modules/reconciliation/pipeline/ingestion-export.js`
- Reference: `src/taskpane/modules/reconciliation/adapters/xml-adapter.js`
- Reference: `src/taskpane/modules/reconciliation/core/xml-query.js`
- Reference: `src/taskpane/modules/reconciliation/core/types.js`

**Step 1: Write minimal implementation**

```js
export function ingestWordOoxmlToPlainText(ooxml, options = {}) { /* parse + paragraph rendering */ }
export function ingestWordOoxmlToMarkdown(ooxml, options = {}) { /* parse + heading/list/bold/italic */ }
```

Implementation requirements:
- Parse with existing adapter parser.
- Never throw for malformed input; return warning array.
- Plain text:
  - Strip tags through run traversal.
  - Preserve paragraph readability using paragraph separators.
- Markdown:
  - Heading from `w:pStyle` (`Heading1..Heading6`).
  - List prefix from `w:numPr` + basic depth indentation.
  - Run formatting from `w:rPr` (`w:b`, `w:i`) only.

**Step 2: Run test to verify it still fails only due missing exports**

Run: `node tests/standalone_ingestion_export_tests.mjs`  
Expected: FAIL with import/export mismatch until entrypoints are updated.

**Step 3: Commit**

```bash
git add src/taskpane/modules/reconciliation/pipeline/ingestion-export.js
git commit -m "feat: add Word OOXML ingestion export helpers"
```

### Task 3: Export Helpers Through Public Entrypoints

**Files:**
- Modify: `src/taskpane/modules/reconciliation/standalone.js`
- Modify: `src/taskpane/modules/reconciliation/index.js`

**Step 1: Wire exports**

```js
export { ingestWordOoxmlToPlainText, ingestWordOoxmlToMarkdown } from './pipeline/ingestion-export.js';
```

**Step 2: Run tests to verify pass**

Run: `node tests/standalone_ingestion_export_tests.mjs`  
Expected: PASS with all assertions green.

**Step 3: Run nearby regression checks**

Run: `node tests/standalone_smoke.mjs`  
Expected: PASS.

Run: `node tests/no_word_api_standalone_check.mjs`  
Expected: PASS.

**Step 4: Commit**

```bash
git add src/taskpane/modules/reconciliation/standalone.js src/taskpane/modules/reconciliation/index.js
git commit -m "feat: export standalone Word OOXML text and markdown ingestion helpers"
```

### Task 4: Final Verification

**Files:**
- Verify: `tests/standalone_ingestion_export_tests.mjs`
- Verify: `tests/standalone_smoke.mjs`
- Verify: `tests/no_word_api_standalone_check.mjs`

**Step 1: Run complete verification command set**

Run:
```bash
node tests/standalone_ingestion_export_tests.mjs
node tests/standalone_smoke.mjs
node tests/no_word_api_standalone_check.mjs
```

Expected:
- all commands exit `0`
- all output lines start with `PASS:`

**Step 2: Commit**

```bash
git add tests/standalone_ingestion_export_tests.mjs src/taskpane/modules/reconciliation/pipeline/ingestion-export.js src/taskpane/modules/reconciliation/standalone.js src/taskpane/modules/reconciliation/index.js
git commit -m "feat: add standalone Word OOXML ingestion exports with tests"
```
