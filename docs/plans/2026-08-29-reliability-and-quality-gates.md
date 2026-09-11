# Reliability, Quality Gates, and Accuracy Verification Plan

**Date:** 2026-08-29  
**Last Updated:** 2026-09-10 (Deepened & Reconciled with `@ansonlai/docx-redline-js` v0.5.4 & Pure OOXML Architecture)  
**Status:** Active — Ready for Execution  
**Prerequisite Baseline:** **The 2026-09-09 Upgrade Plan (`2026-09-09-docx-redline-js-v0.5.4-upgrade.md`) is COMPLETED.** Both the root add-in and `mcp/docx-server` are pinned to exact `@ansonlai/docx-redline-js@0.5.4`, and compatibility suites WP0–WP4 are verified. This plan builds directly on top of that stable baseline.

---

## 1. Guiding Principles & Priority Hierarchy

This plan is governed by three strict priorities, ordered from most important to least important:

1. **Accuracy when making changes (Highest Priority):**  
   Accuracy is the non-negotiable bedrock of the application. Edits must be mathematically sound: target clauses must be matched with exactness, Word Accept All / Reject All lifecycles must be 100% symmetric, formatting must not be stripped or corrupted, and revision IDs must be unique document-wide. Any failure, mismatch, or ambiguous target must immediately **fail closed** and preserve the original document OOXML with zero mutations.
2. **Preserving OOXML-only approach (Second Priority):**  
   Quality gates, unit tests, error classifications, and event logging must operate purely on standard OOXML. **No Office.js dependencies in core tests or reliability layers.** Tests must run 100% offline in Node.js without requiring Microsoft Word or Office.js mocks. All functions must be verified to be completely portable to a web-based document editor, allowing Microsoft Word to be discarded in the future.
3. **Speed and performance (Third Priority):**  
   Fast offline test execution ($<30\text{s}$ for the entire suite), minimal runtime telemetry overhead (lightweight in-memory ring buffer), and single-pass batch validation.

---

## 2. Accuracy Rules & Error Contract Governance

A smaller model implementing or testing changes must adhere strictly to these validation and error handling contracts:

### A. Non-Retryable Fail-Closed Errors
The following error codes emitted by `@ansonlai/docx-redline-js` are permanent failures for the given document state. **They must NEVER be retried automatically, and MUST NEVER trigger an unverified local fallback algorithm:**

| Error Code | Meaning | Required Consumer Action |
|---|---|---|
| `COMMENTED_CONTENT_DELETE` | Operation attempted to delete an entire paragraph that contains active reviewer comments. | **Fail Closed.** Refuse edit. Inform user/agent: *"Cannot delete paragraph with active comments. Resolve or remove comments first."* |
| `FOREIGN_PARAGRAPH_MARK_DELETION` | Operation attempted to edit inside a paragraph mark deleted by another reviewer. | **Fail Closed.** Refuse edit. Preserve original deletion markup intact. |
| `GENERATED_OOXML_INVALID` | Generated OOXML violates ECMA-376 schema or schema constraints. | **Fail Closed.** Discard generated XML; return original OOXML completely untouched. |
| `PATCH_ROUNDTRIP_MISMATCH` | Accepted-view text verification failed (the generated redline does not reconstruct the requested text upon Accept All). | **Fail Closed.** Abort edit; return original OOXML untouched. |
| `UNSAFE_REVISION_BOUNDARY` | Edit crosses unsupported structural boundaries (e.g. complex fields, content controls, or table rows). | **Fail Closed.** Refuse edit. |

### B. Single-Refresh Retryable Errors
| Error Code | Meaning | Required Consumer Action |
|---|---|---|
| `TARGET_NOT_FOUND` | Target paragraph index or anchor text could not be found in the current canonical document text. | Refresh the document context **at most once** and re-attempt. If still not found, abort and report to agent. |
| `AMBIGUOUS_ANCHOR` | `anchorText` appears multiple times in the paragraph and no `occurrence` was specified. | Abort edit. Request explicit `occurrence` (e.g. `occurrence: 2`) from caller. |

---

## 3. Work Packages

### WP1 — Unified Offline Test Runner (`npm test`)

**Goal:** Provide a single, deterministic test runner that verifies all Node suites offline without Microsoft Word.

**File to Create:** `scripts/run-all-tests.mjs`  
**File to Update:** `package.json`

**Step-by-Step Instructions for a Smaller Model:**
1. Create `scripts/run-all-tests.mjs`:
   ```javascript
   import { spawnSync } from 'child_process';

   const suites = [
       // 1. Result & Contract Suites
       'tests/redline_result_contract_tests.mjs',
       'tests/docx_redline_v040_compat_tests.mjs',
       'tests/docx_redline_v054_compat_tests.mjs',
       'tests/no_legacy_shared_operation_bridge_tests.mjs',

       // 2. Pure OOXML Bridge Suites
       'tests/addin/word_operation_runner_adapter_tests.mjs',
       'tests/addin/shared_operation_bridge_tests.mjs',

       // 3. Document Visual & Structural Fidelity Suites
       'tests/document_visual_structure_tests.mjs',
       'tests/ooxml_formatting_visual_tests.mjs',
       'tests/comment_threading_visual_tests.mjs',
       'tests/package_opc_boundary_tests.mjs',
       'tests/document_integrity_safety_tests.mjs',

       // 4. Web Portability Suite
       'tests/web_portability_tests.mjs',

       // 5. Offline Model Replay Evals
       'tests/evals/offline_replay_evals.mjs',

       // 6. MCP Service Suites
       'tests/mcp_docx_redline_service_tests.mjs'
   ];

   console.log(`Starting offline verification of ${suites.length} test suites...\n`);

   for (const suite of suites) {
       console.log(`Running: ${suite}`);
       const result = spawnSync(process.execPath, [suite], { stdio: 'inherit' });
       if (result.status !== 0) {
           console.error(`\nFAILED: ${suite}`);
           process.exit(result.status || 1);
       }
   }

   console.log('\nAll offline test suites passed successfully!');
   ```
2. In `package.json`, update `"scripts"`:
   ```json
   "scripts": {
     "test": "node scripts/run-all-tests.mjs",
     "test:visual": "node tests/document_visual_structure_tests.mjs",
     "test:compat": "node tests/docx_redline_v054_compat_tests.mjs",
     "test:perf": "node scripts/benchmark-ooxml-pipeline.mjs",
     "lint": "eslint src/ tests/",
     "build": "webpack --mode production",
     "build:dev": "webpack --mode development"
   }
   ```
3. Run `npm test` and verify that all suites execute and exit with code 0.

---

### WP2 — Unified Gemini Client Implementation

**Goal:** Centralize API communication, retry logic with jitter, timeout abort control, and structured schema parsing into a single class.

**File to Create:** `src/taskpane/modules/chat/gemini-client.js`

**Step-by-Step Instructions for a Smaller Model:**
```javascript
/**
 * Unified, resilient Google Gemini API client.
 */
export class GeminiClient {
    constructor(config = {}) {
        this.apiKey = config.apiKey || '';
        this.model = config.model || 'gemini-1.5-pro';
        this.timeoutMs = config.timeoutMs || 30000;
        this.maxRetries = config.maxRetries || 3;
    }

    async generateContent(messages, options = {}) {
        let attempt = 0;
        let lastError = null;

        while (attempt < this.maxRetries) {
            attempt++;
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), this.timeoutMs);

            try {
                const response = await fetch(
                    `https://generativelanguage.googleapis.com/v1beta/models/${this.model}:generateContent?key=${this.apiKey}`,
                    {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            contents: messages,
                            generationConfig: {
                                responseMimeType: options.responseMimeType || 'application/json',
                                temperature: options.temperature ?? 0.1
                            }
                        }),
                        signal: controller.signal
                    }
                );

                clearTimeout(timeoutId);

                if (!response.ok) {
                    const status = response.status;
                    if (status === 429 || status === 503) {
                        // Exponential backoff with jitter
                        const delay = Math.pow(2, attempt) * 1000 + Math.random() * 500;
                        console.warn(`[GeminiClient] HTTP ${status}. Retrying in ${delay.toFixed(0)}ms (attempt ${attempt}/${this.maxRetries})...`);
                        await new Promise(r => setTimeout(r, delay));
                        continue;
                    }
                    const errorText = await response.text();
                    throw new Error(`Gemini API Error (HTTP ${status}): ${errorText}`);
                }

                const json = await response.json();
                return json;
            } catch (err) {
                clearTimeout(timeoutId);
                lastError = err;
                if (err.name === 'AbortError') {
                    throw new Error(`Gemini API request timed out after ${this.timeoutMs}ms.`);
                }
                if (attempt >= this.maxRetries) break;
            }
        }

        throw lastError || new Error('Failed to generate content after max retries.');
    }
}
```

---

### WP3 — Offline Model Replay Evaluation Suite

To guard against **prompt drift, model thought leakage, markdown fence wrapping, and truncated JSON**, create `tests/evals/offline_replay_evals.mjs`.

**File to Create:** `tests/evals/offline_replay_evals.mjs`

```javascript
import assert from 'node:assert/strict';
import { mapChangeToCanonicalOperation } from '../../src/taskpane/modules/commands/agentic-tools.js';

console.log('Running Offline Model Replay Evaluations...\n');

// HELPER: Clean raw LLM response text
function cleanModelResponseText(rawText) {
    let text = String(rawText || '');
    // 1. Strip thought blocks (<thought>...</thought>)
    text = text.replace(/<thought>[\s\S]*?<\/thought>/gi, '').trim();
    // 2. Strip markdown json code fences (```json ... ```)
    const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
    if (fenceMatch) {
        text = fenceMatch[1].trim();
    }
    return text;
}

// TEST 1: Thought Leakage Stripping
{
    console.log('Test 1: Thought leakage defense...');
    const rawOutputWithThoughts = `<thought>
I need to replace December 31 with November 30 in paragraph 2.
Let me formulate the JSON output.
</thought>
[
  {
    "action": "modify_text",
    "paragraphIndex": 2,
    "anchorText": "December 31",
    "text": "November 30"
  }
]`;

    const cleaned = cleanModelResponseText(rawOutputWithThoughts);
    assert.ok(!cleaned.includes('<thought>'), 'Must strip thought tags');
    const parsed = JSON.parse(cleaned);
    assert.equal(parsed.length, 1);
    assert.equal(parsed[0].text, 'November 30');

    const op = mapChangeToCanonicalOperation(parsed[0], 'AI Reviewer');
    assert.equal(op.type, 'redline');
    assert.equal(op.modified, 'November 30');
    console.log('  -> Passed: Thought block successfully isolated from structured JSON payload.');
}

// TEST 2: Markdown Fence Wrapping Handling
{
    console.log('Test 2: Markdown code fence parsing...');
    const rawFencedOutput = "```json\n[\n  {\n    \"action\": \"delete\",\n    \"paragraphIndex\": 5\n  }\n]\n```";
    const cleaned = cleanModelResponseText(rawFencedOutput);
    const parsed = JSON.parse(cleaned);
    assert.equal(parsed[0].action, 'delete');
    const op = mapChangeToCanonicalOperation(parsed[0], 'AI Reviewer');
    assert.equal(op.type, 'delete');
    assert.equal(op.target.paragraphIndex, 5);
    console.log('  -> Passed: Markdown code fences stripped cleanly.');
}

console.log('\nAll Offline Model Replay Evaluations Passed Successfully!');
```

---

### WP4 — Document Integrity & Safety Test Suite

To verify **how document safety, non-mutating failures, and whole-document ID uniqueness show up in the document**, create `tests/document_integrity_safety_tests.mjs`.

**File to Create:** `tests/document_integrity_safety_tests.mjs`

```javascript
import assert from 'node:assert/strict';
import { applyOperationToDocumentXml, applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js';

console.log('Running Document Integrity & Safety Tests...\n');

const BASELINE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>First paragraph: Initial content.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second paragraph: Untouched content.</w:t></w:r></w:p>
  </w:body>
</w:document>`;

// TEST 1: Non-Mutating Failure Invariant (Bit-for-Bit Preservation on Failure)
{
    console.log('Test 1: Non-mutating failure invariant...');
    const failureResult = await applyOperationToDocumentXml(BASELINE_XML, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        anchor: { exactText: 'NonExistentStringThatCannotBeFound' },
        modified: 'Replaced String',
        author: 'Tester'
    });

    assert.equal(failureResult.status, 'error');
    assert.equal(failureResult.error.code, 'TARGET_NOT_FOUND');

    assert.equal(failureResult.documentXml, BASELINE_XML, 'On error, original OOXML must be returned completely unmodified');
    console.log('  -> Passed: Original OOXML preserved bit-for-bit on failure.');
}

// TEST 2: Whole-Document Revision ID Uniqueness
{
    console.log('Test 2: Whole-document unique revision ID check...');
    const batchResult = await applyOperationsToDocumentXml(BASELINE_XML, [
        {
            type: 'redline',
            target: { paragraphIndex: 0 },
            anchor: { exactText: 'Initial' },
            modified: 'Primary',
            author: 'Tester'
        },
        {
            type: 'redline',
            target: { paragraphIndex: 1 },
            anchor: { exactText: 'Untouched' },
            modified: 'Modified',
            author: 'Tester'
        }
    ], { atomic: true });

    assert.equal(batchResult.status, 'success');
    const xml = batchResult.documentXml;

    const idMatches = [...xml.matchAll(/<(?:w:ins|w:del)[^>]*w:id="([^"]+)"/g)].map(m => m[1]);
    assert.ok(idMatches.length >= 4, 'Must have at least 4 revision tags');

    const uniqueIds = new Set(idMatches);
    assert.equal(uniqueIds.size, idMatches.length, 'Every revision w:id in the document must be strictly unique');
    console.log(`  -> Passed: All ${idMatches.length} generated revision IDs are globally unique.`);
}

// TEST 3: Commented Paragraph Deletion Safeguard (COMMENTED_CONTENT_DELETE)
{
    console.log('Test 3: Commented paragraph deletion protection...');
    const docWithComment = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>Important clause under active review.</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const deleteResult = await applyOperationToDocumentXml(docWithComment, {
        type: 'delete',
        target: { paragraphIndex: 0 },
        author: 'Tester'
    });

    assert.equal(deleteResult.status, 'error');
    assert.equal(deleteResult.error.code, 'COMMENTED_CONTENT_DELETE');
    assert.equal(deleteResult.documentXml, docWithComment, 'Document with active comments must not be deleted');
    console.log('  -> Passed: Paragraph deletion blocked to protect active reviewer comments.');
}

console.log('\nAll Document Integrity & Safety Tests Passed Successfully!');
```

---

### WP5 — Continuous Integration Pipeline Definition

**File to Create:** `.github/workflows/ci.yml`

```yaml
name: Continuous Integration

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  verify:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout Code
        uses: actions/checkout@v4

      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: 20
          cache: 'npm'

      - name: Install Root Dependencies
        run: npm ci

      - name: Install MCP Dependencies
        run: npm --prefix mcp/docx-server ci

      - name: Verify Exact Upstream Version Pin
        run: |
          npm ls @ansonlai/docx-redline-js | grep "0.5.4"
          npm --prefix mcp/docx-server ls @ansonlai/docx-redline-js | grep "0.5.4"

      - name: Run ESLint
        run: npm run lint

      - name: Run Full Offline Test Suites
        run: npm test

      - name: Run Webpack Build
        run: npm run build
```

---

## 4. Master Test Directory Inventory

| Test File Path | Category | Core Assertions & Purpose |
|---|---|---|
| `tests/document_visual_structure_tests.mjs` | **Visual Structure** | Reviewing pane card grouping, multi-level list properties, markdown table grids, and Accept/Reject symmetry. |
| `tests/ooxml_formatting_visual_tests.mjs` | **Visual Formatting** | Formatting inheritance (`<w:b/>`, `<w:i/>`), whitespace preservation (`xml:space`), tabs, breaks, hyperlinks, and section breaks. |
| `tests/comment_threading_visual_tests.mjs` | **Visual Comments** | Comment range wrapping around targeted text; comment replies threaded in `commentsExtended.xml`. |
| `tests/package_opc_boundary_tests.mjs` | **OPC Package** | Content types, relationship mapping, numbering definitions in complete `.docx` packages. |
| `tests/document_integrity_safety_tests.mjs` | **Safety & Invariants** | Bit-for-bit non-mutating failure invariant, global revision ID uniqueness, and commented paragraph deletion protection. |
| `tests/web_portability_tests.mjs` | **Portability** | Confirms 100% of core operations execute in Node with `global.Word = undefined` and `global.Office = undefined`. |
| `tests/evals/offline_replay_evals.mjs` | **Model Boundary** | Thought leakage defense, markdown code fence stripping, and truncated JSON recovery. |
| `tests/docx_redline_v054_compat_tests.mjs` | **Upstream Contract** | Validates `@ansonlai/docx-redline-js@0.5.4` compatibility contracts. |
| `tests/mcp_docx_redline_service_tests.mjs` | **MCP Server** | Verifies file-level paragraph edits, comment replies, and atomic rollback in `mcp/docx-server`. |

---

## 5. Verification and Acceptance Checklist

- [ ] All 10 suites in the Master Test Inventory pass 100% green via `npm test`.
- [ ] Offline runner execution time is $<30\text{s}$.
- [ ] GitHub Actions CI workflow runs cleanly with zero dependency mismatches.
- [ ] Thought leakage and markdown fences are proven stripped before operation mapping.
- [ ] Non-mutating failure invariant is mathematically proven across all test suites.
