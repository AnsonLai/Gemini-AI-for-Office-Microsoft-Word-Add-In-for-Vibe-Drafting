# Reliability, Quality Gates, and Accuracy Verification Plan (v0.8.0 Architecture)

**Date:** 2026-08-29  
**Last Updated:** 2026-09-23 (Reconciled with `@ansonlai/docx-redline-js` v0.8.0 Direct Upgrade)  
**Status:** Active — Ready for Execution  
**Prerequisite Baseline:** **The Direct Upgrade Plan (`2026-09-09-docx-redline-js-v0.5.4-upgrade.md`) is COMPLETED.** Both the root add-in and `mcp/docx-server` are pinned to exact `@ansonlai/docx-redline-js@0.8.0`.

---

## 1. Guiding Principles & Priority Hierarchy

This plan is governed by three strict priorities, ordered from most important to least important:

1. **Accuracy when making changes (Highest Priority):**  
   Accuracy is the non-negotiable bedrock of the application. Edits must be mathematically sound: target clauses must be matched with exactness, Word Accept All / Reject All lifecycles must be 100% symmetric, formatting must not be stripped or corrupted, and revision IDs must be unique document-wide. Any failure, mismatch, or ambiguous target must immediately **fail closed** and preserve the original document OOXML with zero mutations.
2. **Preserving OOXML-only approach (Second Priority):**  
   Quality gates, unit tests, error classifications, and event logging must operate purely on standard OOXML. **No Office.js dependencies in core tests or reliability layers.** Tests must run 100% offline in Node.js without requiring Microsoft Word or Office.js mocks.
3. **Speed and performance (Third Priority):**  
   Fast offline test execution ($<30\text{s}$ for the entire suite), minimal runtime telemetry overhead, and single-pass batch validation.

---

## 2. Accuracy Rules & Error Contract Governance

Adhere strictly to these validation and error handling contracts under v0.8.0:

### A. Non-Retryable Fail-Closed Errors
The following error codes emitted by `@ansonlai/docx-redline-js@0.8.0` are permanent failures for the given document state. **They must NEVER be retried automatically, and MUST NEVER trigger an unverified local fallback algorithm:**

| Error Code | Meaning | Required Consumer Action |
|---|---|---|
| `COMMENTED_CONTENT_DELETE` | Operation attempted to delete an entire paragraph that contains active reviewer comments. | **Fail Closed.** Refuse edit. Inform user/agent: *"Cannot delete paragraph with active comments. Resolve or remove comments first."* |
| `FOREIGN_PARAGRAPH_MARK_DELETION` | Operation attempted to edit inside a paragraph mark deleted by another reviewer. | **Fail Closed.** Refuse edit. Preserve original deletion markup intact. |
| `GENERATED_OOXML_INVALID` | Generated OOXML violates ECMA-376 schema constraints. | **Fail Closed.** Discard generated XML; return original OOXML completely untouched. |
| `PATCH_ROUNDTRIP_MISMATCH` | Accepted-view text verification failed (the generated redline does not reconstruct requested text upon Accept All). | **Fail Closed.** Abort edit; return original OOXML untouched. |
| `UNSAFE_REVISION_BOUNDARY` | Edit crosses unsupported structural boundaries (e.g. complex fields or content controls). | **Fail Closed.** Refuse edit. |

### B. Single-Refresh Retryable Errors
| Error Code | Meaning | Required Consumer Action |
|---|---|---|
| `TARGET_NOT_FOUND` | Target paragraph index, anchor text, or localized replacement target (`find`) could not be found in canonical document text. | Refresh the document context **at most once** and re-attempt. If still not found, abort and report to agent. |
| `AMBIGUOUS_ANCHOR` | `anchorText` appears multiple times in the paragraph and no `occurrence` was specified. | Abort edit. Request explicit `occurrence` (e.g. `occurrence: 2`) from caller. |

---

## 3. Work Packages

### WP1 — Unified Offline Test Runner (`npm test`)

**Goal:** Provide a single, deterministic test runner that verifies all Node suites offline without Microsoft Word.

**File to Create:** `scripts/run-all-tests.mjs`  
**File to Update:** `package.json`

**Step-by-Step Instructions:**
1. Create `scripts/run-all-tests.mjs`:
   ```javascript
   import { spawnSync } from 'child_process';

   const suites = [
       // 1. Result & Contract Suites
       'tests/redline_result_contract_tests.mjs',
       'tests/docx_redline_v080_compat_tests.mjs',
       'tests/no_legacy_shared_operation_bridge_tests.mjs',
       'tests/change_validation_tests.mjs',

       // 2. Pure OOXML Bridge Suites
       'tests/addin/word_operation_runner_adapter_tests.mjs',
       'tests/addin/shared_operation_bridge_tests.mjs',

       // 3. Document Visual & Structural Fidelity Suites
       'tests/ooxml_formatting_visual_tests.mjs',

       // 4. MCP Service Suites
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
     "test:compat": "node tests/docx_redline_v080_compat_tests.mjs",
     "test:visual": "node tests/ooxml_formatting_visual_tests.mjs",
     "test:perf": "node scripts/benchmark-ooxml-pipeline.mjs",
     "build": "webpack --mode production",
     "build:dev": "webpack --mode development"
   }
   ```
3. Run `npm test` and verify that all suites execute and exit with code 0.

---

### WP2 — Unified Gemini Client Implementation

**Goal:** Centralize API communication, retry logic with jitter, timeout abort control, and structured schema parsing into a single class.

**File to Create:** `src/taskpane/modules/chat/gemini-client.js`

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
                            generationConfig: options.generationConfig || {}
                        }),
                        signal: controller.signal
                    }
                );

                clearTimeout(timeoutId);

                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(`HTTP ${response.status}: ${errorText}`);
                }

                return await response.json();
            } catch (err) {
                clearTimeout(timeoutId);
                lastError = err;
                if (attempt < this.maxRetries) {
                    const delayMs = Math.pow(2, attempt) * 500 + Math.random() * 200;
                    await new Promise(resolve => setTimeout(resolve, delayMs));
                }
            }
        }

        throw new Error(`Gemini request failed after ${this.maxRetries} attempts: ${lastError?.message || lastError}`);
    }
}
```

---

### WP3 — Mutation Receipts & Telemetry Observability

Ensure every mutation preserves the compact receipt structure:
- `receipt.revisionIds`: Durable revision IDs assigned by the engine.
- `receipt.disposition`: `applied`, `no-op`, or `rolled-back`.
- `receipt.validationSummary`: Results of baseline-delta verification.
- Pass receipts to the taskpane UI and agent tools so callers understand mutation outcomes without reading raw OOXML.
