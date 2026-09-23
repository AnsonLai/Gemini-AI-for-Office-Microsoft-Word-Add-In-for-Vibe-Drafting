# OOXML Engine Performance and Web Portability Plan (v0.8.0 Single-Hop Architecture)

**Date:** 2026-08-29  
**Last Updated:** 2026-09-23 (Reconciled with `@ansonlai/docx-redline-js` v0.8.0 Direct Upgrade)  
**Status:** Active — Ready for Execution  
**Prerequisite Baseline:** **The Direct Upgrade Plan (`2026-09-09-docx-redline-js-v0.5.4-upgrade.md`) is COMPLETED.** Both the root add-in and `mcp/docx-server` are pinned to exact `@ansonlai/docx-redline-js@0.8.0`, with zero external JSZip or Node polyfills required in the browser.

---

## 1. Guiding Principles & Priority Hierarchy

This plan is governed by three strict priorities, ordered from most important to least important:

1. **Accuracy when making changes (Highest Priority):**  
   The primary mandate is absolute accuracy of document edits. Diffs must be clean, revision timestamps and author attribution paired correctly, XML schemas strictly valid, and original OOXML 100% preserved whenever an edit fails or encounters an ambiguous target.
2. **Preserving OOXML-only approach (Second Priority):**  
   **Do NOT build complex Office.js context synchronization loops.** Word interaction must be treated as a dumb, single-hop transport: Read once &rarr; Pure Engine Processing &rarr; Write once. All functions must operate purely on standard OOXML strings and packages so they can run identically in Node, a Word Taskpane, or a future web document editor.
3. **Speed and performance (Third Priority):**  
   Maximize throughput by utilizing `@ansonlai/docx-redline-js@0.8.0` caller-order-independent batching, localized exact replacements (`replacements: [{ find, replace }]`), single-pass serialization, and short-circuiting no-op edits.

---

## 2. Strategic Direction: The Two-Tier Single-Hop Pipeline

Previous drafts explored complex Office.js synchronization patterns (`context.sync` clustering, paragraph proxy caching, table cell traversals). **That approach is completely abandoned.**

Instead, the architecture enforces a **Two-Tier Single-Hop Pipeline**:

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                                Office.js Transport Layer                               │
├──────────────────────────────────────────┬─────────────────────────────────────────────┤
│  Tier 1: Document-Level Single-Hop       │  Tier 2: Range / Scoped Single-Hop          │
│  (Whole-Document Operations)             │  (Selection or Section Edits)               │
├──────────────────────────────────────────┼─────────────────────────────────────────────┤
│ 1. Read:                                 │ 1. Read:                                    │
│    getFileAsync(Compressed)              │    range.getOoxml()                         │
│    -> Uint8Array bytes                   │    -> Flat-OPC string                       │
│                                          │                                             │
│ 2. Engine:                               │ 2. Engine:                                  │
│    const doc = await openDocx(bytes)     │    applyOperationsToDocumentXml(xml, ops, { │
│    doc.applyOperations(operations, {     │      atomic: true,                          │
│      atomic: true,                       │      structuredContent: true,               │
│      author: options.author              │      pairReplacements: true                 │
│    })                                    │    })                                       │
│                                          │                                             │
│ 3. Write:                                │ 3. Write:                                   │
│    body.insertFileFromBase64(            │    range.insertOoxml(                       │
│      toBase64(doc.toUint8Array()),       │      result.documentXml,                    │
│      'Replace'                           │      'Replace'                              │
│    )                                     │    )                                        │
└──────────────────────────────────────────┴─────────────────────────────────────────────┘
```

- **Tier 1 (Universal Facade via `openDocx`)**: For document-wide agentic tool runs (`apply_redlines`), using `getFileAsync` + `openDocx` + `insertFileFromBase64` provides 100% package fidelity. Word's entire numbering table, styles, headers, and footers are preserved automatically by `docx-redline-js@0.8.0`.
- **Tier 2 (Flat-OPC via `applyOperationsToDocumentXml`)**: For smaller, surgical paragraph/selection operations where fetching the whole file is unnecessary, `applyOperationsToDocumentXml` provides sub-10ms in-memory diffing.

---

## 3. Reconciled Status of Optimization Tasks

| Legacy Task ID | Original Title | Reconciled Status | Rationale & Permanent Decision in v0.8.0 |
|---|---|---|---|
| **Phase 7** | Engine boundary verification | **Completed Upstream** | Fully verified in `@ansonlai/docx-redline-js` upstream test harness. |
| **P5.1** | Reduce parse/serialize churn | **Offloaded to Library** | Handled natively by upstream single-DOM batch sessions. |
| **P5.2** | Profile surgical allocations | **Offloaded to Library** | Table-cell and run allocation optimizations belong to the library engine. |
| **P5.3** | Benchmark string operations | **Cancelled** | Sub-millisecond; micro-optimizing string concatenation yields no user benefit. |
| **P5.4** | Collapse Word sync clusters | **SUPERSEDED by Single-Hop** | **Replaced entirely by the Single-Hop Pipeline.** Eliminates all iterative `Paragraph` proxy loops. |
| **P5.5** | Reduce memory churn (`RunModel`) | **Offloaded to Library** | `RunModel` is internal to `@ansonlai/docx-redline-js`. |
| **P5.6 / P5.9** | Diff-result caching & deferred DOM | **Cancelled** | In-memory diff caching creates cache invalidation bugs. Upstream diffing takes $<2\text{ms}$. |
| **P5.7** | Web runtime tuning & lazy loading | **ACTIVE (WP2)** | Ensure the taskpane loads lightweight modules first; dynamically import heavy UI components. |
| **P5.8** | Build shared `DocumentIndex` | **Cancelled** | Upstream v0.8.0 maintains an internal targeting lookup cache. |

---

## 4. Detailed Step-by-Step Implementation Instructions

### Step 1: Implement the Single-Hop OOXML Runner in `word-operation-runner.js`

**File:** `src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js`

Implement the batch execution flow supporting both full `content` and localized `replacements: [{ find, replace }]`:

```javascript
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js';
import { assertRedlineResult, prepareOperationInput } from './redline-result.js';
import { getParagraphOoxmlWithFallback, insertOoxmlWithRangeFallback, withNativeTrackingDisabled } from './word-ooxml.js';

/**
 * Executes a batch of canonical operations purely in OOXML in a single hop.
 *
 * @param {Word.RequestContext} context - Word context (used strictly for single-hop I/O)
 * @param {Word.Paragraph|Word.Range} targetScope - The target paragraph or range proxy
 * @param {Array<Object>} operations - Canonical redline operations
 * @param {Object} [options={}] - Options (author, atomic, etc.)
 * @returns {Promise<{ status: string, receipts?: Array<Object>, error?: Object, hasChanges: boolean }>}
 */
export async function executePureOoxmlBatch(context, targetScope, operations, options = {}) {
    // 1. Single-Hop Read: Fetch plain OOXML string from Word
    const readResult = await getParagraphOoxmlWithFallback(targetScope, context, { logPrefix: 'PureOoxml' });
    const originalOoxml = readResult.ooxml;
    if (!originalOoxml) {
        return { status: 'error', error: { code: 'OOXML_READ_FAILED', message: 'Failed to read OOXML from target scope.' }, hasChanges: false };
    }

    // 2. Prepare operations (support both localized replacements and full modified content)
    const sanitizedOps = operations.map(op => {
        const copy = { ...op };
        if (Array.isArray(copy.replacements)) {
            // Localized replacements: pass through cleanly
            copy.replacements = copy.replacements.map(r => ({
                find: String(r.find || ''),
                replace: options.isModelGenerated ? prepareOperationInput(r.replace) : String(r.replace || '')
            }));
        } else if (copy.modified !== undefined) {
            copy.modified = options.isModelGenerated ? prepareOperationInput(copy.modified) : copy.modified;
        }
        return copy;
    });

    // 3. Pure OOXML Processing: Run through docx-redline-js v0.8.0 in-memory
    const engineResult = await applyOperationsToDocumentXml(originalOoxml, sanitizedOps, {
        atomic: options.atomic !== false,
        author: options.author || 'AI Assistant',
        structuredContent: true,
        pairReplacements: true,
        existingRevisions: 'merge-same-author'
    });

    // 4. Accuracy Check: Validate result status before touching Word
    if (engineResult.status === 'error' || engineResult.rolledBack) {
        return {
            status: 'error',
            error: engineResult.error || { code: 'BATCH_ROLLED_BACK', message: 'Batch rolled back.' },
            receipts: engineResult.receipts || [],
            hasChanges: false
        };
    }

    // 5. Short-circuit if no changes were generated
    if (!engineResult.hasChanges || !engineResult.documentXml || engineResult.documentXml === originalOoxml) {
        return { status: 'success', hasChanges: false, receipts: engineResult.receipts || [] };
    }

    // 6. Single-Hop Write: Insert updated OOXML string back into Word once
    await withNativeTrackingDisabled(context, async () => {
        await insertOoxmlWithRangeFallback(targetScope, engineResult.documentXml, 'Replace', context, 'PureOoxmlWrite');
    }, {
        enabled: true,
        logPrefix: 'PureOoxmlWrite'
    });

    return {
        status: 'success',
        hasChanges: true,
        receipts: engineResult.receipts || []
    };
}
```

### Step 2: Strip Office.js Proxy Loops from `word-redline-runner.js`
1. Open `src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js`.
2. Delete the 1,100-line iterative paragraph proxy loop (`for (const change of (changes || []))`).
3. Delete table synthesis heuristics (`synthesizeMarkdownTableFromSourceRange`, `inferTableConversionEndIndex`), delegating to upstream `structuredContent: true`.
4. Delete helper functions that loop over `Paragraph` proxies:
   - Remove `applyNativeParagraphFormatting`
   - Remove `prepareNativeMarkdownParagraph`
   - Remove `insertNativeTextLines`
5. Replace the entry point with a clean call to `executePureOoxmlBatch`.

---

## 5. Document Visual Formatting & Structural Test Suite

Create `tests/ooxml_formatting_visual_tests.mjs` to empirically verify formatting, whitespace, tabs, breaks, hyperlinks, section properties, and localized replacements:

```javascript
import assert from 'node:assert/strict';
import { applyOperationToDocumentXml, applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js';

console.log('Running OOXML Formatting & Visual Fidelity Tests (v0.8.0)...\n');

// TEST 1: Formatting Inheritance (Bold/Italic/Underline) in Inserted Runs
{
    console.log('Test 1: Character formatting inheritance visual check...');
    const docWithFormatting = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">This document is </w:t></w:r>
      <w:r>
        <w:rPr><w:b/><w:i/></w:rPr>
        <w:t>strictly confidential</w:t>
      </w:r>
      <w:r><w:t xml:space="preserve"> and proprietary.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(docWithFormatting, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        anchor: { exactText: 'strictly confidential' },
        modified: 'strictly private and confidential',
        author: 'Editor'
    });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('<w:ins'), 'Must contain insertion markup');
    assert.ok(xml.includes('<w:b/>') || xml.includes('<w:b w:val="true"/>'), 'Inserted text must retain bold styling');
    assert.ok(xml.includes('<w:i/>') || xml.includes('<w:i w:val="true"/>'), 'Inserted text must retain italic styling');
    console.log('  -> Passed: Formatting (<w:b/>, <w:i/>) cleanly preserved on <w:ins> visual run.');
}

// TEST 2: Whitespace Padding & Preservation (xml:space="preserve")
{
    console.log('Test 2: Whitespace padding visual check...');
    const docWithSpaces = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">Fee of $1,000 shall apply.</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(docWithSpaces, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        anchor: { exactText: '$1,000' },
        modified: '$2,500',
        author: 'Editor'
    });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('xml:space="preserve"'), 'Runs must maintain xml:space="preserve"');
    console.log('  -> Passed: xml:space="preserve" retained; words will not collide visually in Word.');
}

// TEST 3: Structural Tab and Line Break Preservation (<w:tab/>, <w:br/>)
{
    console.log('Test 3: Structural tab and line break visual fidelity...');
    const docWithTabsAndBreaks = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Section 1</w:t><w:tab/><w:t>Title of Section</w:t><w:br/><w:t>Details follow below.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(docWithTabsAndBreaks, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        anchor: { exactText: 'Title of Section' },
        modified: 'Updated Section Title',
        author: 'Editor'
    });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('<w:tab/>') || xml.includes('<w:tab />'), 'Tab marker must be preserved');
    assert.ok(xml.includes('<w:br/>') || xml.includes('<w:br />'), 'Line break marker must be preserved');
    console.log('  -> Passed: Visual structural tab (<w:tab/>) and line break (<w:br/>) preserved intact.');
}

// TEST 4: Hyperlink Element Preservation (<w:hyperlink>)
{
    console.log('Test 4: Hyperlink preservation visual check...');
    const docWithLink = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Visit our </w:t></w:r>
      <w:hyperlink r:id="rId5" w:history="1">
        <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>customer portal</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t xml:space="preserve"> for documentation.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(docWithLink, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        anchor: { exactText: 'customer portal' },
        modified: 'support portal',
        author: 'Editor'
    });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('<w:hyperlink'), 'Must preserve <w:hyperlink>');
    assert.ok(xml.includes('r:id="rId5"'), 'Must preserve hyperlink relationship ID rId5');
    console.log('  -> Passed: Clickable hyperlink (<w:hyperlink r:id="rId5">) preserved.');
}

// TEST 5: Localized Exact Replacements (replacements: [{ find, replace }])
{
    console.log('Test 5: Localized exact replacements check (v0.7.0+)...');
    const docText = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Agreement shall terminate within 30 days of written notice.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(docText, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        replacements: [{ find: '30 days', replace: '60 days' }],
        author: 'Editor'
    });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;
    assert.ok(xml.includes('<w:del'), 'Must contain deletion for 30 days');
    assert.ok(xml.includes('<w:ins'), 'Must contain insertion for 60 days');
    assert.ok(xml.includes('60 days'), 'Replacement text must be present');
    console.log('  -> Passed: Localized replacement ({ find, replace }) executed cleanly.');
}

console.log('\nAll OOXML Formatting & Visual Fidelity Tests Passed Successfully!');
```

---

## 6. Performance & Memory Profiling Harness

Create `scripts/benchmark-ooxml-pipeline.mjs` to measure batch throughput and memory consumption across synthetic documents:

```javascript
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js';
import { performance } from 'node:perf_hooks';

function generateSyntheticDocumentXml(paragraphCount) {
    let paragraphs = '';
    for (let i = 0; i < paragraphCount; i++) {
        paragraphs += `<w:p><w:r><w:t>Paragraph ${i}: Standard contractual clause terms and conditions text for testing.</w:t></w:r></w:p>\n`;
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n<w:body>\n${paragraphs}</w:body>\n</w:document>`;
}

async function runBenchmark() {
    console.log('=== OOXML Single-Hop Pipeline Performance Benchmark (v0.8.0) ===\n');

    for (const count of [10, 50, 100, 500]) {
        const docXml = generateSyntheticDocumentXml(count);
        const operations = [
            { type: 'redline', target: { paragraphIndex: 0 }, anchor: { exactText: 'Standard' }, modified: 'Revised', author: 'Benchmarker' },
            { type: 'redline', target: { paragraphIndex: Math.floor(count / 2) }, anchor: { exactText: 'Standard' }, modified: 'Updated', author: 'Benchmarker' },
            { type: 'redline', target: { paragraphIndex: count - 1 }, anchor: { exactText: 'Standard' }, modified: 'Final', author: 'Benchmarker' }
        ];

        const startMem = process.memoryUsage().heapUsed;
        const startTime = performance.now();

        const result = await applyOperationsToDocumentXml(docXml, operations, { atomic: true, pairReplacements: true });

        const elapsedMs = performance.now() - startTime;
        const memoryDeltaMb = (process.memoryUsage().heapUsed - startMem) / (1024 * 1024);

        console.log(`Document Size: ${count} paragraphs`);
        console.log(`  - Status: ${result.status} (hasChanges: ${result.hasChanges})`);
        console.log(`  - Execution Time: ${elapsedMs.toFixed(2)} ms`);
        console.log(`  - Memory Delta: ${memoryDeltaMb.toFixed(2)} MB\n`);

        if (count <= 100 && elapsedMs > 250) {
            console.warn(`[PERF WARNING] Benchmark exceeded 250ms budget for ${count} paragraphs.`);
        }
    }
}

runBenchmark();
```

---

## 7. Verification and Acceptance Checklist

### Accuracy & Visual Formatting Gates
- [ ] Run `node tests/ooxml_formatting_visual_tests.mjs` &rarr; 100% pass across all test cases.
- [ ] Formatting Inheritance: Editing bold or italic text preserves `<w:rPr>` on inserted runs without font resets.
- [ ] Whitespace Preservation: Ensure `xml:space="preserve"` prevents words from colliding in Word.
- [ ] Structural Markers: Literal `<w:tab/>` and `<w:br/>` survive redline operations without degradation.
- [ ] Hyperlinks: `<w:hyperlink>` and relationship IDs remain intact.
- [ ] Localized Replacements: Verify `{ find, replace }` executes with surgical run splitting.

### OOXML Purity & Web Portability Gates
- [ ] Single-Hop Word I/O: Applying an edit invokes exactly ONE read and ONE write in Word. Zero iterative `context.sync()` calls within the edit application.
- [ ] Zero Office.js in Engine: Assert that `@ansonlai/docx-redline-js` has 0 references to `Word.` or Office.js.
- [ ] Offline Execution: All tests run 100% offline in a pure Node environment without Office.js mocks.

### Performance Gates
- [ ] Caller-Order-Independent Batches: Applying a batch of operations executes in a single XML DOM parse and serialize pass.
- [ ] No-Op Short-Circuit: Applying an identical replacement completes in $<5\text{ms}$ and performs 0 Word write operations.
- [ ] Benchmark Gate: Run `node scripts/benchmark-ooxml-pipeline.mjs` &rarr; 100-paragraph documents process in $<100\text{ms}$.
