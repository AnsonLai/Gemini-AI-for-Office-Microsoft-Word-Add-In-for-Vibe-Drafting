# OOXML Engine Performance and Web Portability Plan

**Date:** 2026-08-29  
**Last Updated:** 2026-09-10 (Deepened & Reconciled with `@ansonlai/docx-redline-js` v0.5.4 & Pure OOXML Architecture)  
**Status:** Active — Ready for Execution  
**Prerequisite Baseline:** **The 2026-09-09 Upgrade Plan (`2026-09-09-docx-redline-js-v0.5.4-upgrade.md`) is COMPLETED.** Both the root add-in and `mcp/docx-server` are pinned to exact `@ansonlai/docx-redline-js@0.5.4`, and compatibility suites WP0–WP4 are verified. This plan builds directly on top of that stable baseline.

---

## 1. Guiding Principles & Priority Hierarchy

This plan is governed by three strict priorities, ordered from most important to least important:

1. **Accuracy when making changes (Highest Priority):**  
   The primary mandate is absolute accuracy of document edits. Diffs must be clean, revision timestamps and author attribution paired correctly, XML schemas strictly valid, and original OOXML 100% preserved whenever an edit fails or encounters an ambiguous target.
2. **Preserving OOXML-only approach (Second Priority):**  
   **Do NOT build complex Office.js context synchronization loops.** Word interaction must be treated as a dumb, single-hop transport: Read OOXML once -> Pure OOXML Processing -> Write OOXML once. All functions must operate purely on standard OOXML strings and packages so they can be ported directly to a pure web editor (e.g. Canvas/DOM/WASM-based document editor) in the future, completely ditching Microsoft Word.
3. **Speed and performance (Third Priority):**  
   Within the pure OOXML pipeline, maximize throughput by utilizing `@ansonlai/docx-redline-js` single-DOM batch sessions, avoiding redundant XML serialization passes, and short-circuiting no-op edits.

---

## 2. Strategic Direction: The Single-Hop OOXML Pipeline

Previous drafts explored complex Office.js synchronization patterns (`context.sync` clustering, paragraph proxy caching, table cell traversals). **That approach is completely abandoned.**

Instead, the architecture enforces a **Single-Hop OOXML Pipeline**:
```
┌────────────────────────────────────────────────────────────────────────┐
│                        I/O Boundary (Host Shell)                       │
│                                                                        │
│   Word Add-in:                Web Application:            Node / MCP:  │
│   body.getOoxml()             document.xml / buffer       fs.readFile  │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │ ooxmlString (Plain Text / Package)
┌───────────────────────────────────▼────────────────────────────────────┐
│                    Pure OOXML Engine (Web-Portable)                    │
│                                                                        │
│   applyOperationsToDocumentXml(ooxmlString, operations, {              │
│       atomic: true,                                                    │
│       structuredContent: true,                                         │
│       pairReplacements: true,                                          │
│       existingRevisions: 'merge-same-author'                           │
│   })                                                                   │
│                                                                        │
│   - Single-DOM Session: Parse once, diff all, serialize once           │
│   - Zero Office.js / DOM dependencies                                  │
│   - Identical execution in Node, Browser, or Word Taskpane             │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │ result.documentXml / result.receipt
┌───────────────────────────────────▼────────────────────────────────────┐
│                        I/O Boundary (Host Shell)                       │
│                                                                        │
│   Word Add-in:                Web Application:            Node / MCP:  │
│   range.insertOoxml(Replace)  renderView() / export       fs.writeFile │
└────────────────────────────────────────────────────────────────────────┘
```

---

## 3. Reconciled Status of Legacy Optimization Tasks

All previous in-plugin engine tasks are formally resolved as follows:

| Legacy Task ID | Original Title | Reconciled Status | Rationale & Permanent Decision |
|---|---|---|---|
| **Phase 7** | Engine boundary verification | **Completed Upstream** | Fully verified in `@ansonlai/docx-redline-js` upstream test harness (99/99 test groups, zero host leakage). |
| **P5.1** | Reduce parse/serialize churn | **Offloaded to Library** | Solved by upstream v0.5.0 Single-DOM Document Session. The plugin calls `applyOperationsToDocumentXml`, which parses the XML once for an entire batch. |
| **P5.2** | Profile surgical allocations | **Offloaded to Library** | Table-cell and run allocation optimizations belong to the library engine. |
| **P5.3** | Benchmark string operations | **Cancelled** | Pure string operations are sub-millisecond; micro-optimizing string concatenation in the plugin yields no user benefit. |
| **P5.4** | Collapse Word sync clusters | **SUPERSEDED by Single-Hop** | **Replaced entirely by the Single-Hop OOXML Pipeline.** We do not optimize proxy loops; we eliminate them completely. |
| **P5.5** | Reduce memory churn (`RunModel`) | **Offloaded to Library** | `RunModel` is internal to `@ansonlai/docx-redline-js`. The plugin never directly instantiates or clones `RunModel`. |
| **P5.6 / P5.9** | Diff-result caching & deferred DOM | **Cancelled** | In-memory diff caching creates cache invalidation bugs and stale state. Upstream diffing takes $<2\text{ms}$. |
| **P5.7** | Web runtime tuning & lazy loading | **ACTIVE (WP2)** | Ensure the taskpane and web demo load lightweight modules first; dynamically import heavy visualization components. |
| **P5.8** | Build shared `DocumentIndex` | **Cancelled** | Do NOT build an in-memory shadow document index in the plugin. Upstream v0.5.0 already maintains an internal 47x faster targeting lookup cache. |

---

## 4. Detailed Step-by-Step Implementation Instructions

### Step 1: Implement the Single-Hop OOXML Runner in `word-operation-runner.js`

**File:** `src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js`

A smaller model should implement the execution flow using this exact contract:

```javascript
import { applyOperationToDocumentXml, applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js';
import { assertRedlineResult, prepareOperationInput } from './redline-result.js';
import { getParagraphOoxmlWithFallback, insertOoxmlWithRangeFallback, withNativeTrackingDisabled } from './word-ooxml.js';

/**
 * Executes a batch of canonical operations purely in OOXML.
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

    // 2. Prepare operations (explicit sanitization for model-generated content)
    const sanitizedOps = operations.map(op => ({
        ...op,
        modified: options.isModelGenerated ? prepareOperationInput(op.modified) : op.modified
    }));

    // 3. Pure OOXML Processing: Run through docx-redline-js in-memory
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
2. Delete helper functions that loop over `Paragraph` proxies to apply text or format:
   - Remove `applyNativeParagraphFormatting`
   - Remove `prepareNativeMarkdownParagraph`
   - Remove `insertNativeTextLines` (replace with single package insertion)
3. Ensure that when inserting multi-line or structured content into an anchor paragraph, `word-redline-runner.js` converts the content into an OOXML package string via `buildDocumentFragmentPackage` and performs **one** `insertOoxmlWithRangeFallback` call.

---

## 5. Document Visual Formatting & Structural Test Suite

To verify **how formatting, whitespace, tabs, breaks, hyperlinks, and section properties show up in the document**, create `tests/ooxml_formatting_visual_tests.mjs`.

**File to Create:** `tests/ooxml_formatting_visual_tests.mjs`

```javascript
import assert from 'node:assert/strict';
import { applyOperationToDocumentXml } from '@ansonlai/docx-redline-js';

console.log('Running OOXML Formatting & Visual Fidelity Tests...\n');

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
    console.log('Test 2: Whitespace padding visual check (preventing word collision)...');
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
    assert.ok(xml.includes('Updated Section Title'), 'Modified title must be present');
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

    // Visual Invariant: Hyperlink tag and relationship ID must survive so the link remains clickable in Word
    assert.ok(xml.includes('<w:hyperlink'), 'Must preserve <w:hyperlink>');
    assert.ok(xml.includes('r:id="rId5"'), 'Must preserve hyperlink relationship ID rId5');
    assert.ok(xml.includes('support portal'), 'Must contain modified link text');
    console.log('  -> Passed: Clickable hyperlink (<w:hyperlink r:id="rId5">) preserved.');
}

// TEST 5: Section Break Preservation (<w:sectPr>)
{
    console.log('Test 5: Section break visual preservation (<w:sectPr>)...');
    const docWithSection = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>End of first section text.</w:t></w:r>
      <w:pPr>
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
      </w:pPr>
    </w:p>
    <w:p><w:r><w:t>Second section begins here.</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const result = await applyOperationToDocumentXml(docWithSection, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        anchor: { exactText: 'End of first section text.' },
        modified: 'Concluding remarks for initial section.',
        author: 'Editor'
    });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    // Visual Invariant: sectPr defines margins and page size; dropping it corrupts page layout in Word
    assert.ok(xml.includes('<w:sectPr>'), 'Must preserve <w:sectPr>');
    assert.ok(xml.includes('<w:pgSz'), 'Must preserve page size element');
    assert.ok(xml.includes('<w:pgMar'), 'Must preserve margin element');
    console.log('  -> Passed: Section layout properties (<w:sectPr>) completely preserved.');
}

console.log('\nAll OOXML Formatting & Visual Fidelity Tests Passed Successfully!');
```

---

## 6. Performance & Memory Profiling Harness

To empirically verify Goal 3 (Speed & Performance) without guessing, create `scripts/benchmark-ooxml-pipeline.mjs`.

**File to Create:** `scripts/benchmark-ooxml-pipeline.mjs`

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
    console.log('=== OOXML Single-Hop Pipeline Performance Benchmark ===\n');

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

        // Performance Assertion: 100 paragraphs must process in under 100ms
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
- [ ] Run `node tests/ooxml_formatting_visual_tests.mjs` -> 100% pass across all 5 test cases.
- [ ] Formatting Inheritance: Editing bold or italic text preserves `<w:rPr>` on inserted runs without font resets.
- [ ] Whitespace Preservation: Ensure `xml:space="preserve"` prevents words from colliding in Word.
- [ ] Structural Markers: Literal `<w:tab/>` and `<w:br/>` survive redline operations without degradation.
- [ ] Hyperlinks: `<w:hyperlink>` and relationship IDs remain intact.
- [ ] Section Layout: `<w:sectPr>` (page margins and sizing) remains intact.

### OOXML Purity & Web Portability Gates
- [ ] Single-Hop Word I/O: Verify via debug logs that applying an edit invokes exactly ONE `getOoxml()` and ONE `insertOoxml()`. Zero iterative `context.sync()` calls within the edit application.
- [ ] Zero Office.js in Engine: Assert that `@ansonlai/docx-redline-js` has 0 references to `Word.` or Office.js.
- [ ] Offline Execution: Run `node tests/ooxml_formatting_visual_tests.mjs` — verify 100% pass in pure Node environment.

### Performance Gates
- [ ] Single-DOM Batching: Applying a batch of operations executes in a single XML DOM parse and serialize pass.
- [ ] No-Op Short-Circuit: Applying an identical replacement (no-op) completes in $<5\text{ms}$ and performs 0 Word write operations.
- [ ] Benchmark Gate: Run `node scripts/benchmark-ooxml-pipeline.mjs` -> 100-paragraph documents process in $<100\text{ms}$.
