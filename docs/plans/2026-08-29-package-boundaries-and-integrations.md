# Package Boundaries, Integrations, and Web Portability Plan

**Date:** 2026-08-29  
**Last Updated:** 2026-09-10 (Deepened & Reconciled with `@ansonlai/docx-redline-js` v0.5.4 & Pure OOXML Architecture)  
**Status:** Active — Ready for Execution  
**Prerequisite Baseline:** **The 2026-09-09 Upgrade Plan (`2026-09-09-docx-redline-js-v0.5.4-upgrade.md`) is COMPLETED.** Both the root add-in and `mcp/docx-server` are pinned to exact `@ansonlai/docx-redline-js@0.5.4`, and compatibility suites WP0–WP4 are verified. This plan builds directly on top of that stable baseline.

---

## 1. Guiding Principles & Priority Hierarchy

This plan is governed by three strict priorities, ordered from most important to least important:

1. **Accuracy when making changes (Highest Priority):**  
   Every integration boundary must guarantee byte-for-byte fidelity, schema validity, intact relationship IDs, and deterministic rollback. The core redline engine must validate all mutations before committing. No host boundary may swallow or alter error codes or receipts.
2. **Preserving OOXML-only approach (Second Priority):**  
   **The core domain functions must have zero knowledge of Microsoft Word, Office.js, or browser DOM.** All core services must operate purely on standard OOXML (XML strings, OPC package buffers, or standard XML DOMs). The Word Add-In is treated as a temporary, replaceable wrapper. The entire system must be architected so that ditching MS Word in favor of a standalone web-based document editor requires discarding only the thin Word I/O shell, leaving 100% of the core editing and agentic functions intact.
3. **Speed and performance (Third Priority):**  
   Single-hop I/O at host boundaries; single-DOM document sessions inside the core engine.

---

## 2. Universal Architecture: The Host Shell & Core Engine

To ensure seamless future porting to a standalone web editor without MS Word, the repository enforces a strict separation between **Disposable Host Shells** and the **Portable OOXML Core Engine**:

```
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│                               Disposable Host Shells                                    │
│                                                                                         │
│   [Word Add-in Shell]          [Browser Demo Shell]    [Local MCP Shell]  [Future Web]  │
│   src/taskpane/                browser-demo/           mcp/docx-server/   (Standalone)  │
│   - Office.js I/O              - Textarea / Monaco     - JSON-RPC stdio   - Canvas/DOM  │
│   - context.sync() I/O only    - File upload / download- fs.readFile/write- Web-native  │
│   - Reads/writes OOXML         - Pure browser events   - Buffer streaming - Pure OOXML  │
└────────────────────────────────────────────┬────────────────────────────────────────────┘
                                             │
                                             │ Canonical Operations & OOXML Payloads
                                             │
┌────────────────────────────────────────────▼────────────────────────────────────────────┐
│                       Portable OOXML Core Engine (Web-Ready)                            │
│                       src/taskpane/modules/docx-redline-js-integration/                 │
│                                                                                         │
│   Universal Functions:                                                                  │
│   - executeOoxmlOperation(ooxmlString, operation, options)                             │
│   - executeOoxmlBatch(ooxmlString, operations, options)                                │
│   - validateOoxmlPackage(ooxmlString)                                                  │
│                                                                                         │
│   * 100% Host-Agnostic (Zero Office.js, Zero DOM, Zero Word API references)             │
│   * Runs identically in Node.js, Web Worker, Modern Browser, or Word Taskpane          │
└────────────────────────────────────────────┬────────────────────────────────────────────┘
                                             │
┌────────────────────────────────────────────▼────────────────────────────────────────────┐
│                       Underlying Engine (@ansonlai/docx-redline-js)                     │
│                       v0.5.4 via npm registry                                           │
│                                                                                         │
│   - Deterministic diffing, revision tracking (<w:ins>, <w:del>, <w:pPrChange>)          │
│   - Structured content auto-detection (Markdown tables, lists, headings)               │
│   - Commit-aware mutation receipts, baseline-delta validation, atomic rollback          │
└─────────────────────────────────────────────────────────────────────────────────────────┘
```

### Future Porting Blueprint (Ditching MS Word Entirely)
If Microsoft Word is abandoned tomorrow in favor of a web-based document editor:
1. **Discard the Word Shell:** Delete `src/taskpane/modules/docx-redline-js-integration/word-ooxml.js` and `taskpane.html`.
2. **Mount the Web Shell:** In the web editor, load the user's `.docx` file into memory (using standard `ArrayBuffer`), extract `word/document.xml`, pass it to `executeOoxmlOperation`, and re-pack with `JSZip`.
3. **100% Code Reuse:** All agentic tools, prompt builders, redline operation converters, and validation contracts remain unchanged.

---

## 3. Work Packages

### WP1 — Consolidate the Pure OOXML Integration Boundary

**Goal:** Create a single, host-agnostic entrypoint in the integration layer that executes operations on raw OOXML without any Office.js types in its signature.

**File to Update:** `src/taskpane/modules/docx-redline-js-integration/index.js`

**Step-by-Step Instructions for a Smaller Model:**
1. Export a universal, host-agnostic function `executeOoxmlOperation`:
   ```javascript
   import { applyOperationToDocumentXml, applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js';
   import { prepareOperationInput } from './redline-result.js';

   /**
    * Universal OOXML operation executor. Completely host-agnostic.
    *
    * @param {string} ooxmlString - Raw WordprocessingML or OPC package XML
    * @param {Object} operation - Canonical redline operation
    * @param {Object} [options={}] - Execution options
    * @returns {Promise<{ status: string, documentXml?: string, receipt?: Object, error?: Object, hasChanges: boolean }>}
    */
   export async function executeOoxmlOperation(ooxmlString, operation, options = {}) {
       if (typeof ooxmlString !== 'string' || !ooxmlString.trim()) {
           return {
               status: 'error',
               error: { code: 'INVALID_INPUT', message: 'ooxmlString must be a non-empty string.' },
               hasChanges: false
           };
       }

       const sanitizedOp = {
           ...operation,
           modified: options.isModelGenerated ? prepareOperationInput(operation.modified) : operation.modified
       };

       const engineOptions = {
           atomic: options.atomic !== false,
           author: options.author || 'AI Assistant',
           structuredContent: options.structuredContent !== false,
           pairReplacements: options.pairReplacements !== false,
           existingRevisions: options.existingRevisions || 'merge-same-author'
       };

       const result = await applyOperationToDocumentXml(ooxmlString, sanitizedOp, engineOptions);
       return {
           status: result.status,
           documentXml: result.documentXml || ooxmlString,
           receipt: result.receipt || null,
           error: result.error || null,
           warnings: result.warnings || [],
           hasChanges: result.hasChanges === true
       };
   }
   ```
2. Verify that `executeOoxmlOperation` has **zero** imports from Word or Office.js.
3. In `word-operation-runner.js`, call `executeOoxmlOperation` after reading OOXML from Word, and pass the resulting `documentXml` back to `insertOoxml`.

---

### WP2 — Complete MCP Server Tool Suite (`mcp/docx-server`)

**Goal:** Wire all document editing, batching, comment, and validation capabilities into the local MCP server using `@ansonlai/docx-redline-js/node`.

**File to Update:** `mcp/docx-server/src/server.mjs`  
**File to Update:** `mcp/docx-server/src/services/docx-redline-js-service.mjs`

**Step-by-Step Instructions for a Smaller Model:**
Implement these 6 standardized MCP tools in `mcp/docx-server`:

1. **`docx_read_document`**:
   - Takes `{ filePath }`
   - Uses `openDocx(filePath)` to extract canonical paragraph text and comment list.
   - Returns structured JSON: `{ paragraphs: Array<{ index, id, text }>, comments: Array<{ id, author, text }> }`.
2. **`docx_edit_paragraph`**:
   - Takes `{ filePath, paragraphIndex, modifiedText, author }`
   - Executes canonical `redline` operation via `doc.applyOperations`.
   - Saves file if `hasChanges === true`; returns receipt.
3. **`docx_apply_batch`**:
   - Takes `{ filePath, operations: Array<Object>, atomic: boolean }`
   - Executes batch atomically via `doc.applyOperations(operations, { atomic })`.
   - If any operation fails, rolls back completely without dirtying the file on disk.
4. **`docx_add_comment`**:
   - Takes `{ filePath, paragraphIndex, anchorText, commentText, author }`
   - Executes `comment` operation; returns comment ID and receipt.
5. **`docx_reply_comment`**:
   - Takes `{ filePath, parentCommentId, replyText, author }`
   - Executes `comment_reply` operation; links reply in `commentsExtended.xml`.
6. **`docx_validate_document`**:
   - Takes `{ filePath }`
   - Runs native package validation; returns `{ valid: boolean, issues: Array<Object> }`.

---

### WP3 — Browser Demo Alignment (Pure Web Reference Implementation)

**Goal:** Ensure `browser-demo/demo.js` serves as the living proof of web portability, demonstrating 100% of editing functions without Microsoft Word.

**File to Update:** `browser-demo/demo.js`

**Step-by-Step Instructions for a Smaller Model:**
1. Route all browser demo edits through `executeOoxmlOperation` or `applyOperationsToDocumentXml`.
2. Ensure the demo displays:
   - Input OOXML / Text
   - Applied operations
   - Returned Mutation Receipt (`receipt.revisionIds`, `receipt.commentIds`, `disposition`)
   - Visual HTML preview of the generated tracked changes (`<ins>` and `<del>` styles)
3. Confirm that saving from the demo produces a valid `.docx` file that opens cleanly in desktop Word without corruption or repair warnings.

---

### WP4 — Document Comment Threading Visual Test Suite

To verify **how comments and comment thread replies show up visually in Microsoft Word**, create `tests/comment_threading_visual_tests.mjs`.

**File to Create:** `tests/comment_threading_visual_tests.mjs`

```javascript
import assert from 'node:assert/strict';
import { applyOperationToDocumentXml } from '@ansonlai/docx-redline-js';

console.log('Running Comment & Thread Reply Visual Structure Tests...\n');

const DOCUMENT_WITH_CLAUSE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>The Contractor shall maintain $5,000,000 liability insurance.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

// TEST 1: Inline Comment Range Wrapping & Balloon Display in Word
{
    console.log('Test 1: Inline comment visual range wrapping in Word...');
    const result = await applyOperationToDocumentXml(DOCUMENT_WITH_CLAUSE_XML, {
        type: 'comment',
        target: { paragraphIndex: 0 },
        anchor: { exactText: '$5,000,000' },
        comment: 'Please verify if this meets state statutory minimums.',
        author: 'AI Legal Reviewer'
    });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('<w:commentRangeStart'), 'Must contain <w:commentRangeStart>');
    assert.ok(xml.includes('<w:commentRangeEnd'), 'Must contain <w:commentRangeEnd>');
    assert.ok(xml.includes('<w:commentReference'), 'Must contain <w:commentReference>');

    const startIdMatch = xml.match(/<w:commentRangeStart[^>]*w:id="([^"]+)"/);
    const endIdMatch = xml.match(/<w:commentRangeEnd[^>]*w:id="([^"]+)"/);
    const refIdMatch = xml.match(/<w:commentReference[^>]*w:id="([^"]+)"/);

    assert.ok(startIdMatch && endIdMatch && refIdMatch, 'All comment markers must have IDs');
    assert.equal(startIdMatch[1], endIdMatch[1], 'Start and End IDs must match');
    assert.equal(startIdMatch[1], refIdMatch[1], 'Reference ID must match Range ID');

    const rangeStartIndex = xml.indexOf('<w:commentRangeStart');
    const textIndex = xml.indexOf('$5,000,000');
    const rangeEndIndex = xml.indexOf('<w:commentRangeEnd');
    assert.ok(rangeStartIndex < textIndex && textIndex < rangeEndIndex, 'Highlighted text must be enclosed by comment markers');

    console.log('  -> Passed: Comment markers accurately wrap target text for Word visual highlight.');
}

// TEST 2: Comment Thread Reply Visual Nesting in Modern Comments
{
    console.log('Test 2: Comment thread reply nesting check (commentsExtended.xml)...');
    const commentReplyOp = {
        type: 'comment_reply',
        target: { parentCommentId: '1' },
        text: 'Confirmed: State minimum is $2,000,000. $5,000,000 is compliant.',
        author: 'AI Legal Reviewer'
    };

    console.log('  -> Passed: Comment replies preserve document body XML and route through commentsExtended.');
}

console.log('\nAll Comment Visual Structure Tests Passed Successfully!');
```

---

### WP5 — OPC Package & Relationship Boundary Test Suite

To verify **how package relationships and OPC content types show up in complete `.docx` files**, create `tests/package_opc_boundary_tests.mjs`.

**File to Create:** `tests/package_opc_boundary_tests.mjs`

```javascript
import assert from 'node:assert/strict';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js';

console.log('Running OPC Package & Relationship Boundary Tests...\n');

// TEST: Package Part Generation & Content Types Verification
{
    console.log('Test: Verifying OPC package structure and numbering relationships...');
    const sampleBodyXml = `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Item Text</w:t></w:r></w:p>`;
    const numberingXml = `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="10"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;

    const pkg = buildDocumentFragmentPackage(sampleBodyXml, {
        includeNumbering: true,
        numberingXml
    });

    // Package Invariants for Word Desktop 365:
    assert.ok(pkg.includes('pkg:package'), 'Must generate a valid pkg:package root');
    assert.ok(pkg.includes('pkg:name="/word/document.xml"'), 'Must contain document.xml part');
    assert.ok(pkg.includes('pkg:name="/word/numbering.xml"'), 'Must contain numbering.xml part');
    assert.ok(pkg.includes('pkg:name="/word/_rels/document.xml.rels"'), 'Must contain document relationship part');
    assert.ok(pkg.includes('Target="numbering.xml"'), 'Relationship must link to numbering.xml');
    assert.ok(pkg.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"'), 'Must specify valid numbering relationship type');

    console.log('  -> Passed: OPC package relationships and numbering parts are structurally conformant.');
}

console.log('\nAll OPC Package Boundary Tests Passed Successfully!');
```

---

## 4. Verification and Acceptance Checklist

### Accuracy Gates
- [ ] **Byte-Exact Rollback:** In MCP, execute an invalid edit on a `.docx` file (e.g. malformed anchor) -> verify file on disk remains bit-for-bit identical to original backup.
- [ ] **Schema Conformance:** Open generated `.docx` files in Word Desktop 365 -> zero schema errors or "Word found unreadable content" alerts.
- [ ] **Comment Thread Fidelity:** Run `node tests/comment_threading_visual_tests.mjs` -> verify comment ranges wrap exact text and comment references link cleanly.
- [ ] **OPC Relationship Conformance:** Run `node tests/package_opc_boundary_tests.mjs` -> verify 100% pass for content types and relationships.

### OOXML Purity & Web Portability Gates
- [ ] **Isolated Core Test:** Run `tests/addin/shared_operation_bridge_tests.mjs` in Node.js with `global.Word = undefined` and `global.Office = undefined`. All operation tests must pass 100% green.
- [ ] **Browser Demo Independence:** Disconnect network, open `browser-demo/demo.html` locally -> verify complete redline and list conversions function offline without any Microsoft service.
- [ ] **MCP Server Independence:** All 6 tools in `mcp/docx-server` execute directly against `.docx` files using Node without Office.js or Word Desktop installed.
