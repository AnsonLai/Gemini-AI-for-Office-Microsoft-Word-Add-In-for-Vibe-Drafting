# Agentic Tools and List Reliability Plan (Pure OOXML Architecture)

**Date:** 2026-08-29  
**Last Updated:** 2026-09-10 (Deepened & Reconciled with `@ansonlai/docx-redline-js` v0.5.4 & Pure OOXML Architecture)  
**Status:** Active — Ready for Execution  
**Prerequisite Baseline:** **The 2026-09-09 Upgrade Plan (`2026-09-09-docx-redline-js-v0.5.4-upgrade.md`) is COMPLETED.** Both the root add-in and `mcp/docx-server` are pinned to exact `@ansonlai/docx-redline-js@0.5.4`, and compatibility suites WP0–WP4 are verified. This plan builds directly on top of that stable baseline.

---

## 1. Guiding Principles & Priority Hierarchy

This plan is governed by three strict priorities, ordered from most important to least important:

1. **Accuracy when making changes (Highest Priority):**  
   Every document modification must be exact. Anchor text, occurrences, sub-paragraph offsets, paragraph attributes, character formatting, and tracked change metadata must be preserved without distortion. An edit must never corrupt document XML, duplicate revision IDs, or alter untargeted content. If an edit cannot be applied with 100% certainty, it must fail closed and preserve the original OOXML.
2. **Preserving OOXML-only approach (Second Priority):**  
   **No Office.js object-model synchronization or paragraph proxy manipulation in core logic.** All agent tools, operations, list conversions, table creations, diff reconciliations, and anchor validations must operate strictly on **plain OOXML** (XML strings, OPC package parts, and DOM nodes). Office.js is strictly relegated to an ephemeral I/O transport shell (`getOoxml()` on read, `insertOoxml()` on write). All functions must be 100% portable to a pure web editor in the future, ditching Microsoft Word entirely.
3. **Speed and performance (Third Priority):**  
   Leverage `@ansonlai/docx-redline-js` single-DOM batching, single-pass serialization, and byte-exact short-circuiting when `hasChanges: false`.

---

## 2. Strategic Context & What Has Changed

With the 2026-09-09 upgrade completed, `@ansonlai/docx-redline-js` v0.5.4 provides all the heavy lifting natively:
- **v0.4.0:** Multi-level list & bullet reconciliation natively supported without corrupting numbering definitions.
- **v0.5.0:** **Structured Content Auto-Detection (`structuredContent: true` by default)** converts Markdown tables, headings (`#`), and lists into native Word elements (`w:tbl`, `w:pStyle`, `w:numPr`) directly inside the OOXML. Also exported `planStructuredReplacement` and `analyzeStructuredContent`.
- **v0.5.0:** **Commit-Aware Mutation Receipts** (`receipt.revisionIds`, `receipt.commentIds`, `disposition`, `validationSummary`).
- **v0.5.0:** **Smart Same-Author Revision Merging (`merge-same-author` by default)**: Subsequent edits cleanly re-diff against the pre-revision baseline ($T_0$), avoiding nested or piled-up revisions.
- **v0.5.3–v0.5.4:** Added explicit paragraph `restore` operations, rejected-view insertions (`del(A) / ins(B) / del(A)`), and whole-document duplicate revision ID detection.

### The Architectural Shift: The "Pure OOXML Core"
Instead of `agentic-tools.js` interacting with `Word.Paragraph` proxies or managing custom `<w:numbering>` XML packages, the flow is:
```
┌────────────────────────────────────────────────────────────┐
│                    Host Transport Shell                    │
│  - Word Add-in: context.document.getOoxml()                │
│  - Web App: document.xml from .docx buffer or web state    │
│  - Node/MCP: fs.readFile() / openDocx()                    │
└─────────────────────────────┬──────────────────────────────┘
                              │ Plain OOXML String
┌─────────────────────────────▼──────────────────────────────┐
│                  Pure OOXML Core Domain                    │
│  1. Convert agent tool call to Canonical Operation         │
│  2. Execute applyOperationToDocumentXml(ooxml, operation)  │
│  3. Validate OOXML and evaluate Mutation Receipt           │
└─────────────────────────────┬──────────────────────────────┘
                              │ Updated OOXML String
┌─────────────────────────────▼──────────────────────────────┐
│                    Host Transport Shell                    │
│  - Word Add-in: range.insertOoxml(updatedOoxml, 'Replace') │
│  - Web App: update document state / download .docx         │
│  - Node/MCP: fs.writeFile()                                │
└────────────────────────────────────────────────────────────┘
```

---

## 3. Reconciled Offloading & Deletion Matrix

| Local Plugin File / Function | Problem in Current Code | Native Library Replacement | Action for Implementation |
|---|---|---|---|
| [`src/taskpane/modules/docx-redline-js-integration/word-structured-list.js`](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/src/taskpane/modules/docx-redline-js-integration/word-structured-list.js) | 176 lines of hardcoded `<pkg:package>` and `<w:numbering>` strings with `abstractNumId="100"`. Violates OOXML purity. | Upstream `structuredContent: true` and `wrapInDocumentFragment(ooxml, { includeNumbering: true })`. | **DELETE FILE COMPLETELY.** |
| [`src/taskpane/modules/commands/list-level-utils.js`](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/src/taskpane/modules/commands/list-level-utils.js) | Regex parsing for list marker levels (`resolveInsertListItemLevel`). | Upstream list reconciliation natively calculates `ilvl` from marker depth and indents. | **DELETE FILE COMPLETELY.** |
| `detectRequestedContentKind` in `agentic-tools.js` | 25 lines of regex heuristics checking if the user asked for a "table". | Upstream `structuredContent: true` auto-detects markdown tables and transforms them to `w:tbl`. | **DELETE FUNCTION.** |
| `normalizeListItemsWithLevels` & `buildListMarkdown` in `agentic-tools.js` | Custom list normalization and markdown reconstruction. | Replaced by direct canonical operations with markdown text. | **DELETE FUNCTIONS.** |
| `verifyAnchor` & sliding window in `change-validation.js` | Complex index sliding trying to compensate for off-by-one paragraph indexes. | Use library's `getParagraphText` for canonical context and pass `anchor.exactText`, `anchor.occurrence`, `anchor.offset`. | **SIMPLIFY TO EXACT MATCHING.** If target does not match, fail closed with `TARGET_NOT_FOUND`. |
| Bespoke block segmentation in `word-redline-runner.js` (`segmentNativeInsertionBlocks`, `buildListBlockFragment`, etc.) | Custom regex splitting of text vs list vs table blocks. | Use public `planStructuredReplacement` and `wrapInDocumentFragment` from library. | **REPLACE WITH LIBRARY CALLS.** |

---

## 4. Agent Tool to Canonical Operation Mapping Specification

Every tool called by the Gemini model maps deterministically to a canonical `docx-redline-js` operation. Below is the exact mapping specification for all 5 core tools:

| Agent Tool Name | Model Tool Call Arguments | Canonical `docx-redline-js` Operation Payload | Expected Behavior in OOXML & Word |
|---|---|---|---|
| **`modify_text`** | `{ paragraphIndex, anchorText, text, occurrence }` | `type: 'redline'`, `target: { paragraphIndex }`, `anchor: { exactText: anchorText, occurrence }`, `modified: text` | Emits paired `<w:del>` and `<w:ins>` with identical timestamp. If `text` contains markdown list or table, auto-detects into `w:numPr` or `w:tbl`. |
| **`replace_range`** | `{ startParagraphIndex, endParagraphIndex, text }` | `type: 'redline'`, `target: { paragraphIndex: startParagraphIndex, endParagraphIndex }`, `modified: text` | Replaces multiple paragraphs with the modified content. Preserves following section breaks (`w:sectPr`). |
| **`insert_paragraph`** | `{ targetParagraphIndex, text, position: 'before' \| 'after' }` | `type: 'insert'`, `target: { paragraphIndex: targetParagraphIndex }`, `position`, `modified: text` | Inserts a new tracked paragraph with `<w:pPrChange>` and `<w:ins>` on all child runs. |
| **`delete_paragraph`** | `{ paragraphIndex }` | `type: 'delete'`, `target: { paragraphIndex }` | Emits `<w:del>` over paragraph text and marks the paragraph mark deleted (`w:pPrChange/w:rPr/w:del`). Fails closed if paragraph has comments (`COMMENTED_CONTENT_DELETE`). |
| **`add_comment`** | `{ paragraphIndex, anchorText, commentText, occurrence }` | `type: 'comment'`, `target: { paragraphIndex }`, `anchor: { exactText: anchorText, occurrence }`, `comment: commentText` | Encloses targeted anchor in `<w:commentRangeStart/End>` and adds entry to `word/comments.xml`. |

### Standard Tool Response Format Returned to LLM
Every tool execution returns a structured, predictable response object to the AI agent:

#### Success Response
```json
{
  "success": true,
  "action": "modify_text",
  "paragraphIndex": 4,
  "receipt": {
    "revisionIds": ["101", "102"],
    "disposition": "applied",
    "validationSummary": { "valid": true, "issues": [] }
  },
  "message": "Clause successfully updated with tracked changes."
}
```

#### Error Response (Actionable Error Guidance)
```json
{
  "success": false,
  "action": "modify_text",
  "paragraphIndex": 4,
  "errorCode": "TARGET_NOT_FOUND",
  "message": "Anchor text 'December 31' was not found at paragraph index 4.",
  "hint": "The document text may have changed. Please inspect the current paragraph text before retrying."
}
```

---

## 5. Detailed Step-by-Step Implementation Instructions

### Step 1: Remove Bespoke Numbering & List Construction Modules
1. Delete `src/taskpane/modules/docx-redline-js-integration/word-structured-list.js`.
2. Delete `src/taskpane/modules/commands/list-level-utils.js`.
3. In `src/taskpane/modules/docx-redline-js-integration/index.js`, remove exports of `applyStructuredListDirectOoxml`.
4. In `tests/no_legacy_shared_operation_bridge_tests.mjs`, add `word-structured-list.js` and `list-level-utils.js` to the resurrection guard list.

### Step 2: Refactor `agentic-tools.js` to a Pure Intent-to-Operation Mapper
1. Open `src/taskpane/modules/commands/agentic-tools.js`.
2. Delete:
   - `detectRequestedContentKind`
   - `resolveInsertListItemLevel` import and calls
   - Direct `ReconciliationPipeline` instantiations
   - Inline prompt definitions (all prompts must import from `redline-prompt.js`)
3. Refactor `applyRedlineChangeSet`:
   - Take `aiChanges` array.
   - Map each change to a canonical operation using the exact mapping function:
     ```javascript
     export function mapChangeToCanonicalOperation(change, author = 'AI Assistant') {
       const op = {
         type: change.action === 'delete' ? 'delete' : 'redline',
         author,
         options: {
           structuredContent: true,
           pairReplacements: true,
           existingRevisions: 'merge-same-author'
         }
       };
       if (Number.isInteger(change.paragraphIndex)) {
         op.target = { paragraphIndex: change.paragraphIndex };
       }
       if (change.paragraphId) {
         op.target = op.target || {};
         op.target.paragraphId = change.paragraphId;
       }
       if (change.anchorText) {
         op.anchor = {
           exactText: change.anchorText,
           occurrence: change.occurrence || 1,
           offset: change.offset || 0
         };
       }
       if (change.action !== 'delete') {
         op.modified = change.text || change.modified || '';
       }
       return op;
     }
     ```
4. Pass the array of canonical operations to `word-operation-runner.js` for single-hop execution.

### Step 3: Pure OOXML Paragraph Targeting in `change-validation.js`
1. Open `src/taskpane/modules/commands/change-validation.js`.
2. Replace index sliding heuristics with strict canonical text verification:
   - Extract canonical text using `getParagraphText(paragraphNode)` from `@ansonlai/docx-redline-js`.
   - If `anchorText` is specified, verify that the paragraph at `paragraphIndex` contains `anchorText`.
   - If it does not match, return `{ valid: false, error: 'TARGET_NOT_FOUND', message: `Anchor "${anchorText}" not found at paragraph index ${paragraphIndex}.` }`.
   - Do NOT guess or slide the index to neighboring paragraphs. Accuracy comes first.

### Step 4: Streamline Native Insertion Fallback in `word-redline-runner.js`
1. Open `src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js`.
2. Remove:
   - `segmentNativeInsertionBlocks`
   - `isMarkdownTableStart`
   - `bulletLineToLiteralText`
   - `buildListBlockFragment`
   - `buildTableBlockFragment`
3. Replace with library-provided structured fragment generation:
   - Call `planStructuredReplacement(content, options)` or `wrapInDocumentFragment` with `includeNumbering: true`.
   - Insert the resulting OOXML package string once into the target range.

---

## 6. Document-Level Visual & Structural Test Suite

To verify **how changes show up in the document**, create `tests/document_visual_structure_tests.mjs`. This suite inspects the generated OOXML structure and verifies that Word's visual presentation requirements are satisfied.

**File to Create:** `tests/document_visual_structure_tests.mjs`

```javascript
import assert from 'node:assert/strict';
import {
    applyOperationToDocumentXml,
    acceptTrackedChangesInOoxml,
    rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';

console.log('Running Document Visual & Structural Tests...\n');

const BASELINE_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Clause 1: The payment is due on December 31.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Clause 2: Deliverables shall be provided as follows.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Clause 3: Fee schedule.</w:t></w:r></w:p>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="SectionAnchor"/>
      <w:r><w:t>Clause 4: See details in note.</w:t></w:r>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
  </w:body>
</w:document>`;

// TEST 1: Reviewing Pane Card Grouping (Paired Replacement Timestamps)
{
    console.log('Test 1: Reviewing Pane paired replacement visual grouping...');
    const result = await applyOperationToDocumentXml(BASELINE_DOCUMENT_XML, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        anchor: { exactText: 'December 31' },
        modified: 'November 30',
        author: 'AI Reviewer'
    }, { pairReplacements: true });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('<w:del'), 'Must contain <w:del>');
    assert.ok(xml.includes('<w:ins'), 'Must contain <w:ins>');

    const delDateMatch = xml.match(/<w:del[^>]*w:date="([^"]+)"/);
    const insDateMatch = xml.match(/<w:ins[^>]*w:date="([^"]+)"/);
    assert.ok(delDateMatch && insDateMatch, 'Both tags must have timestamp attributes');
    assert.equal(delDateMatch[1], insDateMatch[1], 'del and ins must share exact date timestamp for Reviewing Pane grouping');
    console.log('  -> Passed: <w:del> and <w:ins> share identical timestamp for unified card display.');
}

// TEST 2: Multi-Level List Rendering (Dynamic Numbering vs Raw Text)
{
    console.log('Test 2: Multi-Level List Visual Rendering in Word...');
    const listMarkdown = "1. First Item\n   a. Sub-Item Alpha\n   b. Sub-Item Beta\n2. Second Item";
    const result = await applyOperationToDocumentXml(BASELINE_DOCUMENT_XML, {
        type: 'redline',
        target: { paragraphIndex: 1 },
        modified: listMarkdown,
        author: 'AI Reviewer'
    }, { structuredContent: true });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('w:val="ListParagraph"'), 'List items must have ListParagraph style in Word');
    assert.ok(xml.includes('<w:ilvl w:val="0"/>'), 'Top-level list must have ilvl 0');
    assert.ok(xml.includes('<w:ilvl w:val="1"/>'), 'Sub-item must have ilvl 1');

    assert.ok(!xml.includes('<w:t>1. First Item</w:t>'), 'Raw marker "1." must be stripped from text node');
    assert.ok(!xml.includes('<w:t>a. Sub-Item Alpha</w:t>'), 'Raw marker "a." must be stripped from text node');
    assert.ok(xml.includes('<w:t>First Item</w:t>'), 'Text must contain clean item body');
    assert.ok(xml.includes('<w:t>Sub-Item Alpha</w:t>'), 'Text must contain clean sub-item body');
    console.log('  -> Passed: Native Word list properties applied, raw markers stripped, indentation configured.');
}

// TEST 3: Table Structure and Grid Rendering
{
    console.log('Test 3: Markdown Table Visual Grid Rendering...');
    const tableMarkdown = "| Role | Person |\n|---|---|\n| Architect | Alice |\n| Reviewer | Bob |";
    const result = await applyOperationToDocumentXml(BASELINE_DOCUMENT_XML, {
        type: 'redline',
        target: { paragraphIndex: 2 },
        modified: tableMarkdown,
        author: 'AI Reviewer'
    }, { structuredContent: true });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('<w:tbl>'), 'Must contain <w:tbl> element');
    assert.ok(xml.includes('<w:tblPr>'), 'Must contain table properties');
    assert.ok(xml.includes('<w:tblBorders>'), 'Must contain table borders');
    assert.ok(xml.includes('<w:tr>'), 'Must contain table rows');
    assert.ok(xml.includes('<w:tc>'), 'Must contain table cells');

    const rowCount = (xml.match(/<w:tr[\s>]/g) || []).length;
    assert.equal(rowCount, 3, 'Table must render exactly 3 rows');
    console.log('  -> Passed: Valid <w:tbl> with borders, cells, and 3 rows rendered.');
}

// TEST 4: Bookmark & Footnote Reference Preservation
{
    console.log('Test 4: Bookmark and footnote reference preservation...');
    const result = await applyOperationToDocumentXml(BASELINE_DOCUMENT_XML, {
        type: 'redline',
        target: { paragraphIndex: 3 },
        anchor: { exactText: 'note' },
        modified: 'explanatory note',
        author: 'AI Reviewer'
    });

    assert.equal(result.status, 'success');
    const xml = result.documentXml;

    assert.ok(xml.includes('<w:bookmarkStart w:id="0" w:name="SectionAnchor"/>'), 'Bookmark start must survive edit');
    assert.ok(xml.includes('<w:bookmarkEnd w:id="0"/>'), 'Bookmark end must survive edit');
    assert.ok(xml.includes('<w:footnoteReference w:id="1"/>'), 'Footnote reference must survive edit');
    console.log('  -> Passed: Bookmarks and footnotes intact across targeted redline.');
}

// TEST 5: Accept All / Reject All Visual Symmetry Invariant
{
    console.log('Test 5: Word Accept All / Reject All Visual Symmetry Invariant...');
    const editedResult = await applyOperationToDocumentXml(BASELINE_DOCUMENT_XML, {
        type: 'redline',
        target: { paragraphIndex: 0 },
        anchor: { exactText: 'December 31' },
        modified: 'November 30',
        author: 'AI Reviewer'
    });

    assert.equal(editedResult.status, 'success');
    const modifiedXml = editedResult.documentXml;

    const acceptedXml = acceptTrackedChangesInOoxml(modifiedXml);
    assert.ok(!acceptedXml.includes('<w:del>'), 'Accept All must remove all <w:del>');
    assert.ok(!acceptedXml.includes('<w:ins>'), 'Accept All must unwrap all <w:ins>');
    assert.ok(acceptedXml.includes('November 30'), 'Accept All must display modified text');
    assert.ok(!acceptedXml.includes('December 31'), 'Accept All must not display original text');

    const rejectedXml = rejectTrackedChangesInOoxml(modifiedXml);
    assert.ok(!rejectedXml.includes('<w:del>'), 'Reject All must unwrap all <w:del>');
    assert.ok(!rejectedXml.includes('<w:ins>'), 'Reject All must remove all <w:ins>');
    assert.ok(rejectedXml.includes('December 31'), 'Reject All must display original text');
    assert.ok(!rejectedXml.includes('November 30'), 'Reject All must not display modified text');

    console.log('  -> Passed: Accept All and Reject All produce exact visual and text symmetry.');
}

console.log('\nAll Document Visual & Structural Tests Passed Successfully!');
```

---

## 7. Verification and Acceptance Checklist

### Accuracy Gates
- [ ] Multi-level List Conversion: Convert `1. Header\n  a. Sub-item` -> verify generated OOXML contains `<w:numPr>` with correct `w:ilvl` (0 and 1) and valid `w:numId`.
- [ ] Legal Numbering Preservation: Convert `2.2.1 Section` -> verify it is NOT re-numbered to `1.`.
- [ ] Markdown Table Conversion: Convert a 3x3 markdown table -> verify clean `<w:tbl>` with correct rows, cells, and borders.
- [ ] Exact Anchor Matching: Target paragraph with multiple occurrences of a word; verify `anchor: { exactText: 'term', occurrence: 2 }` modifies the second occurrence only.
- [ ] Structural Metadata Preservation: Run Test 4 -> verify `<w:bookmarkStart/End>` and `<w:footnoteReference>` survive redlines untouched.
- [ ] Failure Invariant: Intentionally supply a non-existent anchor text -> verify operation returns `TARGET_NOT_FOUND` and leaves the document OOXML completely unchanged.

### Document Visual Appearance Gates
- [ ] Run `node tests/document_visual_structure_tests.mjs` -> verify 100% pass for reviewing pane card grouping, list properties, table grids, bookmarks, and accept/reject symmetry.
- [ ] Reviewing Pane: Paired replacements share identical timestamps.
- [ ] List View: Markers stripped from `<w:t>`, `w:numPr` and `w:ilvl` configured.
- [ ] Table View: `w:tbl`, `w:tblBorders`, and correct row/cell count.

### OOXML Purity & Web Portability Gates
- [ ] Zero Office.js in Domain Logic: Grep `src/taskpane/modules/commands/` and `src/taskpane/modules/docx-redline-js-integration/` — verify zero calls to Word paragraph proxy traversal methods outside the designated I/O boundary in `word-ooxml.js`.
- [ ] Web Portability Test: Run all canonical operations against sample XML strings in Node using `tests/addin/word_operation_runner_adapter_tests.mjs` without initializing Office.js or mocking Word globals.
