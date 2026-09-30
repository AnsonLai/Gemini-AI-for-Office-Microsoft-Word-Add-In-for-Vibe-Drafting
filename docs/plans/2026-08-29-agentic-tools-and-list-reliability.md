# Agentic Tools and List Reliability Plan (v0.8.0 Architecture)

**Date:** 2026-08-29  
**Last Updated:** 2026-09-29 (Upgrade target v0.8.1; v0.8.0 architecture and localized replacements)

**Status:** Active — Ready for Execution  
**Prerequisite Baseline:** **The Direct Upgrade Plan (`2026-09-09-docx-redline-js-v0.5.4-upgrade.md`) is COMPLETED.** Both the root add-in and `mcp/docx-server` are pinned to exact `@ansonlai/docx-redline-js@0.8.1`.

---

## 1. Guiding Principles & Priority Hierarchy

This plan is governed by three strict priorities, ordered from most important to least important:

1. **Accuracy when making changes (Highest Priority):**  
   Every document modification must be exact. Anchor text, occurrences, sub-paragraph offsets, paragraph attributes, character formatting, and tracked change metadata must be preserved without distortion. An edit must never corrupt document XML, duplicate revision IDs, or alter untargeted content. If an edit cannot be applied with 100% certainty, it must fail closed and preserve the original OOXML.
2. **Preserving OOXML-only approach (Second Priority):**  
   **No Office.js object-model synchronization or paragraph proxy manipulation in core logic.** All agent tools, operations, list conversions, table creations, diff reconciliations, and anchor validations must operate strictly on **plain OOXML** or canonical DOCX buffers. Office.js is strictly relegated to an ephemeral I/O transport shell (`getOoxml()` / `getFileAsync` on read, `insertOoxml()` / `insertFileFromBase64` on write).
3. **Speed and performance (Third Priority):**  
   Leverage `@ansonlai/docx-redline-js@0.8.1` caller-order-independent batching, localized replacements (`replacements: [{ find, replace }]`), single-pass serialization, and byte-exact short-circuiting when `hasChanges: false`.

---

## 2. Strategic Context & What Has Changed

With `@ansonlai/docx-redline-js@0.8.1`, the library natively handles what previously required complex add-in glue code:
- **v0.5.0 Structured Content (`structuredContent: true`)**: Converts Markdown tables, headings (`#`), and lists into native Word elements (`w:tbl`, `w:pStyle`, `w:numPr`) directly inside OOXML.
- **v0.6.0 Caller-Order-Independent Batches**: Targets resolve against the immutable initial document state; operations no longer need manual bottom-up sorting.
- **v0.7.0 Localized Exact Replacements (`replacements: [{ find, replace }]`)**: Operations specify exact phrase substitutions inside a target paragraph rather than repeating full paragraphs, **slashing prompt token payloads by 60–90%** and eliminating LLM truncation bugs.
- **v0.8.0 Universal Facade & Zero-Node Lifecycle**: `doc.inspect()` extracts paragraphs in one call; `doc.applyOperations()` handles batch mutation, `numbering.xml`, and `comments.xml` automatically.

---

## 3. Offloading & Deletion Matrix

| Local Plugin File / Function | Problem in Current Code | Native Library Replacement | Action for Implementation |
|---|---|---|---|
| [`src/taskpane/modules/docx-redline-js-integration/word-structured-list.js`](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/src/taskpane/modules/docx-redline-js-integration/word-structured-list.js) | 176 lines of hardcoded `<pkg:package>` and `<w:numbering>` strings with `abstractNumId="100"`. Violates OOXML purity. | Upstream `structuredContent: true` and native numbering management in `doc.applyOperations()`. | **DELETE FILE COMPLETELY.** |
| [`src/taskpane/modules/commands/list-level-utils.js`](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/src/taskpane/modules/commands/list-level-utils.js) | Regex parsing for list marker levels (`resolveInsertListItemLevel`). | Upstream list reconciliation natively calculates `ilvl` from marker depth and indents. | **DELETE FILE COMPLETELY.** |
| `detectRequestedContentKind` in `agentic-tools.js` | 25 lines of regex heuristics checking if the user asked for a "table". | Upstream `structuredContent: true` auto-detects markdown tables and transforms them to `w:tbl`. | **DELETE FUNCTION.** |
| `normalizeListItemsWithLevels` & `buildListMarkdown` in `agentic-tools.js` | Custom list normalization and markdown reconstruction. | Replaced by direct canonical operations with markdown text. | **DELETE FUNCTIONS.** |
| `verifyAnchor` & sliding window in `change-validation.js` | Complex index sliding trying to compensate for off-by-one paragraph indexes. | Use library's `getParagraphText` for canonical context and validate `replacements: [{ find }]` directly against paragraph text. | **SIMPLIFY TO EXACT MATCHING.** If target does not match, fail closed with `TARGET_NOT_FOUND`. |
| Bespoke block segmentation in `word-redline-runner.js` (`synthesizeMarkdownTableFromSourceRange`, etc.) | Custom regex splitting and table heuristics. | Handled natively by upstream engine. | **DELETE HEURISTICS.** |

---

## 4. Agent Tool to Canonical Operation Mapping Specification

Every tool called by the Gemini model maps deterministically to a canonical `docx-redline-js` operation. Below is the updated mapping specification:

| Agent Tool Name | Model Tool Call Arguments | Canonical `docx-redline-js` Operation Payload | Expected Behavior in OOXML & Word |
|---|---|---|---|
| **`modify_text`** | `{ paragraphIndex, anchorText, text, occurrence }` | `type: 'redline'`, `target: { paragraphIndex }`, `anchor: { exactText: anchorText, occurrence }`, `modified: text` | Emits paired `<w:del>` and `<w:ins>` with identical timestamp. Auto-detects markdown lists or tables into `w:numPr` or `w:tbl`. |
| **`replace_paragraph` (Localized)** | `{ paragraphIndex, anchorText, replacements: [{ find, replace }] }` | `type: 'redline'`, `target: { paragraphIndex }`, `replacements: [{ find, replace }]` | Performs surgical run splitting and carrier splitting inside the target paragraph without re-echoing unchanged sentences. |
| **`replace_range`** | `{ startParagraphIndex, endParagraphIndex, text }` | `type: 'redline'`, `target: { paragraphIndex: startParagraphIndex, endParagraphIndex }`, `modified: text` | Replaces multiple paragraphs with modified content. Preserves following section breaks (`w:sectPr`). |
| **`insert_paragraph`** | `{ targetParagraphIndex, text, position: 'before' \| 'after' }` | `type: 'insert'`, `target: { paragraphIndex: targetParagraphIndex }`, `position`, `modified: text` | Inserts a new tracked paragraph with `<w:pPrChange>` and `<w:ins>` on all child runs. |
| **`delete_paragraph`** | `{ paragraphIndex }` | `type: 'delete'`, `target: { paragraphIndex }` | Emits `<w:del>` over paragraph text and marks the paragraph mark deleted. Fails closed if paragraph has comments (`COMMENTED_CONTENT_DELETE`). |
| **`add_comment`** | `{ paragraphIndex, anchorText, commentText, occurrence }` | `type: 'comment'`, `target: { paragraphIndex }`, `anchor: { exactText: anchorText, occurrence }`, `comment: commentText` | Encloses targeted anchor in `<w:commentRangeStart/End>` and adds entry to `word/comments.xml`. |

### Standard Tool Response Format Returned to LLM

#### Success Response
```json
{
  "success": true,
  "action": "replace_paragraph",
  "paragraphIndex": 4,
  "receipt": {
    "revisionIds": ["101", "102"],
    "disposition": "applied",
    "validationSummary": { "valid": true, "issues": [] }
  },
  "message": "Clause successfully updated with tracked changes."
}
```

#### Error Response (Actionable Guidance)
```json
{
  "success": false,
  "action": "replace_paragraph",
  "paragraphIndex": 4,
  "errorCode": "TARGET_NOT_FOUND",
  "message": "Phrase 'December 31' was not found in paragraph 4.",
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

### Step 2: Refactor `agentic-tools.js` to Pure Intent-to-Operation Mapper
1. Open `src/taskpane/modules/commands/agentic-tools.js`.
2. Delete:
   - `detectRequestedContentKind`
   - `resolveInsertListItemLevel` import and calls
   - Direct `ReconciliationPipeline` instantiations
   - Inline prompt definitions (import from `redline-prompt.js`)
3. Refactor `applyRedlineChangeSet`:
   - Map `aiChanges` to canonical operations:
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
       if (Array.isArray(change.replacements)) {
         op.replacements = change.replacements;
       } else if (change.newContent || change.content) {
         op.modified = change.newContent || change.content;
       }
       return op;
     }
     ```
   - Pass the mapped array directly to the batch runner.

### Step 3: Modernize Prompting & Schema (`redline-prompt.js` & `change-validation.js`)
1. In `src/taskpane/modules/commands/redline-prompt.js`, update `REDLINE_DIFF_SCHEMA`:
   - Support `replacements: [{ find: string, replace: string }]`.
   - Update instructions: *"For edits inside an existing paragraph, provide `replacements: [{ find, replace }]` instead of repeating the entire paragraph."*
2. In `src/taskpane/modules/commands/change-validation.js`:
   - Validate that each `find` string in `replacements` exists within the targeted paragraph text.
