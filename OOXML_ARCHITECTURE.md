# OOXML Reconciliation Architecture

This document outlines the architecture of the **OOXML Reconciliation Engine**, located in `src/taskpane/modules/reconciliation/`. This system is responsible for converting AI-generated Markdown text into valid Office Open XML (Word) structures while preserving existing document formatting and generating precise "Track Changes" (Redlines).

## High-Level Architecture

The system operates in two distinct modes depending on the complexity of the operation:

1.  **Reconstruction Pipeline** (Standard): Parses original OOXML, diffs against new text, and rebuilds the paragraph from scratch. Used for standard text edits and list generation.
2.  **Surgical Hybrid Engine** (Complex): Direct DOM manipulation. Used for Tables (to preserve cell structure) and specific "Format Removal" operations.

```mermaid
graph TD
    Input[OOXML + Markdown Input] --> Router{OxmlEngine Router}
    
    Router --Tables/Surgical--> Hybrid[Hybrid Engine\n(In-Place DOM)]
    Router --Lists/Text--> Pipeline[Reconciliation Pipeline\n(Reconstruction)]
    
    subgraph "Reconciliation Pipeline"
        Ingest[Ingestion\n(XML -> RunModel)] --> Diff[Diff Engine\n(Word-Level)]
        Diff --> Patch[Patching\n(Apply Diffs)]
        Patch --> Serialize[Serialization\n(RunModel -> XML)]
    end
    
    subgraph "Hybrid Engine"
        DOM[DOM Parser] --> Search[Node Search]
        Search --> Replace[Surgical w:ins/w:del]
        Replace --> Wrap[Fragment Wrapping]
    end

    Pipeline --> Output[Final OOXML]
    Hybrid --> Output
```

---

## 1. The Reconciliation Pipeline (`pipeline.js`)

The default path for text and list operations. It treats a paragraph as a sequence of "Runs" and reconstructs it.

### Stage 1: Ingestion (`ingestion.js`)
Parses raw OOXML string into a `RunModel`, a linear representation of the document content that abstracts away XML complexity.

**Key Responsibilities:**
*   Parses recursivley, handling nested structures like `w:hyperlink`, `w:sdt` (Content Controls), and `w:smartTag`.
*   Captures `pPr` (Paragraph Properties) to preserve paragraph-level style.
*   **Offset Generation:** Assigns `startOffset` and `endOffset` to every run, creating a coordinate system for the Diff Engine.

**The RunModel Structure:**
```javascript
[
  { 
    kind: 'text',           // RunKind (TEXT, CONTAINER_START, HYPERLINK, etc.)
    text: 'Hello',          // The actual content contributing to the document text
    startOffset: 0,         // Absolute character start position
    endOffset: 5,           // Absolute character end position
    rPrXml: '<w:rPr>...</w:rPr>', // XML string of the run properties (formatting)
    containerContext: 'sdt_1' // Reference to parent container if any
  },
  ...
]
```

### Stage 2: Markdown Pre-processing (`markdown-processor.js`)
*   Strips Markdown symbols (`**`, `__`) from the new text to create a "Clean" string for comparison.
*   Generates **Format Hints**: An overlay of formatting instructions derived from the Markdown.
    *   `{ start: 0, end: 5, format: { bold: true } }`

### Stage 3: Diffing (`diff-engine.js`)
Computes **Word-Level** diffs between the "Accepted" original text (from the RunModel) and the "Clean" new text.

**Tokenization Strategy (`wordsToChars`):**
*   **Problem:** Standard diff algorithms operate on characters. We want word-level granularity.
*   **Solution:** Uses a hashing trick where every unique word is mapped to a unique Unicode character.
    *   `"the quick brown"` -> `"\u0001\u0002\u0003"`
*   The diff algorithm runs on these "character" strings, and the result is translated back to words.

### Stage 4: Patching (`patching.js`)
Applies the diff operations to the `RunModel`.

**Splitting Logic:**
*   If a diff operation (e.g., DELETE or INSERT) starts in the middle of an existing run (e.g., changing "format**ted**" to "format**ting**"), the engine splits the run at the exact boundary.
*   `{ text: "formatted" }` -> `[{ text: "format" }, { text: "ted" }]`.

**Style Inheritance:**
*   **Insertions** inherit style from the surrounding context.
    *   If appending to a word (`Hello` -> `Hello World`), it inherits from `Hello`.
    *   If prepending, it inherits from the following word.

### Stage 5: Serialization (`serialization.js`)
*   Converts the modified `RunModel` back into valid OOXML XML strings.
*   Handles namespace management (`xmlns:w` is implicitly handled by the wrapping).
*   Reconstructs nested containers (Hyperlinks, SDTs) by tracking `CONTAINER_START` and `CONTAINER_END` tokens.

---

## 2. Hybrid / Surgical Engine (`oxml-engine.js`)

Handles complex scenarios where destroying and recreating the standard paragraph structure would cause data loss (e.g., inside Tables) or where Word's API behavior requires specific XML structures.

### Virtual Grid System (`ingestion.js` -> `ingestTableToVirtualGrid`)
To reconcile tables, the engine converts the hierarchical XML (`w:tr` -> `w:tc`) into a 2D **Virtual Grid**.

*   **Handling Merges:**
    *   Parses `gridSpan` (horizontal) and `vMerge` (vertical) attributes.
    *   Fills the grid: `grid[row][col]` points to a Cell object.
    *   Merged areas reference the same "Origin" cell or "Continuation" markers.

*   **Diffing Grids (`table-reconciliation.js`):**
    *   Compares the Markdown table structure against the Virtual Grid dimensions.
    *   Generates `INSERT_ROW`, `DELETE_ROW`, `INSERT_COLUMN` operations.
    *   *Constraint:* Does not currently support complex re-merging via Markdown; focuses on content updates and simple row/col additions.

### Surgical Logic
*   **Table Cell Isolation**:
    *   When user edits a paragraph provided by `context.document.body.paragraphs`, Word sometimes returns the *entire table* OOXML if the paragraph is in a cell.
    *   The engine detects `<w:tbl>`, finds the target specific paragraph by text match (`originalText`), and performs operations *only* on that node.

*   **Format Removal Workaround**:
    *   Word's `insertOoxml` ignores `<w:rPrChange>` (formatting track changes).
    *   **Workaround**: To "Redline" a format removal (e.g., un-bolding), the engine treats it as a text replacement:
        *   `<w:del><strong>Text</strong></w:del>`
        *   `<w:ins>Text</w:ins>`
    *   This forces Word to show the change visibly as a deletion of the bold text and insertion of plain text.

---

## 3. List Generation Logic

Handled within `pipeline.js` (`executeListGeneration`).

*   **Measurement**: Detects the indentation step of the user's input (2 spaces vs 4 spaces vs tabs) to correctly calculate hierarchy levels.
*   **Numbering Service (`numbering-service.js`)**:
    *   Manages `abstractNum` and `num` definitions.
    *   Generates a virtual `numbering.xml` structure to support custom formats (Legal `1.1.1`, Outline `I. A. 1.`, etc.).
    *   Matches Markdown markers (`1.`, `-`, `A.`) to Word's internal numbering ID system.

## 4. Helper Modules

*   **`numbering-service.js`**: Manages `numId` and `abstractNumId` generation to create valid Word lists on the fly.
*   **`types.js`**: Defines the `RunModel`, `DiffOp` enums, and XML Namespaces (`NS_W`).
*   **`integration.js`**: Bridges the pure logic with the Office.js environment (helper functions for checking availability).

## 5. Critical Data Structures

### RunModel Reference
```typescript
interface RunEntry {
    kind: RunKind;             // TEXT, PARAGRAPH_START, CONTAINER_START...
    text: string;              // The plain text content
    startOffset: number;       // Start position in the accepted text
    endOffset: number;         // End position
    rPrXml?: string;           // Serialized <w:rPr> element
    containerContext?: string; // ID of the parent container
    containerId?: string;      // ID if this IS a container start
    nodeXml?: string;          // Preserved XML for opaque nodes (Bookmarks)
}
```

### Format Hints
Overlay objects derived from Markdown.
```javascript
{
  start: 0,
  end: 5,
  format: { bold: true, italic: false }
}
```
