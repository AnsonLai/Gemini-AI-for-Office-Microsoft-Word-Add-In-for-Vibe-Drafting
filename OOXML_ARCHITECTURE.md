# OOXML Reconciliation Architecture

This document outlines the architecture of the **OOXML Reconciliation Engine**, located in `src/taskpane/modules/reconciliation/`. This system is responsible for converting AI-generated Markdown text into valid Office Open XML (Word) structures while preserving existing document formatting and generating precise "Track Changes" (Redlines).

## High-Level Architecture

The system operates in two distinct modes depending on the complexity of the operation:

1.  **Reconstruction Pipeline** (Standard): Parses original OOXML, diffs against new text, and rebuilds the paragraph from scratch. Used for standard text edits and list generation.
2.  **Surgical Hybrid Engine** (Complex): Direct DOM manipulation. Used for Tables (to preserve cell structure) and **Pure Formatting Changes** (Bold, Italic, etc.) to ensure clean redlines.

```mermaid
graph TD
    Input[OOXML + Markdown Input] --> Router{OxmlEngine Router}
    
    Router --Tables--> Hybrid[Hybrid Engine\n(In-Place DOM)]
    Router --Pure Formatting--> rPrChange[w:rPrChange Engine\n(Property DOM)]
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

    subgraph "w:rPrChange Engine"
        rPrSnap[Take rPr Snapshot] --> rPrApply[Apply Formatting]
        rPrApply --> rPrWrap[Inject w:rPrChange]
    end

    Pipeline --> Output[Final OOXML]
    Hybrid --> Output
    rPrChange --> Output
```

## Current OOXML Capabilities

The OOXML engine currently handles the following operations without Word JS API dependencies:

### ✅ Fully OOXML-Based Operations

1. **Text Editing**: `applyRedlineToOxml()` for paragraph-level text modifications
2. **Highlighting**: `applyHighlightToOoxml()` for surgical highlighting with redline support
3. **List Generation**: Complex nested lists with custom numbering styles
4. **List Range Editing (`executeEditList`)**: Multi-paragraph list replacement via reconciliation-generated OOXML, including `w:ins`/`w:del` when redlines are enabled
5. **Table Generation**: Complete table creation and modification
6. **Format Preservation**: Maintains existing document formatting during edits
7. **Pure Formatting Redlines**: Generates `w:rPrChange` elements for clean formatting-only changes (Bold, Italic, U, Strike)
8. **Track Changes**: Generates proper `w:ins`/`w:del` elements for text edits
9. **Comment Preservation**: Maintains comment positions during text edits

### 🚧 Hybrid Operations (Partial OOXML)

1. **List Conversion**: `executeConvertHeadersToList()` still uses Word list API
2. **Table Editing**: `executeEditTable()` uses Word table API for some operations
3. **Navigation**: `executeNavigate()` uses Word selection API
4. **Comments**: `executeComment()` uses Word comment API

### 🎯 Migration Targets

The following areas need to be migrated from Word JS API to pure OOXML:

1. **List Conversion**: Replace Word list API with OOXML list generation
2. **Table Editing**: Replace Word table API with OOXML table manipulation
3. **Navigation**: Implement OOXML-based position tracking
4. **Comments**: Implement OOXML comment injection

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

### Stage 1.5: Tracked Characters & Sentinels
*   **Position Continuity**: Formatting offsets are tracked via `currentInsertOffset`.
*   **Drift Prevention**: The engine correctly increments this offset for kept (EQUAL) text but skips it for DELETED text. This prevents "formatting drift" where bolding would shift to random parts of the paragraph after an edit.

### Stage 2: Markdown Pre-processing (`markdown-processor.js`)
*   Strips Markdown symbols (`**`, `__`) from the new text to create a "Clean" string for comparison.
*   Generates **Format Hints**: An overlay of formatting instructions derived from the Markdown.
*   **Source of Truth**: In Reconstruction Mode, these hints are treated as the definitive state. Existing formatting is synchronized to match them (allowing both bolding and unbolding).
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

**Serialization strategy: Elementary Segment Splitting**
*   To support nested or overlapping formatting, the engine splits text into "elementary segments" at every hint boundary.
*   Each segment is then assigned the union of all overlapping formatting hints.
*   **Explicit Attributes**: Uses `w:val="1"` (for bold/italic) which ensures Word applies the format explicitly rather than relying on implicit toggles.

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

*   **Highlight Injection**:
    *   The `applyHighlightToOoxml` function (in `ooxml-formatting-removal.js`) performs targeted XML modification to inject `w:highlight` tags into specific runs.
    *   This bypasses the reconstruction pipeline for high-speed, structural-preserving formatting updates.

*   **Pure Formatting Changes**:
    *   To ensure formatting changes (e.g., adding or removing Bold) appear as native "Formatting" edits in Word rather than "Delete + Insert" redlines, the engine uses **Surgical Property Modification**.
    *   **Mechanism**: Directly modifies the `<w:rPr>` element to set the new state (e.g., `<w:b w:val="1"/>`) and injects a `<w:rPrChange>` element as a child of `<w:rPr>`.
    *   **w:rPrChange Content**: This element contains a "snapshot" of the `<w:rPr>` state *before* the change. Word uses this to show the "Formatted: Bold" comment in the margin.
    *   **High Fidelity**: This path bypasses the standard reconstruction pipeline to ensure that non-changed properties (like complex fonts or spacing) are preserved with 100% accuracy.
    *   **Functions**: `applyFormatAdditionsAsSurgicalReplacement` and `applyFormatRemovalAsSurgicalReplacement` in `oxml-engine.js`.

### Comment & Range Preservation
*   **Position Markers**: Elements like `w:commentRangeStart/End` are treated as zero-width position markers during ingestion.
*   **Offset Alignment**: They are stored in a `sentinelMap` without adding character sentinels to the primary comparison string. 
*   **Re-insertion**: During reconstruction, the engine checks the `baseIndex` at every step and re-inserts these markers at their exact original offsets, ensuring comments stay pinned to the correct text even after edits.

---

## 3. List Generation Logic

Handled within `pipeline.js` (`executeListGeneration`).

*   **Measurement**: Detects the indentation step of the user's input (2 spaces vs 4 spaces vs tabs) to correctly calculate hierarchy levels.
*   **Numbering Service (`numbering-service.js`):**
    *   Manages `abstractNum` and `num` definitions.
    *   Generates a virtual `numbering.xml` structure to support custom formats (Legal `1.1.1`, Outline `I. A. 1.`, etc.).
    *   Matches Markdown markers (`1.`, `-`, `A.`) to Word's internal numbering ID system.
*   **Runtime insertion behavior**: For list tools that inject prebuilt OOXML (`executeEditList`), native Word track changes is temporarily disabled during insertion, because redline markup is already embedded in OOXML (`w:ins`/`w:del`).

## 4. Helper Modules

*   **`numbering-service.js`**: Manages `numId` and `abstractNumId` generation to create valid Word lists on the fly.
*   **`types.js`**: Defines the `RunModel`, `DiffOp` enums, and XML Namespaces (`NS_W`).
*   **`integration.js`**: Bridges the pure logic with the Office.js environment (helper functions for checking availability).
*   **`ooxml-formatting-removal.js`**: Provides surgical utilities for adding/removing specific formatting tags from OOXML runs (e.g., highlights).

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

---

## Migration Status and Roadmap

### ✅ Completed Migrations

1. **Text Editing**: Fully migrated to OOXML pipeline
2. **Highlighting**: Fully migrated to OOXML surgical engine
3. **List Generation**: Fully migrated to OOXML pipeline
4. **Table Generation**: Fully migrated to OOXML pipeline
5. **Checkpoint System**: Always used OOXML

### 🚧 In Progress

1. **List Conversion**: `executeConvertHeadersToList()` still uses Word list API for complex multi-paragraph conversion.
2. **Table Editing**: `executeEditTable()` uses Word table API for advanced row/column insertion in existing tables.
3. **Comment Operations**: `executeComment()` still uses Word selection-based search and `insertComment()`.

### 📋 Planned Migrations

1. **Comment Operations**: Replace Word comment API with OOXML comment injection
2. **Navigation**: Replace Word selection API with OOXML position tracking
3. **Search Operations**: Replace Word search API with OOXML text parsing

### Migration Strategy

1. **Identify Word API Dependencies**: Audit all code for Word JS API calls
2. **Implement OOXML Alternatives**: Create pure OOXML implementations
3. **Test Thoroughly**: Ensure OOXML approach works across different document structures
4. **Replace and Remove**: Replace Word API calls with OOXML calls, remove deprecated code
5. **Update Documentation**: Keep architecture docs current with migration status

### Benefits of Pure OOXML Approach

1. **Portability**: Code can run outside Word add-in environment
2. **Consistency**: Same behavior across different Word versions
3. **Maintainability**: Single approach for all document operations
4. **Testability**: Easier to test with mock OOXML documents
5. **Future-proof**: Independent of Word JS API changes

### Implementation Guidelines

1. **Always use OOXML first** for any new document manipulation features
2. **Avoid Word JS API** unless absolutely necessary and document exceptions
3. **Create migration tickets** for any remaining Word JS API usage
4. **Test with complex documents** to ensure OOXML approach handles edge cases
5. **Update architecture documentation** to reflect migration progress
