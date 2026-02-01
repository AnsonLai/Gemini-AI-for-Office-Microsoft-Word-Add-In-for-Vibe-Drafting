# Architecture Documentation

This document outlines the architectural logic of the **Gemini AI for Office Word Add-in**, focusing on the Taskpane, Chat Orchestration, and Command Execution layers.

## High-Level Architecture

The application follows a **Chat-Centric** architecture where the Taskpane acts as the orchestrator between the User, the Word Document (via Office.js), and the Gemini API.

```mermaid
graph TD
    User[User Input] --> Taskpane[Taskpane.js]
    Taskpane --> Context[Context Extraction]
    Context --> Word[Word Document]
    Taskpane --> Gemini[Gemini API]
    Gemini --> Tools[Tool Definitions (JSON)]
    Gemini --Function Call--> Taskpane
    Taskpane --> Agentic[Agentic Tools]
    Agentic --> Logic[Decision Logic]
    Logic --Simple Text--> HTML[HTML Insertion]
    Logic --Complex/List--> OOXML[OOXML Engine]
    Logic --Surgical--> Search[Search & Replace]
    HTML --> Word
    OOXML --> Word
    Search --> Word
```

## Current State: Hybrid Architecture with Migration Path

The system currently operates in a **hybrid mode**, using both Word JS API and OOXML approaches. The goal is to migrate to a **pure OOXML approach** for better portability and consistency.

### Word JS API Dependencies (Current)

The following areas still rely on Word JS API:

1. **Context Extraction**: `extractEnhancedDocumentContext()` uses `Word.run()` and `paragraph.load()`
2. **Document Navigation**: `executeNavigate()` uses `paragraph.select()`
3. **Comment Operations**: `executeComment()` uses `match.insertComment()`
4. **Track Changes Management**: `setChangeTrackingForAi()` and `restoreChangeTracking()`
5. **List Operations**: `executeConvertHeadersToList()` uses Word's list API
6. **Table Operations**: `executeEditTable()` uses Word's table API
7. **Search Operations**: `searchWithFallback()` uses `paragraph.search()`

### OOXML Capabilities (Current)

The following areas already use pure OOXML:

1. **Text Editing**: `applyRedlineToOxml()` for paragraph-level edits
2. **Highlighting**: `applyHighlightToOoxml()` for surgical highlighting
3. **List Generation**: OOXML pipeline for complex list structures
4. **Table Generation**: OOXML pipeline for table creation
5. **Checkpoint System**: Stores entire document body as OOXML

---

## 1. Taskpane Layer (`src/taskpane/taskpane.js`)

The `taskpane.js` file is the entry point and main controller. It manages the UI, chat state, and the primary "Think-Act" loop.

### Key Responsibilities
1.  **State Management**:
    *   `chatHistory`: Maintains the conversation context, trimmed to a rolling window (default 10 turns).
    *   `currentRequestController`: Handles cancellation (AbortController) for long-running requests.
    *   `toolsExecutedInCurrentRequest`: Tracks success for partial failure recovery.

2.  **Context Extraction (`extractEnhancedDocumentContext`)**:
    *   **Current Implementation**: Uses Word JS API (`Word.run()`, `paragraph.load()`)
    *   **Migration Target**: Replace with pure OOXML parsing
    *   Before every API call, the system "reads" the document.
    *   **Reliable Extraction**: Always calls `paragraph.load("text")` before processing. This avoids character truncation issues found in some Word API versions, ensuring the OOXML Engine has a full, accurate string for diffing.
    *   **Enhanced Notation**: It converts Word paragraphs into a metadata-rich textual format for the LLM:
        *   `[P#|Style] Text...`
        *   `[P#|ListNumber|L:level|§] Item...` (Captures list structure)
        *   `[P#|T:row,col] Cell...` (Captures table structure)
    *   **Purpose**: Gives the LLM precise "anchors" (P-numbers) to reference in its tool calls.

3.  **Chat Orchestration (`sendChatMessage`)**:
    *   **Preparation**: Locks UI, extracts context, loads settings.
    *   **The Loop**: A `while(keepLooping)` loop that allows the AI to execute multiple tools in a chain (up to `MAX_LOOPS` = 6).
    *   **Tool Execution**:
        *   Parses `candidate.content.parts` for `functionCall`.
        *   Updates the "Thinking..." UI to show specific actions (e.g., "Applying edits...", "Researching...").
        *   Executes the corresponding function from `agentic-tools.js`.
        *   Pushes the `functionResponse` back to `chatHistory`.
    *   **Recovery Logic**:
        *   Handles generic API errors.
        *   **Specific Recovery**: If Gemini gets confused about turn order (Function Call/Response mismatch), it enters a tiered recovery mode:
            1.  **Tier 1**: Validate/Clean history pairs.
            2.  **Tier 2**: Wipe all function calls from history (keep only text).
            3.  **Tier 3**: Fresh start (wipe history, keep system prompt + user message).
            4.  **Tier 4**: Graceful degradation (stop and show partial success).

---

## 2. Command Layer (`src/taskpane/modules/commands/agentic-tools.js`)

This module translates high-level AI instructions (e.g., "Fix the spelling in paragraph 3") into specific document operations.

### Current Hybrid Implementation

The system currently uses a hybrid approach with both Word JS API and OOXML:

#### Word JS API Operations (To be Migrated)

1. **`executeComment()`**: Uses `match.insertComment()` for comment insertion
2. **`executeNavigate()`**: Uses `paragraph.select()` for navigation
3. **`executeConvertHeadersToList()`**: Uses Word's list API (`startNewList()`, `attachToList()`)
4. **`executeEditTable()`**: Uses Word's table API for table operations
5. **`searchWithFallback()`**: Uses `paragraph.search()` for text search

#### OOXML Operations (Already Portable)

1. **`executeRedline()`**: Uses `applyRedlineToOxml()` for text editing
2. **`executeHighlight()`**: Uses `applyHighlightToOoxml()` for highlighting
3. **`executeEditList()`**: Uses OOXML pipeline for list generation
4. **`executeInsertListItem()`**: Uses OOXML for surgical list item insertion

### Migration Strategy

The goal is to replace all Word JS API calls with pure OOXML operations:

1. **Comment Operations**: Replace `insertComment()` with OOXML comment injection
2. **Navigation**: Replace `paragraph.select()` with OOXML-based position tracking
3. **List Conversion**: Replace Word list API with OOXML list generation
4. **Table Operations**: Replace Word table API with OOXML table manipulation
5. **Search Operations**: Replace `paragraph.search()` with OOXML text parsing

---

## 3. Markdown & Conversion Logic

The system allows the AI to write in Markdown, which is then converted to Word-native formats.

### Markdown Processor (`markdown-processor.js`)
*   **Purpose**: Pre-processing for the **OOXML Engine**.
*   **Logic**:
    1.  **Parse**: Regex-based parsing of Markdown symbols (`**bold**`, `*italic*`, `~~strike~~`, `++underline++`).
    2.  **Strip**: Returns `cleanText` (plain text without markers).
    3.  **Hints**: Returns an array of `FormatHints` (`{ start, end, format: { bold: true } }`).
*   **Usage**: The OOXML engine generates the plain text XML, then "paints" these format hints using an **Elementary Segment Splitting** strategy. This ensures that overlapping formats (e.g., ***Bold+Italic***) are handled correctly.

### Markdown Utils (`markdown-utils.js`)
*   **Purpose**: Pre-processing for the **HTML Fallback** path.
*   **Logic**: Uses established libraries (likely `marked` or custom parsers) to convert Markdown string -> HTML string (e.g., `**text**` -> `<strong>text</strong>`).
*   **Usage**: Used when `insertHtml` is safe/preferred (simple text updates without complex nesting or track changes requirements).
*   **Migration Note**: This HTML fallback should be replaced with pure OOXML generation for consistency.

## 4. Checkpoint System

*   **Location**: `taskpane.js` (`createCheckpoint`, `restoreCheckpoint`).
*   **Mechanism**: Stores the **Entire Document Body OOXML** in `localStorage`.
*   **Trigger**: Automatically triggered before *every* destructive tool call (`apply_redlines`, etc.).
*   **Management**: Implements Quota Management (prunes old checkpoints when 5MB limit is reached).
*   **Portability**: This system is already OOXML-based and portable.

## 5. Critical Data Flows

### The Context Loop
1.  **Read**: `Word.run` -> `extractEnhancedDocumentContext` -> `[P1] Text...`
   - **Migration Target**: Replace Word JS API with pure OOXML parsing
2.  **Think**: Gemini receives context -> Decides to call `apply_redlines(instruction: "fix typo in P1")`.
3.  **Act**: `agentic-tools.js` parses instruction -> Locates Paragraph 1 -> Applies Change.
4.  **Verify**: Loop repeats, updated context effectively verifies the change (AI sees the new text in next turn).

### The Redline/formatted Text Flow
1.  **Input**: AI generates `newContent` with Markdown (e.g., `Hello **World**`).
2.  **Detection**: `agentic-tools` detects Markdown markers.
3.  **Path Selection**:
    *   *Complex (Nested Lists/Direct Bold)?* -> **OOXML Pipeline** -> `markdown-processor` extracts hints -> **Reconstruction Mode** splits the text into pieces where format changes, applying explicit `w:val="1"` attributes.
    *   *Simple?* -> **HTML Fallback** -> `markdown-utils` -> `Hello <strong>World</strong>` -> `paragraph.insertHtml()`.
    *   **Migration Target**: Replace HTML fallback with pure OOXML generation

## 6. Migration Plan to Pure OOXML

### Phase 1: Replace Context Extraction
- **Current**: Uses Word JS API (`Word.run()`, `paragraph.load()`)
- **Target**: Parse OOXML directly to extract paragraph information
- **Benefit**: Eliminates dependency on Word API for document reading

### Phase 2: Replace Comment Operations
- **Current**: Uses `match.insertComment()`
- **Target**: Inject comments directly into OOXML structure
- **Benefit**: Portable comment functionality

### Phase 3: Replace Navigation
- **Current**: Uses `paragraph.select()`
- **Target**: Implement OOXML-based position tracking and scrolling
- **Benefit**: Portable navigation across different Word environments

### Phase 4: Replace List Conversion
- **Current**: Uses Word's list API
- **Target**: Use OOXML list generation for all list operations
- **Benefit**: Consistent list handling across environments

### Phase 5: Replace Table Operations
- **Current**: Uses Word's table API
- **Target**: Use OOXML table manipulation for all table operations
- **Benefit**: Portable table functionality

### Phase 6: Replace Search Operations
- **Current**: Uses `paragraph.search()`
- **Target**: Implement OOXML text parsing and search
- **Benefit**: Portable text search functionality

### Phase 7: Replace HTML Fallback
- **Current**: Uses `markdown-utils` and `insertHtml`
- **Target**: Use pure OOXML generation for all text formatting
- **Benefit**: Consistent formatting across all environments

## 7. Benefits of Pure OOXML Approach

1. **Portability**: Code can run outside Word add-in environment
2. **Consistency**: Same behavior across different Word versions
3. **Maintainability**: Single approach for all document operations
4. **Testability**: Easier to test with mock OOXML documents
5. **Future-proof**: Independent of Word JS API changes

## 8. Implementation Guidelines

### For New Features
1. **Always use OOXML first** for any document manipulation
2. **Avoid Word JS API** unless absolutely necessary
3. **Document exceptions** clearly in architecture files
4. **Create migration tickets** for any Word JS API usage

### For Existing Features
1. **Prioritize migration** of Word JS API dependencies
2. **Test thoroughly** after each migration
3. **Update documentation** to reflect new OOXML approach
4. **Remove deprecated code** after successful migration