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

---

## 1. Taskpane Layer (`src/taskpane/taskpane.js`)

The `taskpane.js` file is the entry point and main controller. It manages the UI, chat state, and the primary "Think-Act" loop.

### Key Responsibilities
1.  **State Management**:
    *   `chatHistory`: Maintains the conversation context, trimmed to a rolling window (default 10 turns).
    *   `currentRequestController`: Handles cancellation (AbortController) for long-running requests.
    *   `toolsExecutedInCurrentRequest`: Tracks success for partial failure recovery.

2.  **Context Extraction (`extractEnhancedDocumentContext`)**:
    *   Before every API call, the system "reads" the document.
    *   **Enhanced Notation**: It converts Word paragraphs into a metadata-rich textual format for the LLM:
        *   `[P#|Style] Text...`
        *   `[P#|ListNumber|L:1] Item...` (Captures list structure)
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

This module translates high-level AI instructions (e.g., "Fix the spelling in paragraph 3") into specific Word JS API operations.

### `executeRedline`
The primary tool for modifying document content. It accepts instructions and delegates to specific operations:

| Operation | Description | Implementation Strategy |
| :--- | :--- | :--- |
| **`edit_paragraph`** | Standard text formatting/rewriting. | **Diff-Match-Patch (DMP)**: Computes the difference between old and new text and applies surgical edits if possible, or replaces the paragraph. |
| **`replace_paragraph`** | Replacing content with complex formatting (lists, tables). | **Hybrid Router**: Checks content type.<br>1. **Lists/Tables**: Routes to **OOXML Pipeline** to preserve structure and track changes.<br>2. **Plain/Simple HTML**: Falls back to `insertHtml` for speed. |
| **`replace_range`** | modifying multiple paragraphs. | **Expansion**: Expands selection from start P# to end P#.<br>**Logic**: Detects if inside a table (replaces table if needed) or just text. |
| **`modify_text`** | Surgical replacement of a specific substring. | **Search**: Uses `paragraph.search(originalText)`.<br>**Range Expansion**: If text > 80 chars, uses a prefix/suffix search strategy to locate the range. |

### Other Tools
*   **`insert_comment`**: Adds comments to specific paragraphs.
*   **`highlight_text`**: altering background color.
*   **`edit_list` / `edit_table`**: Specialized structural editing tools that bypass generic text processing for higher reliability.

---

## 3. Markdown & Conversion Logic

The system allows the AI to write in Markdown, which is then converted to Word-native formats.

### Markdown Processor (`markdown-processor.js`)
*   **Purpose**: Pre-processing for the **OOXML Engine**.
*   **Logic**:
    1.  **Parse**: Regex-based parsing of Markdown symbols (`**bold**`, `*italic*`, `~~strike~~`, `++underline++`).
    2.  **Strip**: Returns `cleanText` (plain text without markers).
    3.  **Hints**: Returns an array of `FormatHints` (`{ start, end, format: { bold: true } }`).
*   **Usage**: The OOXML engine generates the plain text XML, then effectively "paints" these format hints onto the run properties (`<w:rPr>`).

### Markdown Utils (`markdown-utils.js`)
*   **Purpose**: Pre-processing for the **HTML Fallback** path.
*   **Logic**: Uses established libraries (likely `marked` or custom parsers) to convert Markdown string -> HTML string (e.g., `**text**` -> `<strong>text</strong>`).
*   **Usage**: Used when `insertHtml` is safe/preferred (simple text updates without complex nesting or track changes requirements).

## 4. Checkpoint System

*   **Location**: `taskpane.js` (`createCheckpoint`, `restoreCheckpoint`).
*   **Mechanism**: Stores the **Entire Document Body OOXML** in `localStorage`.
*   **Trigger**: Automatically triggered before *every* destructive tool call (`apply_redlines`, etc.).
*   **Management**: Implements Quota Management (prunes old checkpoints when 5MB limit is reached).

## 5. Critical Data Flows

### The Context Loop
1.  **Read**: `Word.run` -> `extractEnhancedDocumentContext` -> `[P1] Text...`
2.  **Think**: Gemini receives context -> Decides to call `apply_redlines(instruction: "fix typo in P1")`.
3.  **Act**: `agentic-tools.js` parses instruction -> Locates Paragraph 1 -> Applies Change.
4.  **Verify**: Loop repeats, updated context effectively verifies the change (AI sees the new text in next turn).

### The Redline/formatted Text Flow
1.  **Input**: AI generates `newContent` with Markdown (e.g., `Hello **World**`).
2.  **Detection**: `agentic-tools` detects Markdown markers.
3.  **Path Selection**:
    *   *Complex (Nested Lists)?* -> **OOXML Pipeline** -> `markdown-processor` extracts hints -> Builds `<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>World</w:t></w:r></w:p>`
    *   *Simple?* -> **HTML Fallback** -> `markdown-utils` -> `Hello <strong>World</strong>` -> `paragraph.insertHtml()`.
