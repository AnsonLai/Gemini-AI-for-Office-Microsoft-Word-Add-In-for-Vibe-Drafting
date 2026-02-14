# Project Overview & Architecture Context

## Project Purpose
This is a **Microsoft Word add-in** that integrates **Gemini AI**. It provides Gemini with **agentic tools** to directly edit and markup the document within Word.

## Core Architectural Principle: Decoupling & Portability
The architecture is designed to be **host-agnostic**. The core reconciliation engine is decoupled from any specific host (like Word), enabling a multi-project ecosystem:
- **`core/`**: The host-independent reconciliation implementation (currently in `src/taskpane/modules/reconciliation`).
- **`word-addin/`**: The Microsoft Word-specific implementation (currently in `src/taskpane`).
- **`browser-demo/`**: A standalone web demonstration.
- **`mcp/`**: Model Context Protocol servers for document automation.

> [!NOTE]
> **Splitting Vision**: While these currently reside in a single repository for ease of cross-module development, the long-term goal is to split these into independent projects once the core engine reaches maturity.
The engine can be used in Node.js or browsers without Office.js:
```javascript
import { applyRedlineToOxml } from './modules/reconciliation/standalone.js';

const result = await applyRedlineToOxml(documentXml, originalText, modifiedText, {
  author: 'AI Assistant',
  generateRedlines: true
});
// result.oxml contains the updated document XML
```

## Implementation Guidelines
- **Pure OOXML Over Word JS API**: To maintain portability, always prefer a **pure OOXML implementation** for document manipulation.
- **Thin Runtime Layers**: Host-specific logic (Word add-in UI/runtime, browser demo, MCP server) should stay thin and focus on **orchestration, I/O, and environment glue**.
- **Logic Inward**: New behavior should be implemented in shared reconciliation modules first. If logic starts outside the engine, it is considered "migration debt" and must be moved inward to the core when possible.
- **Avoid Word JS API**: Strictly avoid using the Word JS API unless there is no other way to achieve the required functionality (e.g., UI interactions).
- **Consult User First**: If a solution seems to require the Word JS API, **you must consult with the user first** before proceeding, as this is contrary to the portability goal.

## Essential Documentation
- [ARCHITECTURE.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/ARCHITECTURE.md): Deep dive into system components and the reconciliation engine.
- [ROADMAP.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/ROADMAP.md): Migration status and future goals.
- [STATE.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/STATE.md): Current technical posture and architectural decisions.
