# Project Overview & Architecture Context

## GSD Documentation Contract

| File | Role |
|------|------|
| `SPEC.md` | Project vision and operating principles (**always loaded**). |
| `ARCHITECTURE.md` | System understanding and technical design map. |
| `ROADMAP.md` | Direction of travel: what is done, in progress, and next. |
| `STATE.md` | Durable memory across sessions: decisions, blockers, and current position. |
| `PLAN.md` | Atomic execution plan with XML-structured task blocks and verification steps. |
| `SUMMARY.md` | Session outcome log: what happened and what changed. |

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
The engine can be used in Node.js or browsers (using `jsdom` or `xmldom`) without Office.js:
```javascript
import { applyRedlineToOxml } from './modules/reconciliation/standalone.js';

const result = await applyRedlineToOxml(documentXml, originalText, modifiedText, {
  author: 'AI Assistant',
  generateRedlines: true
});
// result.oxml contains the updated document XML
```

## Implementation Guidelines
- **Pure OOXML Over Word JS API**: To maintain portability, always prefer a **pure OOXML implementation** (modifying the XML DOM) for document manipulation.
- **Thin Runtime Layers**: Host-specific logic (Word add-in UI, browser demo, MCP server) must stay thin and focus purely on **orchestration, I/O, and environment normalization**. No document-processing logic should live here.
- **Logic Inward**: All new behavior (e.g., list handling, table diffing, property extraction) MUST be implemented in shared reconciliation modules first. If logic starts in `agentic-tools.js` or `demo.js`, it is **migration debt** that must be moved inward to the core engine.
- **Future Goal: Complete Migration Inward**: The long-term technical debt reduction strategy is to move all document-shaping logic (like custom markdown preprocessing or complex structural heuristics currently in `agentic-tools.js`) into the reconciliation engine. This ensures that the Browser Demo and MCP Server eventually reach 100% feature parity with the Word Add-in.
- **Avoid Word JS API**: Strictly avoid using the Word JS API (`Range.text = ...`, `Paragraph.insertHtml(...)`) unless there is no possible OOXML equivalent for the required UI/selection interaction.
- **Consult User First**: If a solution seems to require the Word JS API, **you must consult with the user first** with a justification for why OOXML cannot be used.

## Core Entry Points & Responsibilities

| File | Responsibility | Host Environment |
|------|----------------|------------------|
| **`standalone.js`** | Public facade for the **Reconciliation Engine**. Handles normalization of XML providers and provides structural list fallbacks. | **Any** (Browser, Node, MCP) |
| **`agentic-tools.js`** | Orchestrator for the **Word Add-in**. Handles Word-specific state, batching, and coordination between AI results and the document. | **Word Only** (Office.js) |

> [!IMPORTANT]
> **`agentic-tools.js` should remains thin**. Its primary job is to parse AI JSON, loop through paragraphs, and call the engine. Any logic that touches the *content* of a paragraph (how it is edited, how lists are formed) belongs in the engine via `standalone.js`.

## Essential Documentation
- [ARCHITECTURE.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/ARCHITECTURE.md): Deep dive into system components and the reconciliation engine.
- [ROADMAP.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/ROADMAP.md): Migration status and future goals.
- [STATE.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/STATE.md): Current technical posture and architectural decisions.
