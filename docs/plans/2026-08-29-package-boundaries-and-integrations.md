# Package Boundaries, Integrations, and Web Portability Plan (v0.8.0 Architecture)

**Date:** 2026-08-29  
**Last Updated:** 2026-09-29 (Upgrade target v0.8.1; v0.8.0 universal facade)

**Status:** Active — Ready for Execution  
**Prerequisite Baseline:** **The Direct Upgrade Plan (`2026-09-09-docx-redline-js-v0.5.4-upgrade.md`) is COMPLETED.** Both the root add-in and `mcp/docx-server` are pinned to exact `@ansonlai/docx-redline-js@0.8.1`.

---

## 1. Guiding Principles & Priority Hierarchy

This plan is governed by three strict priorities, ordered from most important to least important:

1. **Accuracy when making changes (Highest Priority):**  
   Every integration boundary must guarantee byte-for-byte fidelity, schema validity, intact relationship IDs, and deterministic rollback. The core redline engine must validate all mutations before committing. No host boundary may swallow or alter error codes or receipts.
2. **Preserving OOXML-only approach (Second Priority):**  
   **The core domain functions must have zero knowledge of Microsoft Word, Office.js, or browser DOM.** All core services must operate purely on standard OOXML or canonical `Uint8Array` DOCX packages. The Word Add-In is treated as a temporary, replaceable wrapper. The entire system must be architected so that ditching MS Word in favor of a standalone web-based document editor requires discarding only the thin Word I/O shell, leaving 100% of the core editing and agentic functions intact.
3. **Speed and performance (Third Priority):**  
   Single-hop I/O at host boundaries; single-DOM document sessions inside the core engine.

---

## 2. Universal Architecture: The Host Shell & Core Engine

The repository enforces a strict separation between **Disposable Host Shells** and the **Portable OOXML Core Engine**:

```
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│                                Disposable Host Shells                                   │
│                                                                                         │
│   [Word Add-in Shell]          [Browser Demo Shell]    [Local MCP Shell]  [Future Web]  │
│   src/taskpane/                browser-demo/           mcp/docx-server/   (Standalone)  │
│   - Office.js I/O              - Textarea / Monaco     - JSON-RPC stdio   - Canvas/DOM  │
│   - context.sync() I/O only    - File upload / download- openDocx()       - Web-native  │
│   - Reads/writes OOXML         - Pure browser events   - Direct fs I/O    - Pure OOXML  │
└────────────────────────────────────────────┬────────────────────────────────────────────┘
                                             │
                                             │ Canonical Operations & Uint8Array / OOXML
                                             │
┌────────────────────────────────────────────▼────────────────────────────────────────────┐
│                       Portable OOXML Core Engine (v0.8.0 Facade)                        │
│                       src/taskpane/modules/docx-redline-js-integration/                 │
│                                                                                         │
│   Universal Functions:                                                                  │
│   - openDocx(uint8Array) &rarr; doc.inspect(), doc.applyOperations(), doc.toUint8Array()│
│   - executeOoxmlOperation(ooxmlString, operation, options)                             │
│   - executePureOoxmlBatch(context, targetScope, operations, options)                   │
│                                                                                         │
│   * 100% Host-Agnostic (Zero Office.js, Zero Node built-ins, Zero DOM dependencies)    │
│   * Built-in fflate ZIP container (No external JSZip needed anywhere)                  │
│   * Runs identically in Node.js, Web Worker, Modern Browser, or Word Taskpane          │
└────────────────────────────────────────────┬────────────────────────────────────────────┘
                                             │
┌────────────────────────────────────────────▼────────────────────────────────────────────┐
│                       Underlying Engine (@ansonlai/docx-redline-js@0.8.1)               │
│                                                                                         │
│   - Caller-order-independent batching with atomic rollback                              │
│   - Localized exact replacements (replacements: [{ find, replace }])                   │
│   - Native numbering.xml, comments.xml, and relationship management                     │
│   - Baseline-delta validation & commit-aware mutation receipts                          │
└─────────────────────────────────────────────────────────────────────────────────────────┘
```

### Future Porting Blueprint (Ditching MS Word Entirely)
If Microsoft Word is abandoned tomorrow in favor of a web-based document editor:
1. **Discard the Word Shell:** Delete `src/taskpane/modules/docx-redline-js-integration/word-ooxml.js` and `taskpane.html`.
2. **Mount the Web Shell:** In the web editor, load the user's `.docx` file into memory (using standard `ArrayBuffer`), pass it directly to `openDocx(arrayBuffer)`, call `doc.applyOperations()`, and save using `doc.toUint8Array()`. Zero JSZip or XML DOM polyfills needed!
3. **100% Code Reuse:** All agentic tools, prompt builders, redline operation converters, and validation contracts remain unchanged.

---

## 3. Work Packages

### WP1 — Consolidate the Pure OOXML Integration Boundary

**File to Update:** `src/taskpane/modules/docx-redline-js-integration/index.js`

Export clean, host-agnostic entry points:
```javascript
// Re-export standalone package surface
export * from '@ansonlai/docx-redline-js';

// Add-in-only integration bridge exports
export {
    getParagraphOoxmlWithFallback,
    insertOoxmlWithRangeFallback,
    withNativeTrackingDisabled
} from './word-ooxml.js';
export {
    assertRedlineResult,
    INPUT_SANITIZED_WARNING,
    prepareOperationInput,
    RedlineOperationError
} from './redline-result.js';
export {
    executePureOoxmlBatch,
    applyWordOperation,
    applySharedOperationToWordParagraph,
    applySharedOperationToWordScope
} from './word-operation-runner.js';
export { applyRedlineChangesToWordContext } from './word-redline-runner.js';
```

---

### WP2 — Modernize MCP Server to Thin Facade (`mcp/docx-server`)

**Goal:** Refactor `mcp/docx-server` into a thin wrapper around `openDocx`, `doc.inspect()`, and `doc.applyOperations()`. Drop `jszip` and `@xmldom/xmldom`.

**File to Update:** `mcp/docx-server/src/server.mjs`  
**Files to Delete:**  
- `mcp/docx-server/src/services/docx-package-service.mjs`  
- `mcp/docx-server/src/services/paragraph-targeting-service.mjs`  

**Step-by-Step Instructions:**
1. In `mcp/docx-server/package.json`:
   - Set `"@ansonlai/docx-redline-js": "0.8.1"`.
   - Remove `"jszip"` and `"@xmldom/xmldom"`.
2. Refactor `mcp/docx-server/src/server.mjs`:
   ```javascript
   import { openDocx } from '@ansonlai/docx-redline-js';
   import fs from 'node:fs/promises';

   // In-memory session store mapping sessionId -> doc instance
   const sessions = new Map();

   // Tool 1: docx_open
   async function handleDocxOpen({ path, generateRedlines = true }) {
       const bytes = await fs.readFile(path);
       const doc = await openDocx(bytes);
       const sessionId = crypto.randomUUID();
       sessions.set(sessionId, { doc, path, generateRedlines });
       return { sessionId, paragraphCount: doc.inspect().paragraphs.length };
   }

   // Tool 2: docx_list_paragraphs
   async function handleDocxListParagraphs({ sessionId, start = 0, limit = 500 }) {
       const session = sessions.get(sessionId);
       if (!session) throw new Error('Session not found');
       const paragraphs = session.doc.inspect().paragraphs;
       return {
           total: paragraphs.length,
           items: paragraphs.slice(start, start + limit).map(p => ({
               id: p.paragraphId || `idx:${p.index}`,
               index: p.index,
               ref: p.ref,
               text: p.exactText
           }))
       };
   }

   // Tool 3: docx_edit_paragraph (supporting both full text and localized replacements)
   async function handleDocxEditParagraph({ sessionId, paragraphIndex, replacements, newText, author }) {
       const session = sessions.get(sessionId);
       if (!session) throw new Error('Session not found');
       const op = {
           type: 'redline',
           target: { paragraphIndex },
           author: author || 'AI Assistant',
           ...(replacements ? { replacements } : { modified: newText })
       };
       const result = await session.doc.applyOperations([op], { atomic: true });
       return { success: result.status === 'success', receipts: result.receipts };
   }

   // Tool 4: docx_save
   async function handleDocxSave({ sessionId, outputPath }) {
       const session = sessions.get(sessionId);
       if (!session) throw new Error('Session not found');
       const targetPath = outputPath || session.path;
       await fs.writeFile(targetPath, session.doc.toUint8Array());
       return { success: true, savedPath: targetPath };
   }
   ```
3. Run `node tests/mcp_docx_redline_service_tests.mjs` and verify all tests pass against the new thin server.
