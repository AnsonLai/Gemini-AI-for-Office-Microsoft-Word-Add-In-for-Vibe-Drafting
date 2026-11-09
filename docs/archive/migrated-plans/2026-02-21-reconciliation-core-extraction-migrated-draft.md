# Reconciliation Core Package Extraction Plan

> **Superseded on 2026-08-29:** This duplicate plan is covered by the completed
> `completed/2026-02-21-reconciliation-core-extraction.md`. No open work was
> carried forward from this document.

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Prepare the OOXML reconciliation core (`src/taskpane/modules/reconciliation/`) for extraction into its own standalone repository. Clean all host-specific (Office.js/Word JS API) and AI-specific (`'Gemini AI'`) contamination from the core, refactor the bloated `standalone.js` entry point, strengthen isolation tests, and set up package infrastructure.

**Guiding Principle:** The core is purely for `.docx` XML structure manipulation and markdown conversions. Nothing Word JS API or AI-related. Callers configure their own author name, platform, XML provider, and logger.

**Tech Stack:** Pure ES modules, `diff-match-patch` (only npm dependency), optional `@xmldom/xmldom` peer dependency for Node.js environments.

---

## Current Architecture Context

### Project Structure
```
AIWordPlugin/
├── src/taskpane/modules/reconciliation/   ← THE CORE (59 files)
│   ├── adapters/          (logger.js, xml-adapter.js)
│   ├── core/              (types, paragraph/list/table targeting, xml-query)
│   ├── engine/            (oxml-engine router, surgical/reconstruction modes, format ops)
│   ├── pipeline/          (5-stage: ingestion → markdown → diff → patch → serialize)
│   ├── services/          (comments, numbering, table-recon, package-builder, standalone plumbing)
│   ├── orchestration/     (route-plan, list-markdown/parsing/fallback, operation converter)
│   ├── integration/       ← WORD-SPECIFIC (stays with add-in, NOT extracted)
│   ├── index.js           ← Word-aware entry point (re-exports integration/)
│   └── standalone.js      ← Host-agnostic entry point (~830 lines, needs refactoring)
├── src/taskpane/ooxml-formatting-removal.js  ← STRAY FILE (belongs in core)
├── browser-demo/demo.js   (consumer: imports from standalone.js)
├── mcp/docx-server/       (consumer: imports from standalone.js)
└── tests/                 (29 test files, mix of core and integration tests)
```

### Dual Entry Points
- **`index.js`**: Exports everything including Word-specific `integration/` helpers. Used by the Word add-in (`agentic-tools.js`).
- **`standalone.js`**: Exports only host-agnostic APIs. Wraps the engine to normalize `useNativeApi` fallbacks. Used by browser-demo and MCP server.

### Consumers and Their Imports

**Word add-in** (`src/taskpane/modules/commands/agentic-tools.js`):
```js
import {
    applyRedlineToOxml, preprocessMarkdown, ReconciliationPipeline,
    wrapInDocumentFragment, getAuthorForTracking, buildListMarkdown,
    normalizeListItemsWithLevels, withNativeTrackingDisabled,
    applySharedOperationToWordParagraph, applySharedOperationToWordScope,
    applyRedlineChangesToWordContext
} from '../reconciliation/index.js';
```

**Browser demo** (`browser-demo/demo.js`):
```js
import {
    configureLogger, getParagraphText, buildTargetReferenceSnapshot,
    findParagraphByBestTextMatch, parseParagraphReference,
    stripLeadingParagraphMarker, splitLeadingParagraphMarker,
    createDynamicNumberingIdState, mergeNumberingXmlBySchemaOrder,
    parseXmlStrictStandalone, getBodyElementFromDocument,
    insertBodyElementBeforeSectPr, normalizeBodySectionOrderStandalone,
    sanitizeNestedParagraphsInTables, ensureNumberingArtifactsInZip,
    ensureCommentsArtifactsInZip, validateDocxPackage
} from '../src/taskpane/modules/reconciliation/standalone.js';
import { applyOperationToDocumentXml } from '../src/taskpane/modules/reconciliation/services/standalone-operation-runner.js';
```

**MCP server** (`mcp/docx-server/src/services/reconciliation-service.mjs`):
```js
import {
    applyRedlineToOxml, configureLogger, configureXmlProvider,
    ingestOoxml, injectCommentsIntoOoxml
} from '../../../../src/taskpane/modules/reconciliation/standalone.js';
```

---

## Issues Found (What Needs Fixing)

### Boundary Violations

1. **`ooxml-formatting-removal.js` lives outside the core** at `src/taskpane/ooxml-formatting-removal.js`. It's pure OOXML manipulation (no Word API) but is imported by `standalone.js:769`. It imports `parseOoxml`/`serializeOoxml` from `./modules/reconciliation/engine/oxml-engine.js`.

2. **`pipeline/pipeline.js:26-29` references `Office.context.platform`** directly:
   ```js
   function detectPlatform() {
       if (typeof Office === 'undefined' || !Office?.context?.platform) {
           return 'Unknown';
       }
       return String(Office.context.platform);
   }
   ```
   This is a host-specific global reference in the core. It fails silently in non-Office environments but shouldn't be there at all.

3. **`'Gemini AI'` hardcoded as default author in ~15 files** across `core/types.js`, `engine/oxml-engine.js`, `engine/format-application.js`, `engine/surgical-mode.js`, `pipeline/serialization.js`, `services/comment-engine.js`, `integration/integration.js`, `integration/word-operation-runner.js`. The canonical default is in `core/types.js:184`:
   ```js
   export function createRevisionMetadata(author = 'Gemini AI') { ... }
   ```

### Structural Issues

4. **`standalone.js` is ~830 lines** and contains substantial business logic that should be in dedicated service modules:
   - Lines 333-747: Numbering ID management (~370 lines) - `createDynamicNumberingIdState`, `reserveNextNumberingId`, `remapNumberingPayloadForDocument`, `mergeNumberingXmlBySchemaOrder`, plus many private helpers
   - Lines 130-216: `resolveParagraphRangeByRefs()` and `inferTableReplacementParagraphBlock()` - pure DOM operations that belong in `core/`

5. **`WORD_MAIN_NS` constant is redeclared** in `standalone.js:489` instead of importing from `core/paragraph-targeting.js` where it's already exported.

### Test Gaps

6. **`no_word_api_standalone_check.mjs` only scans top-level files** in the reconciliation directory. It does NOT recurse into subdirectories (`core/`, `engine/`, `pipeline/`, `services/`, `orchestration/`, `adapters/`). The `Office.context.platform` reference in `pipeline/pipeline.js` is NOT caught by the current test.

7. **No dependency-graph validation** - nothing prevents a future contributor from adding an import to a file outside the reconciliation boundary.

---

## Task 1: Move `ooxml-formatting-removal.js` into the Core

**Files to modify:**
- `src/taskpane/ooxml-formatting-removal.js` → move to `src/taskpane/modules/reconciliation/engine/formatting-removal.js`
- `src/taskpane/modules/reconciliation/standalone.js` (line 769)
- `src/taskpane/modules/reconciliation/index.js` (line ~769)

**Steps:**
1. Create `src/taskpane/modules/reconciliation/engine/formatting-removal.js` with the contents of `src/taskpane/ooxml-formatting-removal.js`
2. In the new file, update the import from `'./modules/reconciliation/engine/oxml-engine.js'` to `'./oxml-engine.js'` (now a sibling in `engine/`)
3. In `standalone.js`, change line 769:
   - FROM: `} from '../../ooxml-formatting-removal.js';`
   - TO: `} from './engine/formatting-removal.js';`
4. In `index.js`, change the corresponding import similarly
5. Search entire repo for any other imports of `ooxml-formatting-removal` and update them
6. Delete `src/taskpane/ooxml-formatting-removal.js`

**Verification:** `node tests/formatting_tests.mjs` and `node tests/standalone_smoke.mjs`

---

## Task 2: Create `adapters/config.js` for Configurable Defaults

**New file:** `src/taskpane/modules/reconciliation/adapters/config.js`

**Purpose:** Centralize configurable defaults (author, platform) that were previously hardcoded or relied on host globals.

**Implementation:**
```js
/**
 * Configurable runtime defaults for the reconciliation core.
 * Callers set these once during bootstrap; all core modules read from here.
 */

let _defaultAuthor = 'Author';
let _platform = 'Unknown';

/** Set the default track-change author for revision metadata. */
export function setDefaultAuthor(author) {
    _defaultAuthor = typeof author === 'string' && author.trim() ? author.trim() : 'Author';
}

/** Get the current default track-change author. */
export function getDefaultAuthor() { return _defaultAuthor; }

/** Set the platform identifier (e.g. 'Win32', 'Mac', 'OfficeOnline'). */
export function setPlatform(platform) {
    _platform = typeof platform === 'string' && platform.trim() ? platform.trim() : 'Unknown';
}

/** Get the current platform identifier. */
export function getPlatform() { return _platform; }
```

**Then export from both entry points:**
- In `standalone.js`: `export { setDefaultAuthor, getDefaultAuthor, setPlatform, getPlatform } from './adapters/config.js';`
- In `index.js`: same

---

## Task 3: Replace All `'Gemini AI'` Defaults with `getDefaultAuthor()`

**Files to modify (every file containing `'Gemini AI'` as a default/fallback):**

| File | Line(s) | Current Pattern | Change To |
|------|---------|-----------------|-----------|
| `core/types.js` | 184 | `author = 'Gemini AI'` | `author` (no default; resolve inside function) |
| `core/types.js` | body of `createRevisionMetadata` | uses `author` param directly | `const resolvedAuthor = author \|\| getDefaultAuthor();` |
| `engine/oxml-engine.js` | 41 | `options.author \|\| 'Gemini AI'` | `options.author \|\| getDefaultAuthor()` |
| `engine/format-application.js` | 85 | `author \|\| 'Gemini AI'` | `author \|\| getDefaultAuthor()` |
| `engine/format-application.js` | 171 | `author \|\| 'Gemini AI'` | `author \|\| getDefaultAuthor()` |
| `engine/surgical-mode.js` | 36 | `author \|\| 'Gemini AI'` | `author \|\| getDefaultAuthor()` |
| `pipeline/serialization.js` | 140, 148, 226, 248 | `'Gemini AI'` | `getDefaultAuthor()` |
| `services/comment-engine.js` | 60 | `author = 'Gemini AI'` | `author` (resolve with `getDefaultAuthor()`) |
| `integration/integration.js` | 45, 53, 215, 218, 220 | `'Gemini AI'` | `getDefaultAuthor()` |
| `integration/word-operation-runner.js` | 137, 243 | `'Gemini AI'` | `getDefaultAuthor()` |

**Each file needs:** `import { getDefaultAuthor } from '../adapters/config.js';` (adjust relative path per depth).

**Note:** JSDoc `@param` annotations that say `@param {string} [options.author='Gemini AI']` should be updated to `@param {string} [options.author]` with a note that it defaults to the configured default author.

**Verification:** Run all tests. Verify track-change author in test output is `'Author'` (the new neutral default) unless explicitly set.

---

## Task 4: Remove `Office.context.platform` from `pipeline/pipeline.js`

**File:** `src/taskpane/modules/reconciliation/pipeline/pipeline.js` (lines 24-30)

**Current:**
```js
function detectPlatform() {
    if (typeof Office === 'undefined' || !Office?.context?.platform) {
        return 'Unknown';
    }
    return String(Office.context.platform);
}
```

**Replace with:**
```js
import { getPlatform } from '../adapters/config.js';
// Remove detectPlatform() function entirely.
// All call sites that used detectPlatform() now call getPlatform() directly.
```

**Then in the Word add-in bootstrap** (wherever `Office.onReady` is called, likely `src/taskpane/taskpane.js`):
```js
import { setPlatform } from './modules/reconciliation/index.js';
// Inside Office.onReady callback:
setPlatform(Office.context.platform);
```

**Verification:** `node tests/standalone_smoke.mjs` — no `Office` reference errors. `node tests/no_word_api_standalone_check.mjs` — should now pass once the test is made recursive (Task 8).

---

## Task 5: Extract Numbering Helpers from `standalone.js` to `services/numbering-helpers.js`

**New file:** `src/taskpane/modules/reconciliation/services/numbering-helpers.js`

**Move these functions from `standalone.js`:**

*Private helpers (not exported by standalone.js but needed internally):*
- `parseIntegerAttribute(element, names)` (line 333)
- `nextAvailableId(startId, occupiedIds, maxPreferred)` (line 344)
- `normalizeNumberingIdState(state)` (line 424)
- `hasXmlParseError(doc)` (line 491)
- `isDirectWordChild(node, localName)` (line 497)
- `insertNumberingNodeInSchemaOrder(root, node, kind)` (line 506)
- `getAttributeFirst(element, names)` (line 525)
- `getElementId(element, names)` (line 533)
- `setElementId(element, preferredName, idValue)` (line 539)
- `setElementVal(element, value)` (line 543)

*Public exports:*
- `createDynamicNumberingIdState(numberingXml, options)` (line 381)
- `reserveNextNumberingId(state, kind)` (line 456)
- `reserveNextNumberingIdPair(state)` (line 481)
- `overwriteParagraphNumIds(paragraphNodes, targetNumId)` (line 553)
- `extractFirstParagraphNumId(paragraphNodes)` (line 569)
- `buildExplicitDecimalMultilevelNumberingXml(numId, abstractNumId, startAt)` (line 588)
- `remapNumberingPayloadForDocument(numberingXml, replacementNodes, numberingIdState)` (line 628)
- `mergeNumberingXmlBySchemaOrder(existingNumberingXml, incomingNumberingXml)` (line 697)

**Dependencies to import in the new file:**
- `import { parseOoxml, serializeOoxml } from '../engine/oxml-engine.js';`
- `import { WORD_MAIN_NS } from '../core/paragraph-targeting.js';` (instead of redeclaring)

**Then in `standalone.js`:** Replace all moved code with re-exports:
```js
export {
    createDynamicNumberingIdState, reserveNextNumberingId, reserveNextNumberingIdPair,
    overwriteParagraphNumIds, extractFirstParagraphNumId,
    buildExplicitDecimalMultilevelNumberingXml, remapNumberingPayloadForDocument,
    mergeNumberingXmlBySchemaOrder
} from './services/numbering-helpers.js';
```

**Also remove** the `WORD_MAIN_NS` redeclaration (line 489) from `standalone.js` — import it from `core/paragraph-targeting.js` where it's needed.

**Update `index.js`** to also re-export these from their new home.

**Verification:** `node tests/standalone_smoke.mjs`, `node tests/list_tests.mjs`, `node tests/standalone_docx_plumbing_tests.mjs`. Also verify browser-demo and MCP server still work (they import these via `standalone.js` re-exports, so paths shouldn't change for them).

---

## Task 6: Extract Paragraph Range & Table Heuristics from `standalone.js` to `core/`

**Move `resolveParagraphRangeByRefs()`** (standalone.js lines 178-216) to `core/paragraph-targeting.js`:
- It already uses `resolveTargetParagraphWithSnapshot` from that file
- Remove the `WORD_MAIN_NS` import since it's already defined there
- Add to the exports list in `core/paragraph-targeting.js`

**Move `inferTableReplacementParagraphBlock()`** (standalone.js lines 130-163) and `isLikelyStructuredTableSourceParagraph()` (lines 109-119) to `core/table-targeting.js`:
- They're table-structure heuristics, logically belong with `synthesizeTableMarkdownFromMultilineCellEdit`
- `inferTableReplacementParagraphBlock` uses `getParagraphText` from `core/paragraph-targeting.js` — add that import

**Then in `standalone.js`:** Replace with re-exports:
```js
export { resolveParagraphRangeByRefs } from './core/paragraph-targeting.js';
export { inferTableReplacementParagraphBlock, isLikelyStructuredTableSourceParagraph } from './core/table-targeting.js';
```

**Update `index.js`** similarly — it currently re-exports these from `standalone.js` (line 103-108); change to import from `core/` directly.

**After this refactor, `standalone.js` should be ~250-300 lines** containing only:
- `applyRedlineToOxml()` wrapper (native-API normalization)
- `reconcileMarkdownTableOoxml()` convenience function
- `applyRedlineToOxmlWithListFallback()` convenience function
- Re-exports from core modules

**Verification:** `node tests/standalone_smoke.mjs`, `node tests/standalone_operation_runner_tests.mjs`

---

## Task 7: Verify `diff-match-patch` and `marked` Dependencies

**Check:** Search for `import` or `require` of `diff-match-patch` within the reconciliation directory.

Expected: Only `pipeline/diff-engine.js` uses it. This is the core's **only npm dependency**.

**Check:** Search for `import` or `require` of `marked` within the reconciliation directory.

Expected: `marked` is NOT used by the reconciliation core. It's used by the browser-demo for markdown→HTML rendering. Confirm this.

**No changes expected** — just verification for the package.json `dependencies` field.

---

## Task 8: Rewrite `no_word_api_standalone_check.mjs` to Be Recursive

**File:** `tests/no_word_api_standalone_check.mjs`

**Problem:** Currently only scans top-level `.js` files in the reconciliation directory. Does NOT recurse into `core/`, `engine/`, `pipeline/`, `services/`, `orchestration/`, `adapters/`.

**Rewrite to:**
1. Recursively find all `.js` files under `reconciliation/`
2. **Exclude** the `integration/` subdirectory entirely (it's Word-specific by design)
3. **Exclude** `index.js` at the reconciliation root (it re-exports integration/)
4. Scan with these forbidden patterns (applied to comment-stripped source):
   - `Office\.` — catches `Office.context.platform` and any other Office globals
   - `Word\.` — catches `Word.run`, `Word.InsertLocation`, etc.
   - `context\.sync` — Word API sync pattern
   - `paragraph\.(getOoxml|insertOoxml)` — Word API OOXML methods
5. Keep existing check: `standalone.js` must NOT import from `integration/`
6. **Add new check:** No file outside `integration/` (and excluding `index.js`) may import from `integration/`
7. **Add new check:** No file outside `integration/` (and excluding `index.js`) may import from outside the `reconciliation/` directory tree (exception: `diff-match-patch` npm package)

**Verification:** Run the rewritten test. It should now catch the `Office.context.platform` reference (which we fixed in Task 4, so it should pass). If run before Task 4, it should fail — confirming the test works.

---

## Task 9: Add Dependency-Graph Validation Test

**New file:** `tests/core_dependency_graph_check.mjs`

**Purpose:** Ensures no file in the core (excluding `integration/` and `index.js`) imports from outside the reconciliation directory tree.

**Implementation:**
1. Recursively find all `.js` files under `reconciliation/` (excluding `integration/`, `index.js`)
2. For each file, extract all `import ... from '...'` statements (and dynamic `import()` calls)
3. Resolve each import path relative to the file
4. Verify the resolved path is either:
   - Within the `reconciliation/` directory tree, OR
   - An npm package name (e.g., `diff-match-patch`) — identified by not starting with `.` or `/`
5. Fail with a clear error message listing the violating file and import

**Verification:** Should pass after all previous tasks are complete. Add a deliberate violation temporarily to confirm it catches it.

---

## Task 10: Package Infrastructure

### 10.1 Create `package.json` for the Core

**New file:** `src/taskpane/modules/reconciliation/package.json`

This enables the reconciliation directory to be treated as a self-contained package during development, and will serve as the basis for the separate repository's `package.json`.

```json
{
  "name": "@gsd/docx-reconciliation",
  "version": "0.1.0",
  "description": "Host-independent OOXML reconciliation engine for .docx manipulation with track changes",
  "type": "module",
  "main": "standalone.js",
  "exports": {
    ".": "./standalone.js",
    "./core/*": "./core/*",
    "./engine/*": "./engine/*",
    "./pipeline/*": "./pipeline/*",
    "./services/*": "./services/*",
    "./orchestration/*": "./orchestration/*",
    "./adapters/*": "./adapters/*"
  },
  "dependencies": {
    "diff-match-patch": "^1.0.5"
  },
  "peerDependencies": {
    "@xmldom/xmldom": ">=0.8.0"
  },
  "peerDependenciesMeta": {
    "@xmldom/xmldom": {
      "optional": true
    }
  },
  "files": [
    "adapters/",
    "core/",
    "engine/",
    "pipeline/",
    "services/",
    "orchestration/",
    "standalone.js",
    "README.md"
  ],
  "keywords": ["docx", "ooxml", "reconciliation", "track-changes", "redlines", "word"],
  "license": "MIT"
}
```

**Notes:**
- Entry point is `standalone.js` (not `index.js` — that stays with the add-in)
- `integration/` and `index.js` are NOT included in `files`
- `@xmldom/xmldom` is optional — browsers have native `DOMParser`
- Sub-path exports allow consumers to import deep modules directly (e.g., `@gsd/docx-reconciliation/services/standalone-operation-runner.js`)

### 10.2 Plan the Import Path Transition

When the core moves to its own repo:

| Consumer | Current Import | Future Import |
|----------|---------------|---------------|
| Word add-in (core functions) | `from '../reconciliation/standalone.js'` | `from '@gsd/docx-reconciliation'` |
| Word add-in (integration/) | `from '../reconciliation/index.js'` | Local `integration/` module in add-in repo |
| Browser demo | `from '../src/.../standalone.js'` | `from '@gsd/docx-reconciliation'` |
| Browser demo (deep) | `from '../src/.../services/standalone-operation-runner.js'` | `from '@gsd/docx-reconciliation/services/standalone-operation-runner.js'` |
| MCP server | `from '../../../../src/.../standalone.js'` | `from '@gsd/docx-reconciliation'` |

The `integration/` directory will be copied into the Word add-in repo as a local adapter layer. Its imports from the reconciliation core will change from relative paths to package imports.

---

## Task 11: Test Classification & Migration Prep

Classify which tests belong with the core package and which stay with the add-in:

| Test File | Belongs To | Reason |
|-----------|-----------|--------|
| `standalone_smoke.mjs` | **Core** | Tests `applyRedlineToOxml` standalone |
| `standalone_ingestion_export_tests.mjs` | **Core** | Tests `ingestWordOoxmlToPlainText/Markdown` |
| `standalone_docx_plumbing_tests.mjs` | **Core** | Tests package plumbing |
| `standalone_operation_runner_tests.mjs` | **Core** | Tests `applyOperationToDocumentXml` |
| `no_word_api_standalone_check.mjs` | **Core** | Isolation boundary test |
| `core_dependency_graph_check.mjs` | **Core** | Dependency validation (new in Task 9) |
| `list_tests.mjs` | **Core** | Tests list reconciliation pipeline |
| `formatting_tests.mjs` | **Core** | Tests formatting engine |
| `table_tests.mjs` | **Core** | Tests table reconciliation |
| `redline_operation_converter_tests.mjs` | **Core** | Tests orchestration converter |
| `comment_tests.mjs` | **Review** | Move if it only tests `services/comment-engine.js` |
| `highlight_tests.mjs` | **Review** | Move if it only tests engine highlight path |
| `integration_tests.mjs` | **Add-in** | Tests Word integration layer |
| `shared_operation_bridge_tests.mjs` | **Review** | Check if it references Word integration |
| `word_operation_runner_adapter_tests.mjs` | **Add-in** | Tests Word adapter |
| `migrated_tool_cutover_tests.mjs` | **Add-in** | Tests command-layer cutover |
| `table_targeting_and_format_flags.mjs` | **Core** | Tests core targeting heuristics |
| `debug_extraction.mjs` | **Neither** | Debug script, not a proper test |
| `debug_test6.mjs` | **Neither** | Debug script |
| `verify_fix.mjs` | **Neither** | One-off verification |

**Action for "Review" items:** Read each file, check its imports. If it only imports from `standalone.js` or core modules (not `integration/`), it goes with the core.

---

## Task 12: Documentation Updates

### 12.1 Update `src/taskpane/modules/reconciliation/ARCHITECTURE.md`

- Remove `integration/` from folder layout and module descriptions (it leaves with the add-in)
- Remove `index.js` references (only `standalone.js` is the package entry)
- Add `adapters/config.js` to module responsibilities
- Add `services/numbering-helpers.js` to module responsibilities
- Add `engine/formatting-removal.js` to module responsibilities
- Update "Public Surfaces" section: only `standalone.js`
- Update "End-to-End Flow" section: remove Word Add-in path

### 12.2 Update root `ARCHITECTURE.md`

- Update the sub-projects diagram to show the core as an external package
- Note that `integration/` now lives in the Word add-in project
- Update import path examples

### 12.3 Update `STATE.md`

- Change "Portability Status" from "80% independent" to "100% independent"
- Update "Migration Debt" to reflect completed work
- Add note about the package extraction

### 12.4 Update `ROADMAP.md`

- Move "Repository Split" from Phase 3 planned to in-progress
- Add the package name and structure

### 12.5 Write package `README.md` (for the core)

Quick sections:
1. **Quick Start**: configure XML provider, apply a redline, read the result
2. **API Overview**: engine, pipeline, services, orchestration
3. **Configuration**: `configureXmlProvider()`, `configureLogger()`, `setDefaultAuthor()`, `setPlatform()`
4. **Hosting Guidance**: Node.js (need `@xmldom/xmldom`) vs browser (native `DOMParser`)
5. **Architecture**: link to internal `ARCHITECTURE.md`

---

## Execution Order

Dependencies between tasks require this order:

```
Task 1  (move formatting-removal.js)     ← no dependencies
Task 2  (create adapters/config.js)       ← no dependencies
Task 3  (replace 'Gemini AI' defaults)    ← depends on Task 2
Task 4  (remove Office.context.platform)  ← depends on Task 2
Task 5  (extract numbering helpers)       ← depends on Task 1 (formatting-removal path is settled)
Task 6  (extract range/table heuristics)  ← depends on Task 5 (standalone.js refactoring)
Task 7  (verify dependencies)             ← can run anytime
Task 8  (rewrite isolation test)          ← best after Tasks 1-4 (so it passes)
Task 9  (add dep-graph test)              ← best after Tasks 1-6 (so it passes)
Task 10 (package.json)                    ← after Tasks 1-9 (structure is final)
Task 11 (test classification)             ← after Task 10 (package boundary is defined)
Task 12 (documentation)                   ← last (reflects final state)
```

**Suggested parallel batches:**
1. Tasks 1 + 2 (independent)
2. Tasks 3 + 4 (both depend on Task 2)
3. Task 5
4. Task 6
5. Tasks 7 + 8 + 9 (verification, mostly independent)
6. Tasks 10 + 11
7. Task 12

---

## Full Verification Checklist

After all tasks, run every test:
```bash
# Core tests (should all pass)
node tests/no_word_api_standalone_check.mjs
node tests/core_dependency_graph_check.mjs
node tests/standalone_smoke.mjs
node tests/standalone_ingestion_export_tests.mjs
node tests/standalone_docx_plumbing_tests.mjs
node tests/standalone_operation_runner_tests.mjs
node tests/list_tests.mjs
node tests/formatting_tests.mjs
node tests/table_tests.mjs
node tests/redline_operation_converter_tests.mjs
node tests/comment_tests.mjs
node tests/highlight_tests.mjs
node tests/table_targeting_and_format_flags.mjs

# Integration tests (should still pass — integration/ layer unchanged)
node tests/integration_tests.mjs
node tests/word_operation_runner_adapter_tests.mjs
node tests/shared_operation_bridge_tests.mjs
node tests/migrated_tool_cutover_tests.mjs

# Manual: Open browser-demo/demo.html and verify chat mode works
# Manual: Run MCP server (npm run mcp:docx) and verify docx_edit_paragraph tool
```

---

## Appendix A: Phase 2 Migration Debt (Future Work)

These items are NOT part of this extraction plan but should be addressed afterward to achieve full OOXML parity across all consumers. They are documented in `ROADMAP.md` as Phase 2/3 items.

### A.1 `executeInsertListItem()` in `agentic-tools.js` (lines 701-872)

**Problem:** Manually constructs list item OOXML by regex-extracting `numId`/`ilvl` from adjacent paragraph OOXML (lines 779-815) and building a `pkg:package` string template. This should be in the core.

**Approach:** Create `services/list-item-builder.js`:
```js
export function buildListItemOoxml(adjacentParaOoxml, text, options = {}) {
    // 1. Parse adjacentParaOoxml to extract w:numPr (numId, ilvl)
    // 2. Optionally adjust ilvl based on options.indentDelta
    // 3. Build properly structured list item paragraph OOXML
    // 4. Wrap in pkg:package via package-builder.js
    // Returns: { oxml: string, numId: string, ilvl: string }
}
```
The add-in's `executeInsertListItem` becomes a thin wrapper: read adjacent OOXML via Word API, call `buildListItemOoxml`, insert result via `insertOoxmlWithRangeFallback`.

### A.2 `executeConvertHeadersToList()` (lines 1076-1229)

**Problem:** Uses Word's `startNewList()`, `setLevelNumbering()`, `attachToList()` APIs. No OOXML equivalent in the core.

**Approach:** Create `services/list-conversion.js`:
```js
export function convertParagraphsToListOoxml(paragraphOoxmlArray, options = {}) {
    // 1. Parse each paragraph's OOXML
    // 2. Inject w:numPr into each paragraph's w:pPr
    // 3. Generate numbering definitions (abstractNum + num)
    // 4. Return { paragraphsOoxml: string[], numberingXml: string }
}
```
This is essentially bulk list-binding. The core already does this for new content in `list-generation.js`. The gap is doing it for *existing* paragraphs with their formatting intact. The `enforceListBindingOnParagraphNodes()` helper in `list-structural-fallback.js` already handles single paragraphs — extend to batches.

### A.3 `executeEditTable()` Row/Column Operations (lines 1247-1450)

**Problem:** `add_row`, `delete_row` use Word Table API. `update_cell` already delegates to reconciliation.

**Approach:** Extend `services/table-reconciliation.js`:
```js
export function addTableRowOoxml(tableOoxml, position, cellCount, options = {}) {
    // 1. Parse table OOXML
    // 2. Build a new w:tr with cellCount w:tc elements
    // 3. Insert at position in the table
    // 4. Return modified table OOXML
}

export function deleteTableRowOoxml(tableOoxml, rowIndex) {
    // 1. Parse table OOXML
    // 2. Remove the w:tr at rowIndex
    // 3. Adjust any vMerge references
    // 4. Return modified table OOXML
}
```

### A.4 `executeEditSection()` (lines 1466-1569)

**Problem:** Uses Word API `listItem.level` for section boundary detection and `paragraph.delete()` for removal.

**Approach:** The core already has `getParagraphListInfo()` in `core/list-targeting.js` which extracts list level from OOXML. Section boundary detection can be built on this + paragraph targeting. Lower priority since section editing is less common.

### A.5 Context Extraction (Phase 3, ROADMAP.md)

Replace `Word.Paragraph.load()` logic with pure OOXML parsing of the document body. The `ingestWordOoxmlToPlainText` and `ingestWordOoxmlToMarkdown` functions in `pipeline/ingestion-export.js` already provide this capability for standalone consumers. The gap is that the Word add-in still uses `paragraph.load('text')` for context gathering rather than parsing the OOXML body.
