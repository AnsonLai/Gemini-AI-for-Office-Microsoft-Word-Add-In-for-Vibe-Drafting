# Reconciliation Core Package Extraction Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Prepare the OOXML reconciliation core (`src/taskpane/modules/reconciliation/`) for extraction into its own standalone repository. Clean all host-specific (Office.js/Word JS API) and AI-specific (`'Gemini AI'`) contamination from the core, refactor the bloated `standalone.js` entry point, strengthen isolation tests, and set up package infrastructure.

**Guiding Principle:** The core is purely for `.docx` XML structure manipulation and markdown conversions. Nothing Word JS API or AI-related. Callers configure their own author name, platform, XML provider, and logger.

**Tech Stack:** Pure ES modules, `diff-match-patch` (only npm dependency), optional `@xmldom/xmldom` peer dependency for Node.js environments.

## Implementation Progress

- [x] **Task 1 complete** (February 22, 2026): moved formatting-removal module into `reconciliation/engine`, updated imports/re-exports, removed legacy file.
- [x] **Task 2 complete** (February 22, 2026): added `adapters/config.js` and exported config APIs from both `standalone.js` and `index.js`.
- [x] **Task 3 complete** (February 22, 2026): replaced `'Gemini AI'` defaults in reconciliation modules with `getDefaultAuthor()` and updated related JSDoc defaults.
- [x] **Task 4 complete** (February 22, 2026): removed core `Office.context.platform` reads, switched pipeline platform default to `getPlatform()`, and wired `setPlatform(Office.context.platform)` in `taskpane.js` startup.
- [x] **Task 5 complete** (February 22, 2026): extracted numbering ID/state helpers into `services/numbering-helpers.js`, replaced in-file implementations in `standalone.js` with re-exports, removed local `WORD_MAIN_NS` redeclaration, and added direct exports from `index.js`.
- [x] **Task 6 complete** (February 22, 2026): moved `resolveParagraphRangeByRefs()` into `core/paragraph-targeting.js`, moved table-source heuristics into `core/table-targeting.js`, replaced implementations in `standalone.js` with core re-exports, and updated `index.js` exports to point at `core/` helpers directly.
- [x] **Task 7 complete** (February 22, 2026): verified reconciliation-core dependency boundaries (`diff-match-patch` used only in diff engine, `marked` not used in reconciliation core) and confirmed root dependency declarations.
- [x] **Task 8 complete** (February 22, 2026): rewrote `no_word_api_standalone_check.mjs` to recurse all reconciliation subfolders (excluding `integration/` and `index.js`), added boundary/import checks, and preserved standalone isolation assertions.
- [x] **Task 9 complete** (February 22, 2026): added `core_dependency_graph_check.mjs` to enforce import-boundary rules across reconciliation core files.
- [x] **Task 10 complete** (February 22, 2026): normalized reconciliation entry points (`index.js` host-agnostic primary, `word-addin-entry.js` add-in entry, `standalone.js` compatibility shim), updated import call sites, and added prep `reconciliation/package.json`.
- [x] **Task 11 complete** (February 22, 2026): reorganized tests by ownership into `tests/core/` and `tests/addin/`, deleted improper one-off/debug scripts, and updated relocated test imports/fixture paths.
- [x] **Task 12 complete** (February 22, 2026): updated architecture/state/roadmap/core README docs to reflect package-ready boundaries, entrypoint naming, and test ownership split.

### Verification Log (Tasks 3-4)

- `node tests/reconciliation_author_platform_defaults_tests.mjs` (new): PASS
- `node tests/reconciliation_config_exports_tests.mjs`: PASS
- `node tests/standalone_smoke.mjs`: PASS
- `node tests/standalone_operation_runner_tests.mjs`: PASS
- `node tests/comment_tests.mjs`: PASS
- `node tests/no_word_api_standalone_check.mjs`: PASS
- `npm run build`: PASS (webpack warnings only)

### Verification Log (Task 5)

- `node tests/numbering_helpers_extraction_tests.mjs` (new): PASS
- `node tests/standalone_smoke.mjs`: PASS
- `node tests/list_tests.mjs`: PASS
- `node tests/standalone_docx_plumbing_tests.mjs`: PASS
- `npm run build`: PASS (webpack warnings only)

### Task 5 Change Notes

- Added new module: `src/taskpane/modules/reconciliation/services/numbering-helpers.js`.
- Replaced inlined numbering helper implementations in `src/taskpane/modules/reconciliation/standalone.js` with re-exports from `services/numbering-helpers.js`.
- Updated `src/taskpane/modules/reconciliation/index.js` to export numbering helper APIs directly from `services/numbering-helpers.js`.
- Added extraction regression test: `tests/numbering_helpers_extraction_tests.mjs`.

### Verification Log (Task 6)

- `node tests/targeting_helpers_extraction_tests.mjs` (new): PASS
- `node tests/standalone_smoke.mjs`: PASS
- `node tests/standalone_operation_runner_tests.mjs`: PASS
- `node tests/list_tests.mjs`: PASS

### Task 6 Change Notes

- Added `resolveParagraphRangeByRefs` to `src/taskpane/modules/reconciliation/core/paragraph-targeting.js`.
- Added `isLikelyStructuredTableSourceParagraph` and `inferTableReplacementParagraphBlock` to `src/taskpane/modules/reconciliation/core/table-targeting.js`.
- Replaced in-file helper implementations in `src/taskpane/modules/reconciliation/standalone.js` with:
  - `export { resolveParagraphRangeByRefs } from './core/paragraph-targeting.js';`
  - `export { inferTableReplacementParagraphBlock, isLikelyStructuredTableSourceParagraph } from './core/table-targeting.js';`
- Updated `src/taskpane/modules/reconciliation/index.js` to export the moved helpers from core modules instead of from `standalone.js`.

### Verification Log (Task 7)

- `rg -n "diff-match-patch|diff_match_patch" src/taskpane/modules/reconciliation`: PASS (`src/taskpane/modules/reconciliation/pipeline/diff-engine.js` is the only runtime import site).
- `rg -n "\\bmarked\\b" src/taskpane/modules/reconciliation`: PASS (no matches in reconciliation core).
- `rg -n "\\bmarked\\b" browser-demo src/taskpane/modules/chat src/taskpane/modules/utils`: PASS (`marked` is used outside reconciliation core for markdown rendering).
- `rg -n "diff-match-patch|marked" package.json`: PASS (root dependencies include both packages).
- `node tests/standalone_smoke.mjs`: PASS
- `npm run build`: PASS (webpack warnings only)

### Task 7 Change Notes

- No production-code edits required; this task was verification-only.

### Verification Log (Task 8)

- `node tests/no_word_api_standalone_check.mjs` with temporary nested violation file (`core/__tmp_word_api_violation__.js` containing `Office.context.platform`): FAIL as expected after rewrite (proves recursive scan catches nested core violations).
- `node tests/no_word_api_standalone_check.mjs`: PASS (clean workspace).

### Task 8 Change Notes

- Rewrote `tests/no_word_api_standalone_check.mjs` to recursively scan all `.js` files under `src/taskpane/modules/reconciliation/` while excluding `integration/`.
- Kept explicit `standalone.js` isolation check against `integration/` imports.
- Added enforcement that non-`integration/` files (excluding root `index.js`) cannot import from `integration/`.
- Added enforcement that non-`integration/` files (excluding root `index.js`) cannot import outside `reconciliation/`, with only `diff-match-patch` allowed as an external package.
- Updated import parsing to include static imports, side-effect imports, export-from specifiers, and dynamic `import()` calls.

### Verification Log (Task 9)

- `node tests/core_dependency_graph_check.mjs` with temporary violation file (`core/__tmp_dep_violation__.js` importing `../../../../taskpane/taskpane.js`): FAIL as expected (detects out-of-bound dependency).
- `node tests/core_dependency_graph_check.mjs`: PASS
- `node tests/standalone_smoke.mjs`: PASS
- `npm run build`: PASS (webpack warnings only)

### Task 9 Change Notes

- Added new test file: `tests/core_dependency_graph_check.mjs`.
- Implemented recursive dependency-graph validation for reconciliation core files, excluding `integration/` and root `index.js`.
- Validation now accepts:
  - Reconciliation-internal relative imports.
  - npm package imports (specifier does not start with `.` or `/`).
- Validation fails on any resolved path that escapes the `reconciliation/` directory tree.

### Verification Log (Task 10)

- `rg -n "reconciliation/(standalone|index|word-addin-entry)\\.js" src tests browser-demo mcp`: PASS (core consumers now point to `index.js`; add-in integration consumer points to `word-addin-entry.js`).
- `node tests/reconciliation_config_exports_tests.mjs`: PASS
- `node tests/reconciliation_author_platform_defaults_tests.mjs`: PASS
- `node tests/no_word_api_standalone_check.mjs`: PASS
- `node tests/core_dependency_graph_check.mjs`: PASS
- `node tests/standalone_smoke.mjs`: PASS
- `node tests/standalone_ingestion_export_tests.mjs`: PASS
- `node tests/standalone_docx_plumbing_tests.mjs`: PASS
- `node tests/standalone_operation_runner_tests.mjs`: PASS
- `node tests/list_tests.mjs`: PASS
- `node tests/numbering_helpers_extraction_tests.mjs`: PASS
- `node tests/targeting_helpers_extraction_tests.mjs`: PASS
- `node tests/integration_tests.mjs`: PASS
- `node tests/word_operation_runner_adapter_tests.mjs`: PASS
- `node tests/shared_operation_bridge_tests.mjs`: PASS
- `npm run build`: PASS (webpack warnings only)

### Task 10 Change Notes

- Renamed host-agnostic entry:
  - `src/taskpane/modules/reconciliation/standalone.js` → `src/taskpane/modules/reconciliation/index.js`
- Renamed Word-aware add-in entry:
  - `src/taskpane/modules/reconciliation/index.js` (previous) → `src/taskpane/modules/reconciliation/word-addin-entry.js`
- Added compatibility shim:
  - `src/taskpane/modules/reconciliation/standalone.js` now re-exports from `./index.js` with a deprecation notice.
- Updated host-agnostic consumers to `index.js` imports (browser demo, MCP service, and core/standalone-oriented tests).
- Updated add-in integration consumer import:
  - `src/taskpane/modules/commands/agentic-tools.js` now imports from `../reconciliation/word-addin-entry.js`.
- Added package prep artifact:
  - `src/taskpane/modules/reconciliation/package.json`
- Updated isolation tests from Task 8/9 to exclude `word-addin-entry.js` (add-in-only root entry) from core-only boundary scans.

### Verification Log (Task 11)

- `node tests/core/no_word_api_standalone_check.mjs`: PASS
- `node tests/core/core_dependency_graph_check.mjs`: PASS
- `node tests/core/standalone_smoke.mjs`: PASS
- `node tests/core/standalone_ingestion_export_tests.mjs`: PASS
- `node tests/core/standalone_docx_plumbing_tests.mjs`: PASS
- `node tests/core/standalone_operation_runner_tests.mjs`: PASS
- `node tests/core/list_tests.mjs`: PASS
- `node tests/core/formatting_tests.mjs`: PASS
- `node tests/core/table_tests.mjs`: PASS
- `node tests/core/redline_operation_converter_tests.mjs`: PASS
- `node tests/core/comment_tests.mjs`: PASS
- `node tests/core/highlight_tests.mjs`: PASS
- `node tests/core/table_targeting_and_format_flags.mjs`: PASS
- `node tests/addin/integration_tests.mjs`: PASS
- `node tests/addin/word_operation_runner_adapter_tests.mjs`: PASS
- `node tests/addin/shared_operation_bridge_tests.mjs`: PASS
- `node tests/addin/migrated_tool_cutover_tests.mjs`: PASS

### Task 11 Change Notes

- Moved classified core tests into `tests/core/`:
  - `standalone_smoke.mjs`, `standalone_ingestion_export_tests.mjs`, `standalone_docx_plumbing_tests.mjs`, `standalone_operation_runner_tests.mjs`
  - `no_word_api_standalone_check.mjs`, `core_dependency_graph_check.mjs`
  - `list_tests.mjs`, `formatting_tests.mjs`, `table_tests.mjs`, `redline_operation_converter_tests.mjs`
  - `comment_tests.mjs`, `highlight_tests.mjs`, `table_targeting_and_format_flags.mjs`
- Moved classified add-in tests into `tests/addin/`:
  - `integration_tests.mjs`, `shared_operation_bridge_tests.mjs`, `word_operation_runner_adapter_tests.mjs`, `migrated_tool_cutover_tests.mjs`
- Deleted improper non-test scripts:
  - `tests/debug_extraction.mjs`
  - `tests/debug_test6.mjs`
  - `tests/verify_fix.mjs`
- Updated relocated test imports:
  - `./setup-xml-provider.mjs` → `../setup-xml-provider.mjs`
  - `../src/...` → `../../src/...`
  - fixture paths updated to `../sample_doc/...` where required after relocation.

### Verification Log (Task 12)

- `rg -n "index\\.js|standalone\\.js|word-addin-entry\\.js|adapters/config\\.js|services/numbering-helpers\\.js|engine/formatting-removal\\.js|End-to-End Flow|Public Surfaces" src/taskpane/modules/reconciliation/ARCHITECTURE.md`: PASS
- `rg -n "@gsd/docx-reconciliation|word-addin-entry\\.js|integration/|reconciliation/index\\.js|reconciliation/standalone\\.js" ARCHITECTURE.md ROADMAP.md STATE.md src/taskpane/modules/reconciliation/README.md`: PASS
- `rg -n "Portability Status|100% independent|Repository Split|tests/core|tests/addin|Quick Start|configureXmlProvider\\(|setDefaultAuthor\\(|setPlatform\\(" STATE.md ROADMAP.md src/taskpane/modules/reconciliation/README.md`: PASS

### Task 12 Change Notes

- Updated `src/taskpane/modules/reconciliation/ARCHITECTURE.md`:
  - Reframed to core-package scope, removed `integration/` from core folder layout/module responsibilities, and clarified entrypoint roles (`index.js`, `standalone.js`, `word-addin-entry.js`).
  - Added explicit responsibilities for `adapters/config.js`, `services/numbering-helpers.js`, and `engine/formatting-removal.js`.
  - Rewrote end-to-end flow to host-agnostic path only and tightened contributor orientation guidance.
- Updated root `ARCHITECTURE.md`:
  - Switched project map to package-ready core (`@gsd/docx-reconciliation`) and clarified that `integration/` is add-in local.
  - Updated import path examples for add-in (`word-addin-entry.js`) vs browser/MCP (`index.js`).
- Updated `STATE.md`:
  - Set portability status to 100% independent at package boundary.
  - Recast migration debt as add-in-specific and added reconciliation extraction-prep milestone notes.
- Updated `ROADMAP.md`:
  - Moved repository split into in-progress status and documented current package structure/entrypoints plus test ownership split.
- Updated `src/taskpane/modules/reconciliation/README.md`:
  - Added required sections: Quick Start, API Overview, Configuration, Hosting Guidance, Architecture link, and browser-demo minimal integration example.

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

## Task 10: Package Prep (In-Repo Only)

### 10.1 Normalize Entry Point Naming

**Goal:** Make the host-agnostic core use standard package naming now, without splitting repositories yet.

**Changes:**
1. Rename host-agnostic entry file:
   - `src/taskpane/modules/reconciliation/standalone.js` → `src/taskpane/modules/reconciliation/index.js`
2. Move current Word-aware add-in entry to an explicit local file:
   - `src/taskpane/modules/reconciliation/index.js` (current) → `src/taskpane/modules/reconciliation/word-addin-entry.js`
3. Keep a compatibility shim so existing local imports do not break immediately:
   - Create `src/taskpane/modules/reconciliation/standalone.js` that re-exports from `./index.js`
   - Add a deprecation comment in the shim: new imports should use `./index.js`

**Rationale:** This keeps nomenclature normal for package consumers (`index.js`) while preserving add-in integration behavior via an explicitly named local entry.

### 10.2 Update Local Import Call Sites for the New Entry Layout

**Core consumers (host-agnostic):**
- Change imports from `.../reconciliation/standalone.js` to `.../reconciliation/index.js` in:
  - Browser demo
  - MCP docx server
  - Core tests that explicitly import `standalone.js`

**Word add-in integration consumers:**
- Change imports from `.../reconciliation/index.js` to `.../reconciliation/word-addin-entry.js` in:
  - `src/taskpane/modules/commands/agentic-tools.js`
  - Any other add-in-only modules that expect integration exports

**Verification:** `rg -n "reconciliation/(standalone|index|word-addin-entry)\\.js" src tests browser-demo mcp` and then run the impacted core and integration tests.

### 10.3 Create Prep `package.json` for Future Extraction

**New file:** `src/taskpane/modules/reconciliation/package.json`

This is a prep artifact for extraction and npm publishing later. Keep it `private` while still embedded in this repository.

```json
{
  "name": "@gsd/docx-reconciliation",
  "version": "0.1.0",
  "description": "Host-independent OOXML reconciliation engine for .docx manipulation with track changes",
  "type": "module",
  "private": true,
  "main": "index.js",
  "exports": {
    ".": "./index.js",
    "./standalone": "./standalone.js",
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
    "index.js",
    "standalone.js",
    "README.md"
  ],
  "keywords": ["docx", "ooxml", "reconciliation", "track-changes", "redlines", "word"],
  "license": "MIT"
}
```

**Notes:**
- This task is prep-only. Do not publish from this repository.
- `integration/` and `word-addin-entry.js` are intentionally excluded from package files.
- `@xmldom/xmldom` remains optional for Node.js; browsers can use native `DOMParser`.
- The extraction/publish sequence is documented in Appendix A.6.

---

## Task 11: Test Classification & Migration Prep

Classify which tests belong with the core package and which stay with the add-in:

| Test File | Belongs To | Reason |
|-----------|-----------|--------|
| `tests/core/standalone_smoke.mjs` | **Core** | Tests `applyRedlineToOxml` standalone |
| `tests/core/standalone_ingestion_export_tests.mjs` | **Core** | Tests `ingestWordOoxmlToPlainText/Markdown` |
| `tests/core/standalone_docx_plumbing_tests.mjs` | **Core** | Tests package plumbing |
| `tests/core/standalone_operation_runner_tests.mjs` | **Core** | Tests `applyOperationToDocumentXml` |
| `tests/core/no_word_api_standalone_check.mjs` | **Core** | Isolation boundary test |
| `tests/core/core_dependency_graph_check.mjs` | **Core** | Dependency validation (new in Task 9) |
| `tests/core/list_tests.mjs` | **Core** | Tests list reconciliation pipeline |
| `tests/core/formatting_tests.mjs` | **Core** | Tests formatting engine |
| `tests/core/table_tests.mjs` | **Core** | Tests table reconciliation |
| `tests/core/redline_operation_converter_tests.mjs` | **Core** | Tests orchestration converter |
| `tests/core/comment_tests.mjs` | **Core** | Imports only reconciliation core comment service modules (no `integration/`) |
| `tests/core/highlight_tests.mjs` | **Core** | Imports only reconciliation core formatting engine module (no `integration/`) |
| `tests/addin/integration_tests.mjs` | **Add-in** | Tests Word integration layer |
| `tests/addin/shared_operation_bridge_tests.mjs` | **Add-in** | Imports `integration/word-operation-runner.js` (Word integration layer) |
| `tests/addin/word_operation_runner_adapter_tests.mjs` | **Add-in** | Tests Word adapter |
| `tests/addin/migrated_tool_cutover_tests.mjs` | **Add-in** | Tests command-layer cutover |
| `tests/core/table_targeting_and_format_flags.mjs` | **Core** | Tests core targeting heuristics |
| `tests/debug_extraction.mjs` | **Deleted** | Improper debug script (removed in Task 11) |
| `tests/debug_test6.mjs` | **Deleted** | Improper debug script (removed in Task 11) |
| `tests/verify_fix.mjs` | **Deleted** | Improper one-off verification script (removed in Task 11) |

**Action for "Review" items:** Completed in Task 11. All previously "Review" items are now classified.

---

## Task 12: Documentation Updates

### 12.1 Update `src/taskpane/modules/reconciliation/ARCHITECTURE.md`

- Remove `integration/` from folder layout and module descriptions (it leaves with the add-in)
- Replace old entry-point assumptions with:
  - `index.js` = host-agnostic package entry (primary)
  - `standalone.js` = compatibility alias (deprecated)
  - `word-addin-entry.js` = add-in-local integration entry (not part of published package)
- Add `adapters/config.js` to module responsibilities
- Add `services/numbering-helpers.js` to module responsibilities
- Add `engine/formatting-removal.js` to module responsibilities
- Update "Public Surfaces" section: `index.js` primary + `standalone.js` compatibility alias
- Update "End-to-End Flow" section: remove Word Add-in path
- Ensure the ARCHITECTURE.md is drafted so other AI agents are able to understand the structure and code without reading the whole codebase for future projects

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
6. Add an example of the browser-demo, describing additional functions to make this package work minimally.

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
Task 10 (entrypoint + package prep)       ← after Tasks 1-9 (structure is final)
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
node tests/core/no_word_api_standalone_check.mjs
node tests/core/core_dependency_graph_check.mjs
node tests/core/standalone_smoke.mjs
node tests/core/standalone_ingestion_export_tests.mjs
node tests/core/standalone_docx_plumbing_tests.mjs
node tests/core/standalone_operation_runner_tests.mjs
node tests/core/list_tests.mjs
node tests/core/formatting_tests.mjs
node tests/core/table_tests.mjs
node tests/core/redline_operation_converter_tests.mjs
node tests/core/comment_tests.mjs
node tests/core/highlight_tests.mjs
node tests/core/table_targeting_and_format_flags.mjs

# Integration tests (should still pass — integration/ layer unchanged)
node tests/addin/integration_tests.mjs
node tests/addin/word_operation_runner_adapter_tests.mjs
node tests/addin/shared_operation_bridge_tests.mjs
node tests/addin/migrated_tool_cutover_tests.mjs

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

### A.6 Repository Split and npm Publish Sequence (After This Plan)

This sequence is intentionally out of scope for Task 10. Task 10 only prepares naming and boundaries.

1. Create a new repository from `src/taskpane/modules/reconciliation/` contents.
2. Exclude add-in-only artifacts from the package repo:
   - `integration/`
   - `word-addin-entry.js`
3. Remove `"private": true` from the extracted repo `package.json`.
4. Add release automation (versioning, changelog, publish workflow).
5. Publish an initial prerelease tag (for example `0.1.0-alpha.1`) and validate in:
   - Word add-in
   - Browser demo
   - MCP docx server
6. In the add-in repo, keep `integration/` local and repoint its core imports to `@gsd/docx-reconciliation`.
7. Remove the `standalone.js` compatibility alias in a later major version after downstream migration is complete.
