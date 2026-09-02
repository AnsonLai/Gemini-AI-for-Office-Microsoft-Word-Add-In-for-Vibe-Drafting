# Agentic Tools Move-Out Plan (High ROI)

> **Migrated on 2026-08-29:** The remaining extraction and list-reliability work
> was consolidated into [`2026-08-29-agentic-tools-and-list-reliability.md`](../2026-08-29-agentic-tools-and-list-reliability.md).
> This document is retained in `migrated/` as historical detail.

## Goals
- Reduce complexity in `src/taskpane/modules/commands/agentic-tools.js` by moving reusable orchestration into reconciliation.
- Keep reconciliation modules Word-agnostic where possible so browser and MCP surfaces can reuse them.
- Preserve current behavior while clarifying the boundary between decision logic and Word application.

## Status Snapshot (2026-02-09)

### Completed (Implemented, but with different module shape than originally proposed)
- Shared list marker detection was extracted to `src/taskpane/modules/reconciliation/pipeline/list-markers.js`.
- Shared content classification/parsing was extracted to `src/taskpane/modules/reconciliation/pipeline/content-analysis.js`.
- List/table generation is centralized in reconciliation:
  - `ReconciliationPipeline.executeListGeneration(...)` delegates to `pipeline/list-generation.js`.
  - Table generation is handled by `services/table-reconciliation.js` and `ReconciliationPipeline.executeTableGeneration(...)`.
- The hybrid router (`engine/oxml-engine.js`) now routes list content through reconciliation pipeline logic rather than command-local parsing.

### Partially Completed
- `executeRedline` and `executeEditList` use reconciliation for many list/table paths, but command-layer branching still contains substantial tool-level orchestration and OOXML write logic.
- Numbering behavior has been centralized in `services/numbering-service.js`; a smaller set of command-local list and header-conversion routines remains to migrate.

### Not Completed
- The originally proposed files:
  - `reconciliation/orchestration/route-change.js`
  - `reconciliation/orchestration/list-table.js`
 were not created.

## Current High-ROI Extraction Targets

1. `routeChangeOperation` plan-builder extraction (highest ROI) - Completed (2026-02-09)
- Current location: `src/taskpane/modules/commands/agentic-tools.js`.
- Extract pure decision logic into a reconciliation planner module.
- Suggested output contract:
  - `kind` (`text`, `list`, `table`, `formatOnly`, `formatRemoval`, `htmlFallback`)
  - `engineInput` (`originalOoxml`, `originalText`, `newContent`, options)
  - `requiresNativeFormatting`
  - `warnings`

Implemented:
- Added planner module: `src/taskpane/modules/reconciliation/orchestration/route-plan.js`.
- Added public exports via `reconciliation/index.js` and `reconciliation/standalone.js`.
- Added shared Word apply helper: `src/taskpane/modules/reconciliation/integration/word-route-change.js`.
- `routeChangeOperation(...)` in `agentic-tools.js` now delegates to reconciliation helper for route selection + Word apply sequencing.
- Structured-list `edit_paragraph` route now uses reconciliation list-generation + wrapped numbering insertion first, with legacy direct insertion as fallback.

2. Remove duplicate list parsing logic in command utils - Completed (2026-02-09)
- Current duplication:
  - `parseMarkdownList(...)` in `src/taskpane/modules/utils/markdown-utils.js`
  - `parseListItems(...)` and marker logic in reconciliation pipeline modules
- Consolidate command-layer list detection to reconciliation exports (`parseListItems`, marker helpers) and keep one canonical regex set.

Implemented:
- Added shared parser module: `src/taskpane/modules/reconciliation/orchestration/list-parsing.js`.
- `markdown-utils.parseMarkdownList(...)` now delegates to `parseMarkdownListContent(...)`.
- Command-layer list marker checks in `agentic-tools.js` now use parsed list data instead of bespoke regex literals in replace paths.

3. Move command-local OOXML list package construction into reconciliation modules - Completed (2026-02-09)
- Current functions in `agentic-tools.js`:
  - `applyStructuredListDirectOoxml`
  - `inferNumberingStyleFromMarker`
  - `escapeXmlText`
- Replacement direction:
  - Reuse `services/numbering-service.js` for format detection/numId decisions.
  - Reuse `services/package-builder.js` for package construction.

Implemented:
- Moved legacy direct structured-list insertion fallback into `src/taskpane/modules/reconciliation/integration/word-structured-list.js`.
- Moved numbering-style inference + list markdown builders into `src/taskpane/modules/reconciliation/orchestration/list-markdown.js`.
- `agentic-tools.js` now imports these helpers from reconciliation instead of defining local copies.

4. Centralize Word-specific OOXML application helpers under reconciliation integration - Completed (2026-02-09)
- Current functions in `agentic-tools.js`:
  - `insertOoxmlWithRangeFallback`
  - OOXML read fallback chain (`paragraph.getOoxml` -> range -> parent cell/table)
  - repeated tracking-mode toggle patterns
- Move into `reconciliation/integration/*` as reusable Word adapter helpers.

Implemented:
- Added `src/taskpane/modules/reconciliation/integration/word-ooxml.js` with:
  - `getParagraphOoxmlWithFallback(...)`
  - `insertOoxmlWithRangeFallback(...)`
  - `withNativeTrackingDisabled(...)`
- Exported these from `reconciliation/index.js`.
- Updated `agentic-tools.js` to consume shared integration helpers across:
  - route-level OOXML read fallback
  - OOXML insertion fallback
  - native tracking toggle wrappers in list/table and surgical insertion paths

5. Consolidate list markdown construction helpers - Completed (2026-02-09)
- Current command-local helpers:
  - `buildListMarkdown`, `toRoman`, `toAlphaSequence`, `buildListMarker`
- Move to a reusable reconciliation utility module to support both command tools and non-Word runtimes.

Implemented:
- Added `buildListMarkdown(...)` and marker-style helpers in `reconciliation/orchestration/list-markdown.js`.
- Added `normalizeListItemsWithLevels(...)` in `reconciliation/orchestration/list-markdown.js` and switched `executeEditList(...)` to use it.
- `executeEditList` now consumes the shared builder via reconciliation exports.

6. Extract paragraph identity parsing utility - Completed (2026-02-09)
- Current command-local helper: `extractParagraphIdFromOoxml`.
- Move to reconciliation `core` utilities for shared use by add-in and MCP/document adapters.

Implemented:
- Added `src/taskpane/modules/reconciliation/core/ooxml-identifiers.js`.
- `agentic-tools.js` now imports `extractParagraphIdFromOoxml(...)` from reconciliation exports.

## Recommended Sequencing
1. Extract planner (`buildReconciliationPlan`) with no Word calls. (Completed)
2. Move insertion/read fallback + tracking toggles into integration adapter. (Completed)
3. Remove list parsing/numbering duplication in command layer. (Completed)
4. Replace remaining command-local OOXML builders with reconciliation services/modules. (Completed for current scoped candidates)

## Expected Outcomes
- `agentic-tools.js` becomes a thin orchestration surface around tool I/O and Word runtime calls.
- Reconciliation owns the reusable document logic for list/table/text/format decisions.
- Browser and MCP surfaces can adopt more of the same planner/generation logic with minimal branching.
