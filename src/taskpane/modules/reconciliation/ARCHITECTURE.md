# Reconciliation Architecture

This document explains how the OOXML reconciliation system is organized and how modules collaborate.

## Goals

- Preserve Word-compatible track changes by editing OOXML directly.
- Support both Word Add-in integration and standalone usage.
- Keep formatting-aware behavior predictable across surgical and reconstruction flows.

## Folder Layout

```text
reconciliation/
├── adapters/
│   ├── logger.js
│   └── xml-adapter.js
├── core/
│   ├── paragraph-offset-policy.js
│   ├── paragraph-targeting.js
│   ├── table-targeting.js
│   ├── ooxml-identifiers.js
│   ├── xml-query.js
│   └── types.js
├── engine/
│   ├── oxml-engine.js
│   ├── surgical-mode.js
│   ├── reconstruction-mode.js
│   ├── reconstruction-mapper.js
│   ├── reconstruction-writer.js
│   ├── format-extraction.js
│   ├── format-application.js
│   ├── format-paragraph-targeting.js
│   ├── format-span-application.js
│   ├── rpr-helpers.js
│   ├── run-builders.js
│   ├── table-mode.js
│   └── table-cell-context.js
├── pipeline/
│   ├── pipeline.js
│   ├── ingestion.js
│   ├── ingestion-paragraph.js
│   ├── ingestion-table.js
│   ├── ingestion-xml.js
│   ├── content-analysis.js
│   ├── diff-engine.js
│   ├── list-generation.js
│   ├── list-markers.js
│   ├── patching.js
│   ├── serialization.js
│   └── markdown-processor.js
├── services/
│   ├── comment-builders.js
│   ├── comment-engine.js
│   ├── comment-locator.js
│   ├── comment-package.js
│   ├── numbering-service.js
│   ├── package-builder.js
│   └── table-reconciliation.js
├── orchestration/
│   ├── route-plan.js
│   ├── list-markdown.js
│   └── list-parsing.js
├── integration/
│   ├── integration.js
│   ├── word-ooxml.js
│   ├── word-route-change.js
│   └── word-structured-list.js
├── index.js
└── standalone.js
```

## Module Responsibilities

- `adapters/xml-adapter.js`
  - Abstracts `DOMParser`/`XMLSerializer`.
  - Allows runtime injection for browser or Node.
- `adapters/logger.js`
  - Central logging surface (`log`, `warn`, `error`) and logger injection.
- `core/types.js`
  - Shared enums/constants (`RunKind`, `DiffOp`, `NS_W`) and revision metadata utilities.
- `core/paragraph-offset-policy.js`
  - Canonical paragraph-boundary separator policy used across extraction/ingestion/reconstruction.
- `core/paragraph-targeting.js`
  - Shared paragraph target resolution for per-paragraph operations.
  - Provides `targetRef` parsing (`P#`), strict/fuzzy text matching, and marker stripping helpers used by standalone consumers.
  - Provides table-targeting helpers (`isMarkdownTableText`, containing-word-element lookup) for table-scope operation routing in browser/Node integrations.
- `core/table-targeting.js`
  - Shared heuristics for table-structure intent from per-cell edits.
  - Provides markdown table synthesis from multiline table-cell edits so standalone callers can promote ambiguous cell redlines into full table-scope reconciliation.
  - Supports symmetric signature-row mirroring (for example `Title:` -> inserted `Date:` across peer columns).
- `core/ooxml-identifiers.js`
  - Shared OOXML identity extractors (`w14:paraId` and related paragraph tokens).
- `core/xml-query.js`
  - Shared namespace-safe XML query helpers for first/all lookups and parser-error detection.
- `pipeline/*`
  - General reconciliation pipeline for run-model diffing/patching/serialization.
  - Used for list generation and compatibility flows.
- `pipeline/ingestion-paragraph.js`
  - Paragraph/run ingestion with node-handler dispatch and numbering context detection.
- `pipeline/ingestion-table.js`
  - Virtual-grid table ingestion and merged-cell parsing.
- `pipeline/ingestion-xml.js`
  - Shared ingestion helpers for child-node traversal and attribute serialization.
- `pipeline/content-analysis.js`
  - Shared text classification/parsing helpers for list/table/paragraph detection.
- `pipeline/list-markers.js`
  - Shared list marker regex/detection helpers used by router/pipeline/patching.
- `pipeline/list-generation.js`
  - Generates list/table blocks from markdown lines.
  - Emits paragraph OOXML + optional `numberingXml` payload.
- `services/table-reconciliation.js`
  - Virtual-grid table diff and OOXML table serialization.
- `services/comment-engine.js`
  - OOXML-only comment insertion logic.
- `services/comment-builders.js`
  - Builds comment XML fragments and comment reference nodes.
- `services/comment-locator.js`
  - Locates target runs/ranges for comment anchoring in OOXML.
- `services/comment-package.js`
  - Handles comments-part wiring and relationship/content-type updates.
- `services/package-builder.js`
  - Shared `pkg:package` builders for document fragments, paragraph-only packages, and comments package variants.
- `orchestration/route-plan.js`
  - Word-agnostic route planner for command adapters (`buildReconciliationPlan`).
  - Classifies content into deterministic apply kinds (`structured_list_direct`, `empty_formatted_text`, `empty_html`, `block_html`, `ooxml_engine`).
- `orchestration/list-markdown.js`
  - Shared list markdown builders/marker-style helpers used by command adapters (`buildListMarkdown`, numbering style inference, list-item normalization).
- `orchestration/list-parsing.js`
  - Shared markdown list parsing for command adapters (`parseMarkdownListContent`).
  - Removes duplicate list marker regex/parsing logic from command utility surfaces.
- `engine/oxml-engine.js`
  - Main router/orchestrator for text + formatting reconciliation.
  - Chooses modes and delegates work.
- `engine/surgical-mode.js`
  - In-place edits for table-heavy/structure-sensitive content.
- `engine/reconstruction-mode.js`
  - Thin orchestration for reconstruction mapping + writing.
- `engine/reconstruction-mapper.js`
  - Builds reconstruction maps (paragraph/property/sentinel/reference) and indexed lookups.
- `engine/reconstruction-writer.js`
  - Applies diffs against mapped context and writes updated content fragments.
- `engine/format-extraction.js`
  - Extracts spans and existing formatting from paragraphs/runs.
- `engine/format-application.js`
  - Orchestrates format-only and surgical format-application flows.
- `engine/format-paragraph-targeting.js`
  - Paragraph text reconstruction/matching helpers used by format-only targeting.
- `engine/format-span-application.js`
  - Span boundary splitting and robust per-span format synchronization.
- `engine/rpr-helpers.js`
  - Canonical `w:rPr` order and format override/addition primitives.
- `engine/run-builders.js`
  - Shared constructors for runs and track-change nodes.
- `engine/table-cell-context.js`
  - Detects table-cell wrapper contexts and paragraph-only serialization.
- `engine/table-mode.js`
  - Table reconciliation/text-to-table transformation flows extracted from router.
- `integration/integration.js`
  - Word API bridge (`paragraph.getOoxml()/insertOoxml()`).
- `integration/word-ooxml.js`
  - Shared Word-only OOXML interop helpers:
    - OOXML read fallback chain (paragraph/range/table-cell/table)
    - OOXML insert fallback (`Paragraph.insertOoxml` -> range fallback on `GeneralException`)
    - temporary native tracking toggles around OOXML apply operations
- `integration/word-structured-list.js`
  - Legacy direct structured-list OOXML insertion fallback extracted from command layer for reuse and isolation.
- `integration/word-route-change.js`
  - Shared Word paragraph route/apply helper used by command adapters to keep `agentic-tools` thin.
  - Owns reconciliation route execution and fallback sequencing for single-paragraph edits.
- `index.js`
  - Main public API surface.
- `standalone.js`
  - Public API surface with no Word API exports.
  - Normalizes native-API fallback responses for standalone consumers (returns unchanged OOXML + warning when Word-native apply is required).
  - Re-exports shared paragraph-targeting helpers for browser/Node integrations.
  - Re-exports shared table-targeting heuristics for browser/Node integrations.

## End-to-End Flow

### 1) Entry

- Word Add-in path: `integration/integration.js` -> `index.js` -> `engine/oxml-engine.js`
- Standalone path: `standalone.js` -> `engine/oxml-engine.js`

### 2) Router (`engine/oxml-engine.js`)

The router:

1. Parses OOXML via `adapters/xml-adapter.js`.
2. Sanitizes AI text and extracts markdown format hints (`pipeline/markdown-processor.js`).
3. Extracts existing formatting and spans (`engine/format-extraction.js`).
4. Chooses a path:
   - Format removal
   - Format-only surgical application
   - Table reconciliation
   - Surgical mode
   - Reconstruction mode
   - List pipeline generation

When list-target content is detected, the router delegates to `ReconciliationPipeline.executeListGeneration(...)`,
which returns OOXML list paragraphs and numbering definitions suitable for `insertOoxml`.

### 3) Mode Execution

- Surgical mode:
  - Maintains existing structure and patches runs in place.
- Reconstruction mode:
  - Rebuilds paragraph content from diff segments and wrappers.
- Format-only path:
  - Splits spans on format boundaries and synchronizes target rPr state.

### 4) Output

- Returns `{ oxml, hasChanges }` to caller.
- Caller decides how/where OOXML is inserted or written back.
- List-generation paths may also return `numberingXml` for package wrapping.

## List Edit Integration (Command Layer)

`executeEditList` in `modules/commands/agentic-tools.js` now uses a two-stage reconciliation strategy:

1. Build normalized list markdown from tool args (`listType`, `numberingStyle`, indentation levels).
2. Run `applyRedlineToOxml(...)` over the full selected range.
3. If reconciliation reports no changes (`hasChanges === false`), force structural conversion via
   `ReconciliationPipeline.executeListGeneration(...)` and replace with wrapped OOXML + numbering.

This fallback is required for cases where source text is already similar (for example manual `A.`, `B.`, `C.` text)
but paragraphs are not true Word list items. In those cases, textual diff can be a no-op while structural list
conversion is still required.

Native Word tracking is intentionally disabled during insertion for these list operations because the OOXML already
contains explicit redline markup (`w:ins`/`w:del`) when redlines are enabled.

## Current Migration Status (Command Layer -> Reconciliation)

- Reconciliation now owns shared list marker parsing, content analysis, list generation, numbering service, and package builders.
- `routeChangeOperation` command-layer decisioning now uses reconciliation `buildReconciliationPlan(...)`.
- Word-specific OOXML fallback/toggle helpers are centralized in `integration/word-ooxml.js` and consumed by command-layer routes.
- Command-layer markdown list parsing now delegates to reconciliation `orchestration/list-parsing.js`.
- Structured list `edit_paragraph` application now prefers reconciliation list-generation output (with numbering package wiring) before legacy direct list insertion fallback.
- Command-layer list markdown builders, paragraph-id extraction, and direct structured-list fallback are now extracted to reconciliation modules.
- Command-layer `routeChangeOperation(...)` is now a thin wrapper over reconciliation `integration/word-route-change.js`.
- Command-layer list item normalization in `executeEditList(...)` now delegates to reconciliation orchestration helpers.
- `modules/commands/agentic-tools.js` still contains migration debt:
  - compatibility wrapper `parseMarkdownList(...)` remains in `markdown-utils` (delegating to reconciliation parser)
- Intended direction:
  - reconciliation modules produce deterministic operation plans/results
  - integration modules own reusable Word-specific apply/read/toggle adapters
  - command modules remain thin tool orchestration and error handling layers

## Key Design Notes

- `RPR_SCHEMA_ORDER` in `engine/rpr-helpers.js` is the single ordering source for run property insertion.
- `snapshotAndAttachRPrChange(...)` in `engine/run-builders.js` is the shared track-change snapshot routine.
- Format-only surgical flow now uses full target-state synchronization for core flags (bold/italic/underline/strikethrough), not additive-only behavior.
- Hot-path lookup indexing is used in patching/format/table loops to reduce repeated O(n) scans.

## Extension Points

- Replace parser/serializer:
  - `configureXmlProvider({ DOMParser, XMLSerializer })`
- Replace logger:
  - `configureLogger({ log, warn, error })`
- Add new mode logic:
  - Extend router decisions in `engine/oxml-engine.js`
  - Keep mode-specific logic inside `engine/*` modules

## Public Surfaces

- Add-in + internal usage: `reconciliation/index.js`
- Standalone usage: `reconciliation/standalone.js`

Keep new exports centralized through one of these entrypoints.
