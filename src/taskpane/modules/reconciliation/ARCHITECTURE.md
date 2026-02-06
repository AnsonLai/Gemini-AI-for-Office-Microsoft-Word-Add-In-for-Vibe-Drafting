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
│   └── types.js
├── engine/
│   ├── oxml-engine.js
│   ├── surgical-mode.js
│   ├── reconstruction-mode.js
│   ├── format-extraction.js
│   ├── format-application.js
│   ├── rpr-helpers.js
│   ├── run-builders.js
│   └── table-cell-context.js
├── pipeline/
│   ├── pipeline.js
│   ├── ingestion.js
│   ├── diff-engine.js
│   ├── patching.js
│   ├── serialization.js
│   └── markdown-processor.js
├── services/
│   ├── comment-engine.js
│   ├── numbering-service.js
│   └── table-reconciliation.js
├── integration/
│   └── integration.js
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
  - Shared enums/constants (`RunKind`, `DiffOp`, `NS_W`, revision IDs).
- `pipeline/*`
  - General reconciliation pipeline for run-model diffing/patching/serialization.
  - Used for list generation and compatibility flows.
- `services/table-reconciliation.js`
  - Virtual-grid table diff and OOXML table serialization.
- `services/comment-engine.js`
  - OOXML-only comment insertion logic.
- `engine/oxml-engine.js`
  - Main router/orchestrator for text + formatting reconciliation.
  - Chooses modes and delegates work.
- `engine/surgical-mode.js`
  - In-place edits for table-heavy/structure-sensitive content.
- `engine/reconstruction-mode.js`
  - Rebuild-oriented flow for non-table content where structure changes are allowed.
- `engine/format-extraction.js`
  - Extracts spans and existing formatting from paragraphs/runs.
- `engine/format-application.js`
  - Applies format-only changes, span splitting, and rPr synchronization.
- `engine/rpr-helpers.js`
  - Canonical `w:rPr` order and format override/addition primitives.
- `engine/run-builders.js`
  - Shared constructors for runs and track-change nodes.
- `engine/table-cell-context.js`
  - Detects table-cell wrapper contexts and paragraph-only serialization.
- `integration/integration.js`
  - Word API bridge (`paragraph.getOoxml()/insertOoxml()`).
- `index.js`
  - Main public API surface.
- `standalone.js`
  - Public API surface with no Word API exports.

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

## Key Design Notes

- `RPR_SCHEMA_ORDER` in `engine/rpr-helpers.js` is the single ordering source for run property insertion.
- `snapshotAndAttachRPrChange(...)` in `engine/run-builders.js` is the shared track-change snapshot routine.
- Format-only surgical flow now uses full target-state synchronization for core flags (bold/italic/underline/strikethrough), not additive-only behavior.

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
