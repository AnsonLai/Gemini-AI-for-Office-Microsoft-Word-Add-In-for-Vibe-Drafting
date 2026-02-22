# Reconciliation Core Architecture

This document describes the host-agnostic OOXML reconciliation core and how to work inside it safely.

## Scope

The package boundary is the host-independent code under this folder:

- Included in core package: `adapters/`, `core/`, `engine/`, `pipeline/`, `services/`, `orchestration/`, `index.js`, `standalone.js`.
- Add-in local only (not part of published core package): `word-addin-entry.js` and `integration/`.

## Goals

- Preserve Word-compatible redlines by editing OOXML directly.
- Keep core logic host-independent (no Office.js globals, no Word API calls).
- Reuse the same engine in browser, Node, and Word-hosted environments.

## Folder Layout (Core Package)

```text
reconciliation/
├── adapters/
│   ├── config.js
│   ├── logger.js
│   └── xml-adapter.js
├── core/
├── engine/
│   └── formatting-removal.js
├── orchestration/
├── pipeline/
├── services/
│   ├── numbering-helpers.js
│   ├── standalone-docx-plumbing.js
│   └── standalone-operation-runner.js
├── index.js
└── standalone.js
```

## Entry Points

- `index.js`: primary host-agnostic package entrypoint.
- `standalone.js`: compatibility alias that re-exports from `index.js` (deprecated for new imports).
- `word-addin-entry.js`: add-in-local entrypoint that re-exports integration helpers; excluded from future published package.

## Module Responsibilities

- `adapters/config.js`
  - Runtime configuration surface for defaults (`setDefaultAuthor`, `getDefaultAuthor`, `setPlatform`, `getPlatform`).
  - Removes hardcoded author/platform values from core modules.
- `adapters/xml-adapter.js`
  - XML parser/serializer injection for browser or Node runtimes.
- `adapters/logger.js`
  - Runtime logger injection and shared log methods.
- `core/*`
  - Shared types, OOXML identity helpers, target resolution, list/table targeting heuristics, and XML query helpers.
- `engine/oxml-engine.js`
  - Main reconciliation router and mode selection.
- `engine/formatting-removal.js`
  - Shared formatting removal/highlight helpers extracted into engine scope.
- `pipeline/*`
  - Ingestion, markdown preprocessing, diffing, patching, and serialization stages.
- `services/numbering-helpers.js`
  - Dynamic numbering ID allocation, numbering payload remapping, and schema-order-safe numbering merges.
- `services/standalone-docx-plumbing.js`
  - Package-level extraction/wiring/validation for `word/document.xml`, `word/numbering.xml`, and `word/comments.xml`.
- `services/standalone-operation-runner.js`
  - Shared host-agnostic operation bridge for `redline`, `highlight`, and `comment`.
- `orchestration/*`
  - Word-agnostic route planning and list fallback orchestration utilities.

## End-to-End Flow

1. Caller imports from `index.js` (or legacy `standalone.js` alias).
2. Caller optionally configures XML provider/logger/defaults via `adapters/*`.
3. Caller invokes reconciliation APIs (`applyRedlineToOxml`, operation runner, ingestion/export helpers).
4. `engine/oxml-engine.js` routes to format, table, list, surgical, or reconstruction flows.
5. Pipeline/services return OOXML and optional package artifacts (`numberingXml`, comments payloads).
6. Host layer is responsible for writing OOXML back to document/package boundaries.

## Public Surfaces

- Primary: `reconciliation/index.js`
- Compatibility alias: `reconciliation/standalone.js` (deprecated)

Keep exports centralized through `index.js`; only maintain `standalone.js` for backward compatibility.

## Fast Orientation For Contributors

Use this sequence to understand or modify behavior without reading everything:

1. Start at `index.js` to locate the exported API.
2. Follow the export to `engine/oxml-engine.js` or the relevant `services/*` module.
3. For targeting bugs, inspect `core/paragraph-targeting.js`, `core/list-targeting.js`, and `core/table-targeting.js`.
4. For package wiring issues, inspect `services/standalone-docx-plumbing.js`.
5. For numbering issues, inspect `services/numbering-helpers.js` and list fallback orchestration.
