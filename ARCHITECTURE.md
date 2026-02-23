# System Architecture and Project Map

Role: system map for the external reconciliation package and its runtime consumers.

## Project Map

This repository consumes one reusable external package.

```mermaid
graph TD
    Core[@ansonlai/docx-redline-js\nexternal dependency] --> Word[Word Add-in\nsrc/taskpane]
    Core --> Demo[Browser Demo\nbrowser-demo]
    Core --> MCP[MCP docx server\nmcp/docx-server]
```

## Subprojects

1. `src/taskpane`: Word add-in host.
2. `src/taskpane/modules/docx-redline-js-integration`: add-in-only bridge layer around Word APIs.
3. `browser-demo`: browser runtime for manual validation and demo workflows.
4. `mcp/docx-server`: Node runtime exposing reconciliation as MCP tools.

## Core Boundary and Entrypoints

- Host-agnostic engine now comes from `@ansonlai/docx-redline-js`.
- Add-in local bridge entrypoint is `src/taskpane/modules/docx-redline-js-integration/index.js`.
- `src/taskpane/modules/reconciliation/` was removed after extraction.

## Import Path Conventions

- Word add-in command modules import the bridge:
  - `src/taskpane/modules/commands/agentic-tools.js` -> `../docx-redline-js-integration/index.js`
- Browser and MCP consumers import package entrypoints directly:
  - `browser-demo/demo.js` -> `@ansonlai/docx-redline-js`
  - `mcp/docx-server/src/services/docx-redline-js-service.mjs` -> `@ansonlai/docx-redline-js`

## Reconciliation Engine Overview

The engine reconciles text/markdown edits into Word-compatible OOXML with track changes.
Major paths include format-only, formatting removal, surgical edits, reconstruction, list generation, and table reconciliation.

Pipeline stages:

1. Ingestion
2. Markdown preprocessing
3. Word-level diffing
4. Patching
5. Serialization

## Portability Status

Core reconciliation code is host-agnostic and consumed as a package:

- No required Office.js or Word API references in package core modules.
- Runtime defaults (author/platform) are caller-configurable.
- Node hosts use `@xmldom/xmldom`; browser hosts use native DOM APIs.

## Operational Guidance

1. Prefer OOXML-first implementations for document manipulation.
2. Keep Word API usage inside `src/taskpane/modules/docx-redline-js-integration/`.
3. Route reusable logic to `@ansonlai/docx-redline-js` rather than command-layer duplication.
4. Update `STATE.md` and `ROADMAP.md` whenever package boundaries or entrypoints change.
