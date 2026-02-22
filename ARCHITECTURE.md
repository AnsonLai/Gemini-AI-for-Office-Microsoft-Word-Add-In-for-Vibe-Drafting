# System Architecture and Project Map

Role: System map for the host-agnostic OOXML reconciliation core and its runtime consumers.

## Project Map

This repository hosts multiple projects around one reusable core package.

```mermaid
graph TD
    Core[@gsd/docx-reconciliation\nsrc/taskpane/modules/reconciliation] --> Word[Word Add-in\nsrc/taskpane]
    Core --> Demo[Browser Demo\nbrowser-demo]
    Core --> MCP[MCP docx server\nmcp/docx-server]
```

## Subprojects

1. `src/taskpane/modules/reconciliation`: package-ready host-agnostic OOXML core.
2. `src/taskpane`: Word add-in host and add-in-only integration surface.
3. `browser-demo`: browser runtime for manual validation and demo workflows.
4. `mcp/docx-server`: Node runtime exposing reconciliation as MCP tools.

## Core Boundary and Entrypoints

The core package entrypoint layout is now:

- Primary host-agnostic entrypoint: `src/taskpane/modules/reconciliation/index.js`
- Compatibility alias (deprecated): `src/taskpane/modules/reconciliation/standalone.js`
- Add-in-local entrypoint: `src/taskpane/modules/reconciliation/word-addin-entry.js`

`integration/` remains local to the add-in side and is intentionally excluded from the future published package.

## Import Path Conventions

- Word add-in integration consumers import add-in entrypoint:
  - `src/taskpane/modules/commands/agentic-tools.js` -> `../reconciliation/word-addin-entry.js`
- Browser and Node host-agnostic consumers import primary core entrypoint:
  - `browser-demo/demo.js` -> `../src/taskpane/modules/reconciliation/index.js`
  - `mcp/docx-server/src/services/reconciliation-service.mjs` -> `../../../../src/taskpane/modules/reconciliation/index.js`

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

Core reconciliation code is now fully host-agnostic:

- No required Office.js or Word API references in core package files.
- Runtime defaults (author/platform) are caller-configurable via `adapters/config.js`.
- Node hosts use `@xmldom/xmldom`; browser hosts use native DOM APIs.

## Operational Guidance

1. Prefer OOXML-first implementations for document manipulation.
2. Keep Word API usage in add-in integration modules only.
3. Route new reusable logic into core package modules rather than command-layer duplication.
4. Update `STATE.md`, `ROADMAP.md`, and `src/taskpane/modules/reconciliation/ARCHITECTURE.md` whenever boundaries or entrypoints change.
