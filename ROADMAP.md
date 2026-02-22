# Project Roadmap

> Role: **Where we are going** and status tracking of what is done, in progress, and planned.

This document tracks migration from Word JS API to pure OOXML and the core package extraction track.

## Current Phase: Phase 2 (Feature Migration)

### ✅ Completed (Phase 1: Core Infrastructure)
- [x] **Text Editing**: Fully migrated to OOXML pipeline (`applyRedlineToOxml`).
- [x] **Highlighting**: Fully migrated to surgical engine.
- [x] **List Generation**: Fully migrated to OOXML pipeline.
- [x] **Table Generation**: Fully migrated to virtual grid pipeline.
- [x] **Pure Formatting**: Migrated to `w:rPrChange` surgical engine.
- [x] **Checkpoint System**: Uses whole-body OOXML storage.

### In Progress (Phase 2: Feature Migration)
- [ ] **List Conversion**: `executeConvertHeadersToList()` still uses Word List API for complex conversion.
- [ ] **Table Editing**: `executeEditTable()` uses Word Table API for row/column insertion in existing tables.

### In Progress (Phase 3: Final Decoupling and Repo Split)
- [~] **Repository Split**: Reconciliation package boundary is prepared in-repo as `@gsd/docx-reconciliation`.
  - Current package-ready structure lives at `src/taskpane/modules/reconciliation/`.
  - Core entrypoint layout is finalized: `index.js` (primary), `standalone.js` (compatibility alias), and `word-addin-entry.js` (add-in local, not for publish).
  - Test ownership is split: `tests/core/` (core) and `tests/addin/` (add-in integration).
  - Next step: move publishable package files into a dedicated repository and repoint add-in/demo/MCP imports to published package coordinates.
- [ ] **Context Extraction**: Replace `Word.Paragraph.load()` logic with pure OOXML parsing of the document body.
- [ ] **Comment Operations**: Replace remaining host-only comment entrypoints with direct OOXML-first flows where applicable.
- [ ] **Navigation**: Implement OOXML-based position tracking.
- [ ] **Search**: Implement pure OOXML text search parser.

### Planned (Phase 3+: Remaining Host-Decoupling Work)
- [ ] **Host Consumer Migration**: complete downstream migration away from `standalone.js` compatibility imports.
- [ ] **Repository Separation**: split `core`, `word-addin`, `browser-demo`, and `mcp` into independent repositories.

## Long-Term Vision
1. **Zero Office.js Logic**: All document manipulation happens via OOXML.
2. **CLI Utility**: Ship the core engine as a standalone `.docx` repair/edit tool.
3. **Server-Side Integration**: Use the engine in a Node.js backend to process large batches of legal documents.
