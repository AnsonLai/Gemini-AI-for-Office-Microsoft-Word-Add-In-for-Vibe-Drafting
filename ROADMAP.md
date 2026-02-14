# Project Roadmap

This document tracks the migration progress from Word JS API to pure OOXML for 100% portability.

## Current Phase: Phase 2 (Feature Migration)

### ✅ Completed (Phase 1: Core Infrastructure)
- [x] **Text Editing**: Fully migrated to OOXML pipeline (`applyRedlineToOxml`).
- [x] **Highlighting**: Fully migrated to surgical engine.
- [x] **List Generation**: Fully migrated to OOXML pipeline.
- [x] **Table Generation**: Fully migrated to virtual grid pipeline.
- [x] **Pure Formatting**: Migrated to `w:rPrChange` surgical engine.
- [x] **Checkpoint System**: Uses whole-body OOXML storage.

### 🚧 In Progress (Phase 2: Feature Migration)
- [ ] **List Conversion**: `executeConvertHeadersToList()` still uses Word List API for complex conversion.
- [ ] **Table Editing**: `executeEditTable()` uses Word Table API for row/column insertion in existing tables.

### 📋 Planned (Phase 3: Final Decoupling & Repo Split)
- [ ] **Context Extraction**: Replace `Word.Paragraph.load()` logic with pure OOXML parsing of the document body.
- [ ] **Comment Operations**: Replace `insertComment()` with direct OOXML comment injection.
- [ ] **Navigation**: Implement OOXML-based position tracking.
- [ ] **Search**: Implement pure OOXML text search parser.
- [ ] **Repository Split**: Segregate `core`, `word-addin`, `browser-demo`, and `mcp` into independent repositories.

## Long-Term Vision
1. **Zero Office.js Logic**: All document manipulation happens via OOXML.
2. **CLI Utility**: Ship the core engine as a standalone `.docx` repair/edit tool.
3. **Server-Side Integration**: Use the engine in a Node.js backend to process large batches of legal documents.
