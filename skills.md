# Project Skills Notes

This file captures practical project-level context that helps when working in this repo.

## Core Principle

The strategic direction of this project is to build a **common, OOXML-first reconciliation engine** that owns as much feature logic as possible.

- New behavior should be implemented in shared reconciliation modules first when feasible.
- Runtime layers (Word add-in UI/runtime, browser demo, MCP server) should stay thin and focus on orchestration, I/O, and environment glue.
- Minimize duplicated business logic outside the engine.
- Minimize environment-specific branches outside the engine unless runtime APIs make them unavoidable.
- If logic starts outside the engine for practical reasons, it should be treated as migration debt and moved inward when possible.

## Demo Environments

This project has two demo environments outside the Word add-in runtime:

1. `browser-demo/`
2. `mcp/docx-server/`

### `browser-demo/` (Standalone Browser Demo)

- Runs reconciliation directly in the browser through `src/taskpane/modules/reconciliation/standalone.js`.
- Opens a `.docx` with JSZip, applies reconciliation operations, then writes package updates back.
- Demonstrates OOXML-first editing flows (text redlines, format changes, list/table transforms, comments/highlights).
- Handles package-level artifacts (for example `word/numbering.xml` and `word/comments.xml`) and validates output package consistency.
- Does not use Word JS APIs.

### `mcp/docx-server/` (Local MCP Demo/Tooling)

- Runs as a local MCP server over stdio for tool-driven `.docx` editing.
- Uses the same standalone reconciliation engine as the browser demo.
- Supports session-based document workflows: create/open, list paragraphs, edit paragraph, add comment, save, close.
- Applies reconciliation results to paragraph targets and merges emitted numbering/comment artifacts into package parts.
- Does not use Word JS APIs; operations requiring native Word fallback cannot be completed in MCP mode.

## How Demos Interact With The Main Project

- Both demos reuse the reconciliation architecture documented in:
  - `src/taskpane/modules/reconciliation/ARCHITECTURE.md`
- Both are OOXML-first integration surfaces for the same core engine used by the main Word add-in project.
- The main Word add-in (`src/taskpane/...`) is the primary product surface and should primarily orchestrate runtime-specific concerns (selection, tracking mode integration, tool routing, chat/UI flow), while core document transformation logic stays in the shared engine.
- Demo behavior is useful for validating reconciliation changes independently of Office runtime constraints.
