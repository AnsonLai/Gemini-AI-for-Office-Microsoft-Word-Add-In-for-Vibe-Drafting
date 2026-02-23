# Project State: Gemini Word Add-in

> Role: **Memory across sessions** for decisions, blockers, and the current technical position.

This document records the current technical posture and "sacred" decisions of the project.

## Current Posture
- **Engine Version**: Hybrid Mode V5.1.
- **Key Strategy**: Surgical DOM manipulation is prioritized over full string serialization to ensure Word detects pre-embedded redlines.
- **Portability Status**: Core reconciliation logic is consumed from external package `@gsd/docx-reconciliation`.

## Sacred Decisions
1. **OOXML > Word JS**: Any document modification MUST be implemented in the OOXML engine if possible. Word JS is reserved for UI and host-specific discovery.
2. **Migrate Logic Inward**: Minimize duplicated business logic outside the engine. The engine should reach "feature completion" so that all runtimes (Word, Browser, MCP) share the same capabilities.
3. **Word-Level Diffing**: Always use word-level granularity for diffs to ensure human-readable redlines.
4. **Explicit Offlining**: Formatting removal must explicitly emit "OFF" attributes (e.g., `w:b w:val="0"`) to override Word's style inheritance.
5. **Pkg:Package Wrapping**: All OOXML fragments must be wrapped in `pkg:package` structure for reliable insertion across different Office platforms.

## Active Blockers
- **Word Online Redline Bug**: Word Online sometimes ignores `w:rPrChange` during insertion, necessitating the "Surgical Replacement" workaround (wrap original in `w:del`, new in `w:ins`).
- **Mixed Content Controls**: Nested Content Controls (`w:sdt`) can sometimes cause offset drift during ingestion; requires careful sentinel tracking.
- **Add-in Migration Debt**: Some add-in command flows still rely on Word APIs for feature-specific operations (for example complex header-to-list conversion and table row/column edits).

## Recent Architectural Shifts
- **Feb 2026**: Transitioned from fragmented docs to the **GSD Workflow** (`SPEC`, `ARCH`, `ROADMAP`, `STATE`).
- **Feb 2026**: Unified router logic in `oxml-engine.js` with specific modes for format removal and text-to-table.
- **Feb 22, 2026**: Completed reconciliation core extraction prep:
  - Entry points normalized (`index.js` primary, `standalone.js` compatibility alias, `word-addin-entry.js` add-in local).
  - Runtime defaults moved to `adapters/config.js` (no hardcoded author/platform and no core Office global reads).
  - Core helpers extracted (`engine/formatting-removal.js`, `services/numbering-helpers.js`, core targeting helpers).
  - Test ownership split into `tests/core/` and `tests/addin/`.
  - Prep package metadata added as `@gsd/docx-reconciliation` (`src/taskpane/modules/reconciliation/package.json`).
- **Feb 23, 2026**: Completed repository split cutover in AIWordPlugin:
  - Installed local package dependency (`file:../../Docx Redline JS`) and removed in-repo `src/taskpane/modules/reconciliation/`.
  - Moved add-in-only bridge to `src/taskpane/modules/reconciliation-integration/`.
  - Updated add-in, browser demo, and MCP imports to `@gsd/docx-reconciliation`.
  - Updated add-in integration tests to reference new package and bridge paths.
- **Jan 2026**: Migrated from HTML fallback to pure OOXML for list and table insertion.
