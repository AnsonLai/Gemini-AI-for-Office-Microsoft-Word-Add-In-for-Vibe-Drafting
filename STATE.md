# Project State: Gemini Word Add-in

This document records the current technical posture and "sacred" decisions of the project.

## Current Posture
- **Engine Version**: Hybrid Mode V5.1.
- **Key Strategy**: Surgical DOM manipulation is prioritized over full string serialization to ensure Word detects pre-embedded redlines.
- **Portability Status**: Core reconciliation logic is 80% independent of Word JS API.

## Sacred Decisions
1. **OOXML > Word JS**: Any document modification MUST be implemented in the OOXML engine if possible. Word JS is reserved for UI and host-specific discovery.
2. **Migrate Logic Inward**: Minimize duplicated business logic outside the engine. The engine should reach "feature completion" so that all runtimes (Word, Browser, MCP) share the same capabilities.
3. **Word-Level Diffing**: Always use word-level granularity for diffs to ensure human-readable redlines.
4. **Explicit Offlining**: Formatting removal must explicitly emit "OFF" attributes (e.g., `w:b w:val="0"`) to override Word's style inheritance.
5. **Pkg:Package Wrapping**: All OOXML fragments must be wrapped in `pkg:package` structure for reliable insertion across different Office platforms.

## Active Blockers
- **Word Online Redline Bug**: Word Online sometimes ignores `w:rPrChange` during insertion, necessitating the "Surgical Replacement" workaround (wrap original in `w:del`, new in `w:ins`).
- **Mixed Content Controls**: Nested Content Controls (`w:sdt`) can sometimes cause offset drift during ingestion; requires careful sentinel tracking.

## Recent Architectural Shifts
- **Feb 2026**: Transitioned from fragmented `.md` docs to the **GSD Workflow** (`SPEC`, `ARCH`, `ROADMAP`, `STATE`, `PLAN`, `SUMMARY`).
- **Jan 2026**: Migrated from HTML fallback to pure OOXML for list and table insertion.
