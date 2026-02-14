# Word Add-in (Taskpane)

This sub-project implements the Microsoft Word UI and integration layer for the **OOXML Reconciliation Core**.

## Architecture Alignment (GSD Workflow)

The Word Add-in is the primary host for the core engine. It follows the global project documentation:
- **[SPEC.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/SPEC.md)**: Host-agnostic vision and project purpose.
- **[ARCHITECTURE.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/ARCHITECTURE.md)**: Global project map and core engine deep-dive.
- **[ROADMAP.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/ROADMAP.md)**: Migration status for removing Office.js dependencies.

## Key Components

- **`taskpane.js`**: Orchestrates chat logic and document context extraction.
- **`agentic-tools.js`**: Translates AI function calls into document operations.
- **`taskpane.html/css`**: The React-based (or vanilla) user interface.

## Host-Specific Logic

While the core engine is portable, this sub-project contains the necessary "glue" to interact with the Word JS API:
- Document context discovery.
- OOXML injection into the active document.
- UI state management (Thinking... indicators, chat history).

## Long-Term Vision
Eventually, this folder will be split into its own repository (`word-addin-host`) using the `core` engine as a dependency.
