# OXML Engine Refactor Task List

## Status Legend
- [ ] Pending
- [~] In progress
- [x] Completed

## Work Items
- [x] Phase 1: Add XML adapter and migrate all `DOMParser`/`XMLSerializer` usage
- [x] Phase 2: Add logger adapter and migrate `console.log/warn/error` usage
- [x] Phase 3: Split `oxml-engine.js` into focused modules
- [x] Phase 4: Deduplicate repeated logic patterns
- [x] Phase 5: Add JSDoc and section headers for all new modules
- [x] Phase 6: Add standalone entrypoint + standalone usage guide
- [~] Phase 7: Update tests, add smoke checks, and verify no Word API leakage

## Running Notes
- Started implementation and baseline code audit.
- Added `xml-adapter.js` and `logger.js` adapters and wired reconciliation modules.
- Split `oxml-engine.js` into new modules: `rpr-helpers.js`, `format-extraction.js`, `format-application.js`, `table-cell-context.js`, `surgical-mode.js`, `reconstruction-mode.js`, `run-builders.js`.
- Added standalone entrypoint `src/taskpane/modules/reconciliation/standalone.js`.
- Added standalone usage documentation `STANDALONE_USAGE.md`.
- Added test harness `tests/setup-xml-provider.mjs`.
- Added smoke test `tests/standalone_smoke.mjs` and no-Word-API check `tests/no_word_api_standalone_check.mjs`.
- `node tests/standalone_smoke.mjs` passes.
- `node tests/no_word_api_standalone_check.mjs` passes.
- Existing `tests/*.mjs` suites requiring `jsdom` could not be executed in this environment because `jsdom` is not installed.
- Follow-up fix: surgical format-only mode now reconciles target core formats (bold/italic/underline/strikethrough) instead of add-only behavior, fixing underline requests on pre-bolded text.
- Reorganized reconciliation modules into subfolders:
  - `adapters/`, `core/`, `engine/`, `pipeline/`, `services/`, `integration/`
  - Kept `index.js` and `standalone.js` as top-level entrypoints.
- Updated imports across `src/` and `tests/` to match new layout.
- Added architecture docs:
  - `src/taskpane/modules/reconciliation/ARCHITECTURE.md`
  - `src/taskpane/modules/reconciliation/README.md`
- Browser demo fix: prevented invalid `word/document.xml` structure by ensuring inserted demo markers are added before `w:sectPr` and normalizing `w:body` child order so `w:sectPr` remains last.
- Browser demo hardening: ignore `w:sectPr` in replacement fragments and avoid overwriting an existing `word/numbering.xml` part during kitchen-sink list insertion.
