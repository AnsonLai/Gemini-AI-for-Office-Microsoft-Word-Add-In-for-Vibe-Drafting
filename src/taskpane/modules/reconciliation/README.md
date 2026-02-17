# Reconciliation Module

Use this folder for all OOXML reconciliation logic.

- Public API for add-in/internal: `index.js`
- Public API for standalone environments: `standalone.js`
- Architecture guide: `ARCHITECTURE.md`

Core code is organized by concern:

- `adapters/` runtime adapters (`xml-adapter`, `logger`)
- `core/` shared types/constants + revision/offset/XML-query policy helpers
  - includes OOXML identifier extraction utilities
- `engine/` router + mode implementations
  - includes focused formatting helpers (`format-paragraph-targeting`, `format-span-application`)
- `pipeline/` run-model reconciliation pipeline stages + shared list marker parser
  - includes hot-path indexed patch lookups for diff application
- `services/` table/comment/numbering services + shared package/plumbing helpers
  - includes `standalone-docx-plumbing.js` for OOXML output extraction, package artifact wiring, and package-level validation used by standalone/browser hosts
  - includes `standalone-operation-runner.js` for applying `redline`/`highlight`/`comment` operations to full `word/document.xml` payloads with shared targeting heuristics (explicit range and single-paragraph concatenation list cases use surgical insertion-only handling to preserve list binding/numbering style)
- `orchestration/` Word-agnostic planning helpers for command adapters
  - includes shared markdown list parsing, list markdown builders, list-item normalization helpers, and single-line structural list fallback helpers
- `integration/` Word API bridge + shared Word-only OOXML interop helpers
  - includes legacy structured-list insertion fallback helpers extracted from command layer
  - includes shared paragraph route/apply helper (`word-route-change.js`) used by command adapters
