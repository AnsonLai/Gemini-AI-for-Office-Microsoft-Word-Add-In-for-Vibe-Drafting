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
  - includes `ingestion-export.js` for Word OOXML -> readable plain text and basic markdown (`ingestWordOoxmlToPlainText`, `ingestWordOoxmlToMarkdown`)
- `services/` table/comment/numbering services + shared package/plumbing helpers
  - includes `standalone-docx-plumbing.js` for OOXML output extraction, package artifact wiring, and package-level validation used by standalone/browser hosts
  - includes `standalone-operation-runner.js` for applying `redline`/`highlight`/`comment` operations to full `word/document.xml` payloads with shared targeting heuristics (explicit range and single-paragraph concatenation list cases use surgical insertion-only handling to preserve list binding/numbering style)
  - standalone runner target resolution supports `targetRef` + text fallback (strict/fuzzy), including turn-snapshot drift correction for multi-operation turns
- `orchestration/` Word-agnostic planning helpers for command adapters
  - includes shared markdown list parsing, list markdown builders, list-item normalization helpers, and single-line structural list fallback helpers
- `integration/` Word API bridge + shared Word-only OOXML interop helpers
  - includes legacy structured-list insertion fallback helpers extracted from command layer
  - includes shared paragraph route/apply helper (`word-route-change.js`) used by command adapters

Formatting behavior note:

- Format-only redlines that cannot be localized surgically by span extraction now retry through OOXML reconstruction fallback in the shared engine before standalone native-fallback normalization is considered.

Useful standalone/add-in ingestion exports:

- `ingestWordOoxmlToPlainText(ooxml)` strips OOXML tags and returns readable paragraph-structured text
- `ingestWordOoxmlToMarkdown(ooxml)` returns a basic markdown projection (headings, bold/italic runs, obvious bullets/numbering)
