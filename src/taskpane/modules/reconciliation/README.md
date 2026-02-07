# Reconciliation Module

Use this folder for all OOXML reconciliation logic.

- Public API for add-in/internal: `index.js`
- Public API for standalone environments: `standalone.js`
- Architecture guide: `ARCHITECTURE.md`

Core code is organized by concern:

- `adapters/` runtime adapters (`xml-adapter`, `logger`)
- `core/` shared types/constants + revision/offset/XML-query policy helpers
- `engine/` router + mode implementations
  - includes focused formatting helpers (`format-paragraph-targeting`, `format-span-application`)
- `pipeline/` run-model reconciliation pipeline stages + shared list marker parser
  - includes hot-path indexed patch lookups for diff application
- `services/` table/comment/numbering services + shared package builder
- `integration/` Word API bridge only
