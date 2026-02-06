# Reconciliation Module

Use this folder for all OOXML reconciliation logic.

- Public API for add-in/internal: `index.js`
- Public API for standalone environments: `standalone.js`
- Architecture guide: `ARCHITECTURE.md`

Core code is organized by concern:

- `adapters/` runtime adapters (`xml-adapter`, `logger`)
- `core/` shared types/constants
- `engine/` router + mode implementations
- `pipeline/` run-model reconciliation pipeline stages
- `services/` table/comment/numbering services
- `integration/` Word API bridge only
