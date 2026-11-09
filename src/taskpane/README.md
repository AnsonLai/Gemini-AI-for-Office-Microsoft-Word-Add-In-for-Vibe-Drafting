# Word taskpane

This directory contains the in-place Word plugin, settings, chat UI, lightweight
inspect/apply harness and Office.js body transport. The add-in pins exact
`@ansonlai/docx-redline-js` (exact version in `package.json`).

`taskpane.js` keeps UI and host concerns. Portable `document-agent-loop.js`
coordinates bounded provider/tool exchanges; `tool-registry.js` declares exactly
two tools. `document-inspection-session.js` issues source-bound opaque handles;
`document-edit-intents.js` maps fully validated typed batches to library
operations. `document-tools.js` requires a persisted exact-source checkpoint,
checks freshness, commits once and verifies the host result. The compact
`document-editing-guide.js` supplies model guidance without a skill framework.

The integration folder retains Flat OPC preparation/assembly and Word
read/insert/tracking state. Native paragraph/list/table edit routes, nested AI
interpretation and orphan compatibility executors are retired. No host write
is replayed after an uncertain result. The existing Gemini transport and
IndexedDB checkpoint store remain local consumer responsibilities.

Read the [tool contract](../../docs/agentic-tool-contracts.md),
[architecture](../../docs/ARCHITECTURE.md) and
[implementation plan](../../docs/plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md)
for supported scope and recorded evidence. Body checkpoints preserve only the
transported body package. Live model and Word Online validation limits remain
explicit.

From the repository root, run `npm test`, `npm run build:dev` and `npm run build`.
Real Office.js and independent desktop Word lanes are separate; commands are in
[scripts/README.md](../../scripts/README.md).
