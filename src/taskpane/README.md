# Word Taskpane

This directory contains the Microsoft Word add-in UI, chat/tool orchestration
and Office.js integration. The UI is implemented in HTML, CSS and JavaScript.

## Main components

- `taskpane.js` coordinates chat, document context, checkpoints and tool
  dispatch. Editing support is loaded on first use; a failed module load can be
  retried.
- `modules/commands/agentic-tools.js` validates tool requests and coordinates
  the document operations.
- `modules/docx-redline-js-integration/consumer-core.js` contains portable
  source inspection, canonical batch preparation and result/package helpers.
- `modules/docx-redline-js-integration/word-operation-runner.js` and
  `word-ooxml.js` own Word proxy reads, insertions, synchronization and tracking
  management. The integration `index.js` preserves established exports.
- `modules/chat/gemini-client.js` provides bounded shared HTTP transport.
- `modules/storage/checkpoint-store.js` stores pre-mutation snapshots in
  IndexedDB.

The add-in pins `@ansonlai/docx-redline-js@0.8.2` for OOXML reconciliation.
Portable preparation is separated from Word I/O, but not every tool has moved
to the canonical atomic batch path. List insertion still uses tested native
fallbacks outside its verified canonical subset. See the root
[architecture](../../ARCHITECTURE.md), [tool contracts](../../docs/agentic-tool-contracts.md)
and [active list plan](../../docs/plans/2026-08-29-agentic-tools-and-list-reliability.md)
for boundaries and remaining limits.

## Local verification

From the repository root:

```powershell
npm run build:dev
npm test
```

Actual desktop Word and Office.js transport checks are separate Windows lanes;
offline tests and mocked Word proxies do not establish all native Word fidelity.
See [validation instructions](../../scripts/README.md).
