# Word Taskpane

This directory contains the Microsoft Word add-in UI, chat/tool orchestration
and Office.js integration. The UI is implemented in HTML, CSS and JavaScript.

## Main components

- `taskpane.js` coordinates chat, document context, checkpoints and tool
  dispatch. Editing support is loaded on first use; a failed module load can be
  retried.
- `modules/commands/agentic-tools.js` validates tool requests and coordinates
  the document operations.
- `modules/docx-redline-js-integration/consumer-core.js` contains add-in
  mapping, source validation and the local Flat OPC preparation/assembly
  bridge; supported DOCX operations and reconciliation run through the pinned
  library.
- `modules/docx-redline-js-integration/word-operation-runner.js` and
  `word-ooxml.js` own Word proxy reads, insertions, synchronization and tracking
  management; `redline-plan.js` maps validated AI changes to source-bound
  operations. The integration `index.js` preserves established exports.
- `modules/chat/gemini-client.js` provides bounded shared HTTP transport.
- `modules/storage/checkpoint-store.js` stores pre-mutation snapshots in
  IndexedDB.

The add-in pins `@ansonlai/docx-redline-js@0.8.3` for OOXML reconciliation.
Portable preparation is separated from Word I/O, but not every tool has moved
to the canonical atomic batch path. The old iterative Word redline loop,
native insertion recovery and manual table synthesis were removed from the
migrated redline route, which now prepares one atomic library batch before the
Word adapter's single confirmed insertion. `insert_list_item` is canonical
only for the verified tracking-on bullet/decimal subset at source/resolved
levels 0–1. General `edit_list` and header conversion, deeper/other-style
insertion, tracking-off requests and unsupported outdent contexts remain on
native or established legacy paths. The local Flat OPC bridge and operation
mapping are not library code. Shared Gemini transport/outcomes, test lanes and
deferred startup are consumer reliability/packaging work. See the [library
offload review](../../docs/library-offload-review.md). Current v0.8.3 checks cover
12 supported list cases (48 Office.js checks) and five native routes (20
Office.js checks); the independent Word oracle reports 92 applicable checks
passed across 17 exports, zero failed, and five not applicable. `npm test`
passes 51 suites with zero failures and four exclusions. See the [v0.8.3
report](../../docs/validation-reports/2026-09-30-docx-redline-v083.md). See the root
[architecture](../../ARCHITECTURE.md), [tool contracts](../../docs/agentic-tool-contracts.md)
and [canonical migration follow-up](../../docs/plans/2026-09-30-canonical-list-migration-follow-up.md)
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
