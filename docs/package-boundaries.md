# Consumer package boundaries

The consumers pin `@ansonlai/docx-redline-js@0.8.2`. The library owns document
inspection, OOXML edits, revision resolution and package serialization.

| Surface | Responsibility | Host dependencies |
| --- | --- | --- |
| `consumer-core.js` | Immutable-source mapping, validation, operation preparation and result handling | Library and its XML adapter; no Word, UI or filesystem |
| `word-operation-runner.js` / `word-redline-runner.js` | Read live scopes, manage tracking, write prepared Flat OPC and report confirmed host outcomes | Word request context |
| Browser document session | Open/inspect, atomic batch mutation, serialize complete DOCX | Public document facade; bytes supplied by caller |
| Browser demo UI | File selection, prompt context, preview and downloads | DOM and browser file APIs |
| MCP document service | File/session management and existing tool contracts | Node filesystem and MCP transport |

The add-in integration `index.js` remains a compatibility surface with Word
exports. Non-Word consumers import the explicit portable entry point instead
of that mixed index. Existing adapter export names remain available.

## Transport and outcome differences

Portable preparation can produce a ready insertion payload. It cannot confirm
that Word accepted a write; the Word adapter sets its mutation outcome only
after host synchronization. Browser/MCP document mutation happens in memory;
download or filesystem save is a separate action.

Word uses range/body Flat OPC transport. Browser and MCP use complete DOCX
document sessions. Whole-document Word binary insertion remains a deferred
transport option, without certification from this plan.

The browser preview retains JSZip because `docx-preview` requires it. Editing
does not use JSZip to rebuild packages. The no-build page uses the library's
published browser bundle; loading its unbundled entry without dependency maps
previously failed on bare XML/ZIP dependency specifiers.

List capability limits and the separately reported library defects remain in
the [agentic list plan](plans/2026-08-29-agentic-tools-and-list-reliability.md).
Portable extraction does not certify unsupported list mutations or native
fallback fidelity. Preview output is not an independent Word oracle.
