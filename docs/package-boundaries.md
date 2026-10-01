# Consumer package boundaries

The consumers pin `@ansonlai/docx-redline-js@0.8.3`. The library owns document
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

The v0.8.2 reports preserve earlier Reject All and historical-property
findings; v0.8.3 validation passes 12 independent Word checks on the two
former public-facade Reject All cases. The release notes describe those fixes.
Remaining canonical limits are plain or text-changing header-to-list
conversion and list-format changes. A marker-prefixed `1. Header` no-op on bare
`document.xml` can fail with `RECEIPT_RECONCILIATION_FAILED`. The migration
status is in the [canonical migration follow-up](plans/2026-09-30-canonical-list-migration-follow-up.md).
Portable extraction does not certify unsupported list mutations or native
fallback fidelity. Preview output is not an independent Word oracle.
