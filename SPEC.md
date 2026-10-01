# Project Specification

## Purpose

Gemini AI for Office is a Microsoft Word add-in that lets a user ask Gemini to
review a document and apply edits through Word. The repository also contains a
browser document demo and a local MCP server for DOCX workflows.

## Runtime boundaries

The consumers pin `@ansonlai/docx-redline-js@0.8.3` exactly. The package owns
OOXML reconciliation and its public document facade. The add-in also has a
portable consumer layer at
`src/taskpane/modules/docx-redline-js-integration/consumer-core.js` for source
inspection, canonical batch preparation and package construction. That module
has no Office.js, UI or filesystem dependency. The Word adapter in the same
directory owns Word proxy reads, insertions, synchronization and tracking
state. Browser and MCP runtimes use the package facade directly for their
document lifecycles.

This is a shared repository, not a claim that every consumer uses one identical
host adapter. Keep host I/O in the consumer that owns it and reusable OOXML
behavior in the package or the tested portable boundary.

## Editing and safety contracts

- Resolve edits against inspected source text and stable paragraph identities
  where available. Refuse stale or ambiguous targets before content writes.
- Prepare migrated Word operation batches from one source snapshot and perform
  one insertion after successful preparation. An empty or refused batch has no
  insertion.
- Preserve engine receipts and distinguish preparation errors from host
  outcomes. A failed synchronization after a write attempt is indeterminate;
  do not replay a document mutation when the host may have applied it.
- Browser and MCP batches use the public facade's atomic operation contract.
  MCP save is explicit; sessions remain in memory until saved or closed.
- Word-only behavior such as selection, tracking toggles and checkpoints stays
  in the Word host layer.

## Current list migration limits

Some list operations still use established native Word paths. Canonical
`insert_list_item` routing is limited to the verified active bullet/decimal
subset and documented preconditions. Other levels/styles, tracking-off requests
and unsupported anchors continue through native capability paths; their tests
and live validation evidence must be described separately.

The broader canonical list migration remains open. The v0.8.2 reports recorded
Reject All paragraph-boundary failures and historical-numbering inspection;
the v0.8.3 release notes report fixes for those behaviors. Those historical
reports preserve the earlier evidence and do not add current host-validation
results. Plain or text-changing header-to-list conversion and list-format
changes still lack supported canonical operations. A marker-prefixed
`1. Header` to `1. Header` operation on bare `document.xml` can fail with
`RECEIPT_RECONCILIATION_FAILED`. See the [library follow-ups](docs/library-issues/README.md)
and [canonical migration follow-up](docs/plans/2026-09-30-canonical-list-migration-follow-up.md).
Do not infer complete canonical migration from passing native fallback checks.

## Provider and startup behavior

Gemini HTTP transport is centralized in `src/taskpane/modules/chat/gemini-client.js`.
Transient network/timeout and selected server/rate-limit responses receive a
bounded retry policy; authorization, invalid-request and malformed-response
failures are not retried. Transport retry never replays a document tool.

Editing support is loaded on first use in the taskpane. Concurrent first loads
share initialization; a failed import can be retried. Platform setup precedes
Word source-baseline parsing, and tool dependencies initialize once after a
successful load.

## Verification

Offline suites, real desktop Word checks, actual Office.js transport checks,
browser session checks, and observational performance measurements are separate
evidence. See [scripts/README.md](scripts/README.md) for commands and
[ARCHITECTURE.md](ARCHITECTURE.md) for component responsibilities. Current
validation limits and known library behavior are recorded in dated reports;
none of those results guarantees behavior for every Word build or DOCX.
