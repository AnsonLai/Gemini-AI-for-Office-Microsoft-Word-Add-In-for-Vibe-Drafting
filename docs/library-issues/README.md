# Library follow-ups

These are separate upstream work items for `@ansonlai/docx-redline-js`.
The add-in and MCP pin exact 0.8.2. A registry check during the 2026-09-30
consumer wrap-up found 0.8.2 still published as latest; the local library
release and changelog also remain at 0.8.2. These Markdown packets are local
reports, not evidence that GitHub issues have been filed or resolved.

## Open requirements for canonical list migration

| Report | Evidence / effect |
| --- | --- |
| [Plain insertion rejection](2026-09-30-list-insertion-rejection.md) | Independent Word Reject All leaves an extra empty paragraph; plain-anchor canonical insertion stays disabled. |
| [List-range rejection](2026-09-30-list-range-rejection.md) | Independent Word Reject All merges source paragraphs; broad canonical list replacement stays disabled. |
| [Canonical list capabilities](2026-09-30-canonical-list-operations.md) | Unmarked/text-changing header conversion and list-format changes lack a verified canonical mapping. |
| [Historical list inspection](2026-09-30-historical-list-inspection.md) | Offline inspection reads historical numbering as current; the reproduced operation refuses without a write or native replay. |

The [agentic plan](../plans/2026-08-29-agentic-tools-and-list-reliability.md)
remains open at the user's request until the library fixes permit full
canonical migration. Passing native Word paths do not resolve these reports.

After an upstream release, update exact pins together, rerun each diagnostic
expecting complete source restoration, run the supported matrix and actual
Office.js transport, then verify accepted/rejected exports in independent Word.
Migrate additional commands only after their numbering and targeting semantics
pass. Keep active helper modules until every caller has migrated.

## Fixed by 0.8.2

- [Hyperlink boundary](2026-09-30-hyperlink-boundary.md): adjacent plain
  punctuation stays outside the hyperlink.
- [Manual line-break rejection](2026-09-30-line-break-rejection.md): Reject All
  restores text on the correct side of the break.

The fixed boundary cases remain in the consumer compatibility/fidelity gates.
No library-side patch or consumer reconstruction workaround is part of this
wrap-up.
