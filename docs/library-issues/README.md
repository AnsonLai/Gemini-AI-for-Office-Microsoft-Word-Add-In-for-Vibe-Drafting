# Library follow-ups

## Current status: installed 0.8.3

The root and MCP consumers pin exact 0.8.3. The release fixes plain insertion
rejection, multi-paragraph and all-empty list rejection, historical inspection,
public-facade same-kind numbering, and list-fragment parsing/start values.
These reports remain as historical reproductions, not current open defects.
See the [0.8.3 validation](../validation-reports/2026-09-30-docx-redline-v083.md).

The remaining open requirement is [canonical list capabilities](2026-09-30-canonical-list-operations.md):
plain-to-list/header conversion and list-format changes remain unsupported by
0.8.3. The bare-XML unchanged-header receipt error also remains a documented
library limitation. Those features are planned in the library's
`docs/plans/2026-09-30-canonical-list-operations.md`, not implemented.
The August agentic plan is closed with deferred scope at the user's request.
Remaining canonical work is tracked in the fresh
[Canonical List Migration Follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md).

## Historical 0.8.2 investigation and unpublished review

These are separate upstream work items for `@ansonlai/docx-redline-js`.
The add-in and MCP pin exact 0.8.2. A registry check during the 2026-09-30
consumer wrap-up found 0.8.2 still published as latest; the local library
release and changelog also remain at 0.8.2. These Markdown packets are local
reports, not evidence that GitHub issues have been filed or resolved.

### Requirements recorded against 0.8.2

| Report | Evidence / effect |
| --- | --- |
| [Plain insertion rejection](2026-09-30-list-insertion-rejection.md) | Independent Word Reject All leaves an extra empty paragraph; plain-anchor canonical insertion stays disabled. |
| [List-range rejection](2026-09-30-list-range-rejection.md) | Independent Word Reject All merges source paragraphs; broad canonical list replacement stays disabled. |
| [Canonical list capabilities](2026-09-30-canonical-list-operations.md) | Unmarked/text-changing header conversion and list-format changes lack a verified canonical mapping. |
| [Historical list inspection](2026-09-30-historical-list-inspection.md) | Offline inspection reads historical numbering as current; the reproduced operation refuses without a write or native replay. |

The [agentic plan](../plans/completed/2026-08-29-agentic-tools-and-list-reliability.md)
remains open at the user's request until the library fixes permit full
canonical migration. Passing native Word paths do not resolve these reports.

After an upstream release, update exact pins together, rerun each diagnostic
expecting complete source restoration, run the supported matrix and actual
Office.js transport, then verify accepted/rejected exports in independent Word.
Migrate additional commands only after their numbering and targeting semantics
pass. Keep active helper modules until every caller has migrated.

### Earlier fixes in 0.8.2

- [Hyperlink boundary](2026-09-30-hyperlink-boundary.md): adjacent plain
  punctuation stays outside the hyperlink.
- [Manual line-break rejection](2026-09-30-line-break-rejection.md): Reject All
  restores text on the correct side of the break.

The fixed boundary cases remain in the consumer compatibility/fidelity gates.
No library-side patch or consumer reconstruction workaround is part of this
wrap-up.

### Historical unpublished source review (2026-09-30)

The local library now contains uncommitted fixes for historical inspection and
the two original paragraph-boundary defects. The original add-in diagnostic
cases pass 12 independent Word checks through the consumer's standalone/merge
path. These are local fixes awaiting release, not fixes in installed 0.8.2.
The [review report](../validation-reports/2026-09-30-unpublished-library-list-review.md)
records focused tests, header receipts and index conventions.

Additional upstream concerns found during that review:

- [All-empty list ranges lose source paragraphs on Reject All](2026-09-30-all-empty-list-range-rejection.md): reproduced offline in the local tree.
- [Public DOCX facade changes unrelated list numbering](2026-09-30-public-facade-list-numbering.md): independent Word failures; identical numbering output in published 0.8.2 proves this predates the new fixes.

The library's proposed `docs/plans/2026-09-30-canonical-list-operations.md`
addresses the remaining conversion/remapping feature work, but it is not
implemented. The agentic plan remains open until released fixes and verified
capabilities enable full canonical migration.
