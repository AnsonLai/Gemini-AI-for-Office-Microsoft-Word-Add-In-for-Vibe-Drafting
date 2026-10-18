# Library follow-ups

Upstream work items for `@ansonlai/docx-redline-js`, which the project owner
maintains. Library defects are documented here with a portable reproducer
instead of being worked around in the add-in. A reproducer exits 0 while its
defect reproduces against the installed package and 1 once fixed. Resolved
reports move to [`completed/`](completed/); once every report is resolved, the
[canonical list follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md)
deletes this folder.

## Current status: installed 0.10.0 (verified 2026-10-17)

The root and MCP consumers pin exact 0.10.0. The last open report, the
underscore collapse (KB-7), is fixed in 0.10.0 (see the
[0.10.0 validation](../validation-reports/2026-10-17-docx-redline-v0100.md)).
Resolved reports are in [`completed/`](completed/) with dated resolution notes.
Their original bodies are historical records of the version they were filed against.

## Open

None. Every recorded report is resolved, so the
[canonical list follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md)
cleanup (keep needed evidence in `docs/validation-reports/`, then delete this
folder and its incoming links) is unblocked.

Consumer-side bugs, host-boundary exceptions and the Word JavaScript API
inventory are tracked in [`KNOWN_BUGS.md`](../plans/KNOWN_BUGS.md).

## Consumer transport incidents (not library defects)

None open. The existing-comments delta assembly incident is resolved and moved to
[`completed/`](completed/2026-10-02-word-existing-comments-delta-transport.md) (2026-10-11).

Consumer adoption of the released canonical list capabilities is tracked in the
[Canonical List Migration Follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md).
That plan owns the cleanup of this folder.

## After an upstream release

Update both exact pins together, rerun every open reproducer (expect exit 1; a few resolved reproducers print NOT REPRODUCED),
the offline suite, the [golden scenario](../../scripts/README.md#original-sixteen-prompt-nda-taskpane-gate)
and the affected Word lanes with independent Accept All/Reject All checks.
Move fixed reports to `completed/` with a dated resolution note and update
the links to them.

## Completed

| Report | Fixed in |
| --- | --- |
| [Manual line-break rejection](completed/2026-09-30-line-break-rejection.md) | 0.8.2 |
| [Hyperlink boundary](completed/2026-09-30-hyperlink-boundary.md) | 0.8.2 |
| [Plain insertion rejection](completed/2026-09-30-list-insertion-rejection.md) | 0.8.3 |
| [List-range rejection](completed/2026-09-30-list-range-rejection.md) | 0.8.3 |
| [All-empty list range rejection](completed/2026-09-30-all-empty-list-range-rejection.md) | 0.8.3 |
| [Historical list inspection](completed/2026-09-30-historical-list-inspection.md) | 0.8.3 |
| [Public facade list numbering](completed/2026-09-30-public-facade-list-numbering.md) | 0.8.3 |
| [Table append tracked formatting](completed/2026-10-01-table-append-tracked-formatting.md) | 0.8.4 |
| [Soft-break list conversion](completed/2026-10-01-soft-break-list-conversion.md) | 0.8.4 |
| [Soft-break localized replacements](completed/2026-10-01-soft-break-localized-replacements.md) | 0.8.4 |
| [Generated bullet numbering collision](completed/2026-10-01-generated-bullet-numbering-collision.md) | 0.8.4 |
| [Format occurrence targeting](completed/2026-10-01-format-occurrence-targeting.md) | 0.8.4 (`textOccurrence`) |
| [Existing comments delta assembly](completed/2026-10-02-word-existing-comments-delta-transport.md) | Consumer fix (not a library defect); moved 2026-10-11 |
| [Separate list operations restart numbering](completed/2026-10-02-separate-list-operations-restart-numbering.md) | 0.9.0 (0.8.4 regression) |
| [Weak-target soft-break replacement deletes the paragraph](completed/2026-10-02-weak-target-soft-break-deletion.md) | 0.9.0 |
| [`generateRedlines: false` cross-author insertion](completed/2026-10-02-generate-redlines-false-cross-author-insertion.md) | 0.9.0 |
| [Word Reject All extra paragraph after a final table append](completed/2026-10-01-word-reject-table-append-paragraph.md) | 0.9.0 at package level; live-insert residual is KB-1 |
| [Canonical list capabilities](completed/2026-09-30-canonical-list-operations.md) | 0.9.0 (`list-format`, `paragraph-format.numbering`); consumer adoption tracked in the list plan and KB-4 |
| [Markdown pre-processor collapses literal underscore runs](completed/2026-10-12-table-reconciliation-underscore-collapse.md) | 0.10.0 (KB-7) |
