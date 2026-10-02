# Library follow-ups

Upstream work items for `@ansonlai/docx-redline-js`, which the project owner
maintains. Library defects are documented here with a portable reproducer
instead of being worked around in the add-in. A reproducer exits 0 while its
defect reproduces against the installed package and 1 once fixed. Resolved
reports move to [`completed/`](completed/); once every report is resolved, the
[canonical list follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md)
deletes this folder.

## Current status: installed 0.8.4 (verified 2026-10-02)

The root and MCP consumers pin exact 0.8.4. Every 2026-10-01 report except
Word's table Reject All paragraph is fixed in 0.8.4, and all earlier 0.8.2/0.8.3
fixes still hold; those reports are in [`completed/`](completed/) with dated
resolution notes.

The add-in adopted the 0.8.4 APIs: `format_text` sends `textOccurrence`,
`edit_list` passes the document's `numberingXml` to `applyRedlineToOxml`, and
the `UNSUPPORTED_TABLE_FORMATTING` refusal now covers only the remaining 0.8.4
limitation (changing the text of a paragraph that carries the author's own
pending formatting while adding a table).

## Open

| Report | Since | Effect on the add-in |
| --- | --- | --- |
| [Canonical list capabilities](2026-09-30-canonical-list-operations.md) | 0.8.2 | Plain/unmarked header conversion and list-format changes are unsupported, so those tools keep native Word routes (`reproduce-plain-header-conversion.mjs` still exits 0 on 0.8.4). |
| [Word Reject All leaves an extra paragraph after a final table append](2026-10-01-word-reject-table-append-paragraph.md) | 0.8.3 | Table removal on Reject All now works in Word, but one empty paragraph remains after the original text. |
| [Separate list operations restart numbering](2026-10-02-separate-list-operations-restart-numbering.md) | **0.8.4 regression** | Non-adjacent headers converted in one batch land in different lists ("A.", "A."). Production `convert_headers_to_list` is still native, so users are unaffected; the canonical fidelity case is skipped. |
| [Weak-target soft-break replacement deletes the paragraph](2026-10-02-weak-target-soft-break-deletion.md) | 0.8.4 | Data loss with `ok` status when targeting by index + text only. The add-in always sends strong targets and is not exposed. |

The library's proposed `docs/plans/2026-09-30-canonical-list-operations.md`
covers the canonical list capabilities; it is not implemented. Remaining
canonical work is tracked in the
[Canonical List Migration Follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md).

## After an upstream release

Update both exact pins together, rerun every open reproducer (expect exit 1),
the offline suite, the [golden scenario](../../scripts/README.md#golden-scenario-desktop-word)
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
