# Library issues closeout (2026-10-17)

`docs/library-issues/` held upstream reports for `@ansonlai/docx-redline-js`,
each with a portable reproducer. Every report was resolved by 0.10.0, so the
folder was deleted as the last cleanup item of the
[canonical list follow-up](../plans/completed/2026-09-30-canonical-list-migration-follow-up.md) (WP3).
The rule it carried (library defects get a portable reproducer, never an add-in
workaround) now lives in [`KNOWN_BUGS.md`](../plans/KNOWN_BUGS.md).

The original reports and reproducers are kept in git history. Read one with:

```powershell
git show ff68152:docs/library-issues/completed/<file>
```

The only reproducer still used by a lane, the KB-1 Word Reject All table-append
exporter, moved to
[`scripts/word-desktop/export-table-append-reject-all.mjs`](../../scripts/word-desktop/export-table-append-reject-all.mjs).
The other reproducers target fixed defects and are not part of any gate.

## Resolved reports

Release evidence is in the validation reports for
[0.8.2](../archive/validation-reports/2026-09-30-docx-redline-js-v0.8.2.md), [0.8.3](2026-09-30-docx-redline-v083.md),
[0.8.4](../archive/validation-reports/2026-10-02-docx-redline-v084.md), [0.9.0](2026-10-12-docx-redline-v090.md)
and [0.10.0](../archive/validation-reports/2026-10-17-docx-redline-v0100.md).

| Report (file in `completed/`) | Fixed in |
| --- | --- |
| Manual line-break rejection (`2026-09-30-line-break-rejection.md`) | 0.8.2 |
| Hyperlink boundary (`2026-09-30-hyperlink-boundary.md`) | 0.8.2 |
| Plain insertion rejection (`2026-09-30-list-insertion-rejection.md`, `2026-09-30-fidelity-reproducer.mjs`) | 0.8.3 |
| List-range rejection (`2026-09-30-list-range-rejection.md`) | 0.8.3 |
| All-empty list range rejection (`2026-09-30-all-empty-list-range-rejection.md`) | 0.8.3 |
| Historical list inspection (`2026-09-30-historical-list-inspection.md`) | 0.8.3 |
| Public facade list numbering (`2026-09-30-public-facade-list-numbering.md`) | 0.8.3 |
| Table append tracked formatting (`2026-10-01-table-append-tracked-formatting.md`) | 0.8.4 |
| Soft-break list conversion (`2026-10-01-soft-break-list-conversion.md`) | 0.8.4 |
| Soft-break localized replacements (`2026-10-01-soft-break-localized-replacements.md`) | 0.8.4 |
| Generated bullet numbering collision (`2026-10-01-generated-bullet-numbering-collision.md`) | 0.8.4 |
| Format occurrence targeting (`2026-10-01-format-occurrence-targeting.md`) | 0.8.4 (`textOccurrence`) |
| Existing comments delta assembly (`2026-10-02-word-existing-comments-delta-transport.md`) | Consumer fix, not a library defect (2026-10-11) |
| Separate list operations restart numbering (`2026-10-02-separate-list-operations-restart-numbering.md`) | 0.9.0 (0.8.4 regression) |
| Weak-target soft-break replacement deletes the paragraph (`2026-10-02-weak-target-soft-break-deletion.md`) | 0.9.0 |
| `generateRedlines: false` cross-author insertion (`2026-10-02-generate-redlines-false-cross-author-insertion.md`) | 0.9.0 |
| Word Reject All extra paragraph after a final table append (`2026-10-01-word-reject-table-append-paragraph.md`) | 0.9.0 at package level; live-insert residual is KB-1 |
| Canonical list capabilities (`2026-09-30-canonical-list-operations.md`, `reproduce-plain-header-conversion.mjs`) | 0.9.0 (`list-format`, `paragraph-format.numbering`) |
| Markdown pre-processor collapses literal underscore runs (`2026-10-12-table-reconciliation-underscore-collapse.md`) | 0.10.0 (KB-7) |

Where a report had a reproducer, it sits beside it under the same date prefix.
