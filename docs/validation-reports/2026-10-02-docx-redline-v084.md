# docx-redline-js 0.8.4 consumer validation (2026-10-02)

The add-in and MCP server now pin exact `@ansonlai/docx-redline-js@0.8.4`.
This report records the verification of the upgrade, the consumer changes it
enabled, and the remaining open library items. Earlier dated reports (for
example [v0.8.3](2026-09-30-docx-redline-v083.md)) remain historical records.

## Library issue verification

Each open report's portable reproducer was rerun against the installed
package (exit 0 = defect reproduces, 1 = fixed), followed by independent checks
that do not rely on the changelog.

| Report | Result |
| --- | --- |
| Soft-break list conversion | Fixed. NDA recitals convert from `\v` and `\n` host text to a 3-item `upperLetter` list; Reject All restores one paragraph with both `<w:br/>`; the implicit identical-text case stays a no-op. |
| Soft-break localized replacements | Fixed. Within-line find applies and keeps breaks; a find spanning a break is refused by design. |
| Generated bullet numbering collision | Fixed. Runner allocates fresh IDs (NDA: new `numId 28`, bullet glyph). `applyRedlineToOxml` needs the new `numberingXml` option (adopted by `edit_list`). |
| Format occurrence targeting | Fixed via new `textOccurrence` (2nd match formatted; out of range → `PATCH_SOURCE_NOT_FOUND`). |
| Table append tracked formatting | Fixed offline and in desktop Word Accept All (underline kept, 3×3 table). |
| Word Reject All after final table append | **Open.** Word now removes the table (row-level `w:trPr/w:ins`) but leaves one extra empty paragraph. |
| 2026-09-30 historical reports (7) | Still fixed on 0.8.4 (fidelity reproducer, offline re-runs; public-facade numbering relies on the dated v0.8.3 Word evidence). |
| Canonical list capabilities | **Open** (`reproduce-plain-header-conversion.mjs` still exits 0). |

New findings, documented in [`docs/library-issues/`](../library-issues/README.md):

- **Regression:** separate list operations in one batch restart numbering
  (non-adjacent header conversion yields "A.", "A."; 0.8.3 gave "A.", "B.").
  The fidelity and planner cases are skipped with a pointer to the report;
  production `convert_headers_to_list` is native and unaffected.
- **Data loss with `ok`:** a weakly targeted (index + text only) replacement
  in a soft-break paragraph deletes the whole paragraph. The add-in always
  sends strong targets and is not exposed.

Desktop Word evidence: Word 16.0, build 16.0.20430, `OpenNoRepairDialog`,
`Revisions.RejectAll()`/`AcceptAll()` on 0.8.4-built table-append packages.

## Consumer changes enabled by 0.8.4

- `format_text` maps `occurrence` to the format operation's `textOccurrence`
  (any 1-based occurrence; previously first occurrence only).
- `edit_list` reads the document's `word/numbering.xml` from `body.getOoxml()`
  and passes it as `numberingXml` to `applyRedlineToOxml`.
- `UNSUPPORTED_TABLE_FORMATTING` now refuses only the remaining 0.8.4
  limitation: changing the text of a paragraph that carries the author's own
  pending formatting while adding a table. New inline formatting with a table,
  and appends after an underlined paragraph, are allowed.
- Localized `replacements`/`modify_text` operations are planned with
  `structuredContent: false`: a literal edit in a lettered soft-break paragraph
  (`A. … / B. …`) must not be reread as a Markdown list (verified: with
  `true`, Accept All turned the recitals into three list paragraphs).

Tests updated to the fixed behavior: `ooxml_formatting_visual_tests` (tab +
break + hyperlink replacement now applies and round-trips),
`within_turn_context_refresh_tests` (0.8.4 keeps the final paragraph in place
on table append, so staleness is exercised on a new table cell),
`table_append_batch_tests` (formatting with tables preserved; text-edit
limitation still refused). New: `edit_list_numbering_tests`.

## Results

- Offline: `npm test` — 57 suites passed, 0 failed, 4 entrypoints excluded;
  suite-internal skips: the header-continuity regression (2) and the v0.5.4
  version-specific lane.
- Desktop Word golden scenario (replay, 16 steps, Word 16.0 build 16.0.20430):
  every content check passes on the final export (29/29), including steps 3
  (recitals → A/B/C list), 11 (intro + bullets) and 16 (Titan rewrite) that
  failed on 0.8.3. The only failures are the redline-author checks on steps 6,
  9 and 12, whose tools still use native Word APIs (Word stamps the Office
  user name); that migration is tracked in the canonical list follow-up.
- The live-model golden lane was not run (no provider key in this run).
