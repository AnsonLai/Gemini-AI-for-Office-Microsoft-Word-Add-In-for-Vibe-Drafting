# Literal underscore runs are collapsed by the Markdown pre-processor

> **Resolved in 0.10.0 (2026-10-17):** Fixed in `@ansonlai/docx-redline-js@0.10.0`. Underscore emphasis now follows CommonMark flanking rules, and a caller's `structuredContent: false` keeps inline Markdown literal. The reproducer prints "NOT REPRODUCED" and exits 1. The add-in and MCP pin exact 0.10.0, and the "Add Date Row" known-defect gate is removed. The body below is the historical 0.9.0 report. See the [0.10.0 validation](../../validation-reports/2026-10-17-docx-redline-v0100.md).

Repository: `AnsonLai/docx-redline-js`. Reproduces in installed 0.9.0 and in
0.8.4 (checked 2026-10-12 with a scratch install outside the repo). Not yet
filed upstream.

## Impact

Silent text corruption with a success status. Underscore signature lines
(`_________________________`) are parsed as Markdown bold (`__x__`) and
italic (`_x_`) and shrunk in the Accept All text. The library reports
`status: 'ok'` and nothing warns.

- A `table-reconciliation` whose modified Markdown repeats the table's own
  cells verbatim (a no-change reconcile) turns a 25-character underscore cell
  into `_`. That includes cells the edit did not touch, and any new cell. The
  surviving `_` is also marked bold.
- A lone `_` in one cell can swallow text across cells: `| _ | x_y |` yields a
  cell `xy`, because the italic pattern runs across the `|` separator.
- `*****` is collapsed the same way (to `*`).
- Backslash escapes (`\_\_\_\_\_`) do not help.
- The same parser also changes a plain paragraph rewrite with
  `structuredContent: false`, which the README describes as treating the text
  "strictly as plain text". Appending words to `Signed: _________________________`
  gives `Signed: _____ (print name)`. `__Schedule A__` loses its underscores
  and becomes bold.

## Expected

Text that the caller did not mark up keeps its literal characters. Either
treat a run of underscores (or asterisks) that has no content between the
delimiters as literal text, or honour `structuredContent: false` /
`literalTextEdit` by skipping inline Markdown extraction. A reconcile whose
modified text equals the table's current text must be a no-op.

## Actual

```powershell
node docs/library-issues/completed/2026-10-12-table-reconciliation-underscore-collapse-reproducer.mjs
```

Prints the Accept All cells for a no-change reconcile (`["_","_"]` against
25-underscore cells) and the rewritten paragraph
(`"Signed: _____ (print name)"`). Exits 0 while the defect reproduces.

## Cause

`pipeline/markdown-processor.js`, `preprocessMarkdown`, pattern list
`MARKDOWN_PATTERNS`. Two patterns match underscore runs:

- `{ regex: /__(.+?)__/, format: { bold: true } }`. Five underscores match as
  `__` + `_` + `__`, so the run becomes one bold `_`. Longer runs shrink by
  four characters per match.
- `{ regex: /_(?!_)(.+?)_(?!_)/, format: { italic: true } }`. It has no word
  boundary or same-line/same-cell restriction.

The table route runs the pre-processor twice, so a 25-character run goes
25 to 5 to 1:

1. `engine/oxml-engine.js` line 382 (`const { cleanText: cleanModifiedText, formatHints } = preprocessMarkdown(sanitizedText)`)
   cleans the whole Markdown table, and the result goes to
   `applyTableReconciliation` (`engine/table-mode.js`, called at line 573).
2. `services/table-reconciliation.js` calls `preprocessMarkdown` again per cell
   (`generateTableOoxml` line 67, `serializeVirtualGridToOoxml` line 242,
   `reconcileCellContent` line 368, `reconcileMultiBlockCellContent` lines 343
   and 352).

The first pass runs even when `structuredContent: false` or `literalTextEdit`
is set, which is why the paragraph rewrite is affected too.

## Consumer exposure

See `KNOWN_BUGS.md` KB-7. The add-in cannot reach the table route from its
retained intents, but `rewrite_paragraph` reaches the paragraph-scope
variant. A localized `replace` is refused by the library with
`PATCH_ROUNDTRIP_MISMATCH` instead of corrupting text.
