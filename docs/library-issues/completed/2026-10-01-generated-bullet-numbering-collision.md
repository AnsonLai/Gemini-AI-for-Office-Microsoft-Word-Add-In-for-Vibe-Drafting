# Generated Markdown bullets reuse existing numbering IDs

> **Resolved:** Fixed in 0.8.4 (verified 2026-10-02): the runner allocates fresh IDs from the supplied `numberingXml`. `applyRedlineToOxml` needs the new `numberingXml` option, which the add-in's `edit_list` now passes from `body.getOoxml()`. Follow-up regression: [separate list operations restart numbering](../2026-10-02-separate-list-operations-restart-numbering.md).

Repository: `AnsonLai/docx-redline-js`. Reproduces in installed 0.8.3.
Not yet filed upstream.

## User impact

Golden-scenario prompt: *"rewrite the entire required disclosure provision …
Use a few bullets after the initial paragraph"*. The redline
(`replace_paragraph` with `intro\n- …\n- …\n- …`, `structuredContent: true`)
is applied in desktop Word, but the three new items are attached to `numId 1`,
which in the NDA is the existing **decimal** list of Confidential Information
examples. Word renders them as continuing numbered items (5., 6., 7.) instead
of bullets.

## Expected

New bullet paragraphs reference a numbering instance whose level 0 is
`bullet`, allocated so it does not collide with any `w:num`/`w:abstractNum` in
the supplied source `numberingXml`.

## Actual

The generated numbering part defines a bullet `abstractNum 0` and `num 1`, and
the new paragraphs reference `numId 1`. Both IDs already exist in the source
numbering (decimal), so after merging the part (or after Word's `insertOoxml`
reconciles it) the existing decimal definition wins.

```powershell
node docs/library-issues/completed/2026-10-01-generated-bullet-numbering-collision-reproducer.mjs
```

Exits 0 while the defect reproduces (minimal synthetic source: one decimal
list and one plain paragraph rewritten as an intro plus two bullets).

The same allocation path (`NumberingService` constructed without the source
numbering) is used by direct list generation in `applyRedlineToOxml`, so other
Markdown-list conversions into documents with existing lists are likely
affected too.

## Consumer state

The add-in passes the source `numberingXml`/`stylesXml` to the document
operation runner and does not remap numbering IDs itself. No workaround is
applied; the golden scenario's step 11 check ("intro paragraph then bullets")
fails until a release allocates fresh IDs.
