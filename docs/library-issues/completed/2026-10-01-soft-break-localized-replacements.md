# Localized replacements are refused in paragraphs with soft line breaks

> **Resolved:** Fixed in 0.8.4 (verified 2026-10-02); a find spanning a break is refused by design. Related open issue: [weak-target soft-break deletion](../2026-10-02-weak-target-soft-break-deletion.md). The add-in now plans localized replacements with `structuredContent: false` so a lettered soft-break paragraph is not reread as a Markdown list.

Repository: `AnsonLai/docx-redline-js`. Reproduces in installed 0.8.3.
Not yet filed upstream.

## User impact

Golden-scenario prompt: *"rewrite 'potential business relationship or
transaction' to be a specific project, codenamed Titan"*. The phrase sits in
the recitals paragraph, whose items are separated by soft line breaks
(`<w:br/>`, because the recitals list conversion is blocked by
[the soft-break list conversion issue](2026-10-01-soft-break-list-conversion.md)).
A one-line `replacements` edit is refused with `INVALID_OPERATION`, the batch
rolls back, and the user's request silently fails.

## Expected

A localized find/replace on text within one line of a soft-break paragraph is
applied like any other localized replacement; the break is preserved.

## Actual

`services/document-operation-contract.js` rejects localized replacements when
`/\r|\n/.test(target.text)`. The accepted-view `exactText` represents `w:br`
as `"\n"`, so every paragraph containing a manual line break is excluded,
regardless of where the replacement is.

```powershell
node docs/library-issues/completed/2026-10-01-soft-break-localized-replacements-reproducer.mjs
```

Exits 0 while the defect reproduces.

## Consumer state

No workaround. The add-in's own validation accepts the change (the find exists
in the paragraph), and the engine's refusal is surfaced as
`INVALID_OPERATION`. The golden scenario's step 16 is linked to this report.
