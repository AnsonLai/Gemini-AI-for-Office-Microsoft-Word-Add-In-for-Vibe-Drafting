# Installed 0.8.3 cannot convert soft-break, manually lettered text into a list

> **Resolved:** Fixed in 0.8.4 (verified 2026-10-02 on the NDA recitals with `` and `
` host text; Reject All restores both `<w:br/>`).

Repository: `AnsonLai/docx-redline-js`. Reproduces in installed 0.8.3.
Not yet filed upstream.

## User impact

Golden-scenario prompt against `tmp/Sample NDA.docx`: *"can you turn the
recitals into a proper ordered list? A, B, C..."*. The recitals are a single
paragraph whose items `A. …`, `B. …`, `C. …` are separated by soft line breaks
(`<w:br/>`). Both add-in routes fail:

- `edit_list` → `applyRedlineToOxml` on the Word range refuses with
  `TARGET_NOT_FOUND`.
- `apply_redlines` → canonical batch reports `no_change`; the loop guard then
  stops the request.

## Independent expected result

Converting the paragraph to an `upperLetter` list produces three real list
paragraphs (`First/Second/Third recital.` with `w:numPr`), while Reject All
restores the single source paragraph including its two `w:br` breaks.

## Actual result

Run the portable reproducer from the repository root:

```powershell
node docs/library-issues/completed/2026-10-01-soft-break-list-conversion-reproducer.mjs
```

It exits 0 while both defects reproduce:

1. **Soft breaks are invisible to the target check.** Word's
   `Paragraph.text` reports `w:br` as `"\v"`. `engine/oxml-engine.js` builds
   its target text from `extractFormattingFromOoxml`, whose spans contain only
   `w:t` text, so the source reads `…recital.B. Second…` and no normalization
   matches. Result: `TARGET_NOT_FOUND`.
2. **Identical manual markers are a no-op even when structure is explicit.**
   `buildListMarkdown(…, 'upperAlpha')` emits `A. …\nB. …`, byte-identical to
   the manually lettered source. The `hasTextChanges` gate compares raw strings
   and returns `noChange` before list routing, both for
   `applyRedlineToOxml(…, { explicitStructuredContent: true })` and for a
   document operation with `structuredContent: true`. The same applies to
   separate paragraphs (`1. Heading` → numbered list).

A related fidelity gap appears once (1) is fixed: deleted source runs are
serialized with a raw `"\n"` inside `w:delText` instead of `<w:br/>`, so Word's
Reject All would restore the recitals as one run-on line.

## Proposed library fix (local, unreleased)

An uncommitted fix exists in the local library checkout
(`engine/oxml-engine.js`, `pipeline/serialization.js`) with the regression
`tests/soft_break_list_conversion_tests.mjs`. On that tree:

- the target check uses break-aware spans (`buildTextSpansFromParagraphs`);
- with explicit structured content, a list target whose source paragraphs are
  not already an equivalent list counts as a change (implicit requests stay
  no-ops so an unchanged `1. Heading` is never silently auto-numbered);
- in-run `"\n"`/`"\t"` serialize back to `<w:br/>`/`<w:tab/>`.

The library's full suite passed (129/129) with the fix, and the new regression
fails on 0.8.3 with the add-in's exact `TARGET_NOT_FOUND`.

## Consumer state

The add-in already requests explicit structure for `edit_list`
(`explicitStructuredContent: true`, and `structuredContent: true` in
`list-operation-plan.js`), which is the documented library API. No consumer
workaround is applied; the scenario fails until a release containing the fix
is pinned. After release: bump the exact pins, rerun this reproducer
(expect exit 1), the list suites, and the Word lane with Accept/Reject checks.
