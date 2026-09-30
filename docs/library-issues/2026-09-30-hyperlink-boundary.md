# Localized replacement absorbs adjacent plain punctuation into hyperlink

Repository: `AnsonLai/docx-redline-js`. Reproduces in installed 0.8.1 and clean
source HEAD `c4db17a8a622852c0596ec8725573ac0101d9a09` (version 0.8.1).

Filed as [library issue #3](https://github.com/AnsonLai/docx-redline-js/issues/3).

**Resolution:** Fixed in v0.8.2. The portable reproducer and mandatory consumer
regressions now pass. The observations below describe the v0.8.1 defect.

## Independent expected result

Source paragraph has a hyperlink displaying `example.org`, followed by a plain
`.` run. Replace `example.org` with `example.net` using localized `replacements`
and `insertionAffinity: { hyperlink: 'preserve', formatting: 'left' }`.

Accepted paragraph text must be `example.net.`; only `example.net` must be in
the hyperlink. The existing hyperlink address must remain unchanged.

## Actual result

Apply succeeds. Accepted paragraph text is correct, but the hyperlink displays
`example.net.`. The trailing plain punctuation becomes clickable.

**Real desktop Word confirms the defect**, version 16.0/build 16.0.20326:
Word opens without repair; its own Accept All yields hyperlink Range.Text
`example.net.`. Reject All restores the source link boundary.

The independent Word oracle also observes the defect opening engine-accepted
output. This is not merely an XML inspector disagreement or a resolver-only bug.

## Reproduction and investigation

See the adjacent portable `2026-09-30-fidelity-reproducer.mjs`. It uses synthetic
source XML and package operations, without the add-in bridge. Run its hyperlink
case against either the installed package or the repository source entrypoint.

Existing `localized_replacement_tests.mjs`, `structural_tab_field_tests.mjs` and
`insertion_affinity_tests.mjs` pass. Add a regression combining a whole-link
localized replacement with adjacent plain punctuation. Investigation entry
points include `engine/reconstruction-mapper.js` (hyperlink mapping) and
`engine/reconstruction-writer.js` (reference wrappers). These are not established
root causes.

Required fix belongs in the library. No consumer-side XML rewrite is proposed.
