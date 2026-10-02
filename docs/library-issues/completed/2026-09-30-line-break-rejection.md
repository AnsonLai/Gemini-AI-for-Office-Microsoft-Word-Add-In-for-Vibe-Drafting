# Reject All reorders text around a manual line break after localized replacement

> **Resolved:** Fixed in 0.8.2; re-verified against 0.8.4 on 2026-10-02 (`2026-09-30-fidelity-reproducer.mjs` passes).

Repository: `AnsonLai/docx-redline-js`. Reproduces in installed 0.8.1 and clean
source HEAD `c4db17a8a622852c0596ec8725573ac0101d9a09` (version 0.8.1).

Filed as [library issue #4](https://github.com/AnsonLai/docx-redline-js/issues/4).

**Resolution:** Fixed in v0.8.2. The portable reproducer and mandatory consumer
regressions now pass. The observations below describe the v0.8.1 defect.

## Independent expected result

Source paragraph contains a tab, `tabbed`, a manual line break, then `Line with `:

```text
\ttabbed\nLine with 
```

Replace only `tabbed` with `aligned` using localized `replacements`.
Accepted text must be `\taligned\nLine with ` and Reject All must exactly restore
`\ttabbed\nLine with `, including tab, break and trailing space.

## Actual result

Apply and acceptance succeed. Reject All returns:

```text
\ttabbedLine\n with 
```

The unchanged `Line` moves before the source break.

**Real desktop Word confirms the defect**, version 16.0/build 16.0.20326:
Word opens the tracked package without repair. Its own Reject All yields the
same incorrect text. The engine-resolved rejected package also opens with that
incorrect text. Therefore investigation must include emitted revision markup,
rather than assuming only the library resolver is wrong.

## Reproduction and investigation

See the adjacent portable `2026-09-30-fidelity-reproducer.mjs`. It uses synthetic
source XML and package operations, without the add-in bridge. Run its structural
case against either the installed package or the repository source entrypoint.

Existing localized replacement, structural tab/field and insertion affinity
suites pass. The structural tab/field suite lacks this manual-line-break case.
Investigation entry points include `engine/reconstruction-mapper.js`
(`preserveStructuralBreaks`, run mapping), `engine/reconstruction-writer.js`
(reference-map node placement), and revision resolution. These are not
established root causes.

Required fix belongs in the library. No consumer-side workaround is proposed.
