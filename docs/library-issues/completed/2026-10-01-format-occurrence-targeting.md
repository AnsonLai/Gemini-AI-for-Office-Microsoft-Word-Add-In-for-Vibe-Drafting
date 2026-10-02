# `format` operations cannot select a later occurrence inside a targeted paragraph

> **Resolved:** Fixed in 0.8.4 via the new `textOccurrence` field on `format`/`character-format` (verified 2026-10-02). The add-in maps `format_text.occurrence` to `textOccurrence`; `target.occurrence` still selects a paragraph by design.

Repository: `AnsonLai/docx-redline-js`. Reproduces in installed 0.8.3.
Not yet filed upstream.

## User impact

Golden-scenario prompt: *"can you unbold BC in governing law?"*. The governing
law paragraph contains "British Columbia" twice; only the first is bold (via
the `Strong` character style). The add-in's `format_text` change maps to the
library's `format` operation. When the text to format repeats, the add-in can
only address the **first** match, so a request to format a later occurrence
must be rejected back to the model ("extend `find` to make it unique").

## Expected

A `format` operation should accept an in-paragraph occurrence for
`textToFormat` while the paragraph itself is strongly targeted (index,
`exactText`, `paragraphId`, `fingerprint`), like `replacements[].occurrence`
does for localized text edits.

## Actual

`applyFormatOperation` reads the in-paragraph occurrence from
`targetDescriptor.occurrence`, but `normalizeTargetDescriptor` and paragraph
resolution use that same field to pick the N-th paragraph matching the target
text. With `occurrence: 2`, paragraph resolution fails (`TARGET_NOT_FOUND`)
before formatting runs. Omitting it formats the first match only.

```powershell
node docs/library-issues/completed/2026-10-01-format-occurrence-targeting-reproducer.mjs
```

Exits 0 while the defect reproduces.

## Consumer state

`format_text` accepts `occurrence: 1` (the library default) and rejects other
occurrences with `format_find_ambiguous`, telling the model to extend `find`.
No consumer-side occurrence resolution or OOXML reconstruction is applied.
Suggested library change: a dedicated field such as `textOccurrence` on the
`format` operation, independent of paragraph target occurrence.
