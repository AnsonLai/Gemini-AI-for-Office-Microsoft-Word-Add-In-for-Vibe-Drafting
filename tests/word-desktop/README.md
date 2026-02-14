# Word Desktop List Inspector

`docxjs` preview is useful for quick feedback, but Microsoft Word desktop is the source of truth for list interpretation.

Use this script to inspect how Word itself sees paragraph numbering:

```powershell
powershell -ExecutionPolicy Bypass -File tests/word-desktop/list-inspector.ps1 -DocxPath "C:\path\to\file.docx" -OnlyLists
```

Optional flags:

- `-OnlyLists`: emit only paragraphs that Word treats as list items
- `-MaxRows N`: limit output rows

Output fields (JSON):

- `paragraphIndex`
- `isList`
- `listType`
- `listId`
- `listString`
- `listValue`
- `listLevel`
- `singleList`
- `singleListTemplate`
- `style`
- `text`

Use this output to verify whether two paragraphs are in the same Word list chain (`listId`) and what numbering Word actually renders (`listString`, `listValue`, `listLevel`).

## Automated Regression

Run the end-to-end Word-grounded list regression harness:

```powershell
powershell -ExecutionPolicy Bypass -File tests/word-desktop/list-regression.ps1
```

If Word COM is unstable on a machine/session, you can still run the OOXML-side regression build/assertions only:

```powershell
node tests/word-desktop/list-regression.mjs
```

What it does:

- Copies `tests/sample_doc` into a temporary working folder
- Applies two high-risk list scenarios with reconciliation helpers:
  - non-contiguous section header conversions into ordered list numbering
  - nested insertion under archival exception (`2.2.1`)
- Validates OOXML invariants (for example no orphan `numId`, preserved nested list depth, explicit start overrides for converted headers)
- Rebuilds a `.docx`
- Inspects final list output using Word COM (`list-inspector.ps1`)
- Fails if numbering/bullets are incorrect (for example merged header numbering or missing list markers)

Artifacts are written under `tests/word-desktop/.tmp/`:

- `list-regression-output.docx`
- `list-regression-inspector.json`
