# Word Desktop List Inspector

`docxjs` preview is useful for quick feedback, but Microsoft Word desktop is the source of truth for list interpretation.

Use this script to inspect how Word itself sees paragraph numbering:

```powershell
powershell -ExecutionPolicy Bypass -File tests/word-desktop/list-inspector.ps1 -DocxPath "C:\path\to\file.docx" -OnlyLists
```

Optional flags:

- `-OnlyLists`: emit only paragraphs that Word treats as list items
- `-MaxRows N`: limit output rows
- `-RetryCount N`: retry COM inspection on transient Word RPC failures (default `3`)
- `-KillWordBeforeStart`: terminate existing `WINWORD` processes before each attempt (useful for unstable COM sessions)
- `-OutputPath <file>`: write JSON output directly to file (in addition to stdout)

The inspector now proactively prepares `TEMP/TMP` and Word cache folders (`INetCache\Content.Word`) before launching COM to reduce "Word could not create the work file" failures.

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

By default this harness now uses `tests/Sample NDA.docx` when present (it extracts it into `.tmp`), otherwise it falls back to `tests/sample_doc`.

Use a specific source docx:

```powershell
powershell -ExecutionPolicy Bypass -File tests/word-desktop/list-regression.ps1 -SourceDocx "tests/Sample NDA.docx"
```

Run OOXML-only regression (skips Word COM inspector, no Word popup risk):

```powershell
powershell -ExecutionPolicy Bypass -File tests/word-desktop/list-regression.ps1 -OpenXmlOnly
```

If Word shows `Word could not create the work file. Check the temp environment variable.`, run:

```powershell
powershell -ExecutionPolicy Bypass -File tests/word-desktop/repair-word-workfile.ps1
```

Then retry the inspector/regression command.

If Word COM is detected as blocked, the harness writes:

- `tests/word-desktop/.tmp/word-com-blocked.txt`

Delete that marker after fixing the local Word environment, then rerun.

If Word COM is unstable on a machine/session, you can still run the OOXML-side regression build/assertions only:

```powershell
node tests/word-desktop/list-regression.mjs
```

What it does:

- Uses extracted source OOXML (`tests/Sample NDA.docx` by default when available, otherwise `tests/sample_doc`) and copies it into a temporary working folder
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
