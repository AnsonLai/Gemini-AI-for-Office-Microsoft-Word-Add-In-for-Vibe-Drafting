# DOCX XML Harness

This harness lets you ingest a `.docx` (or an already-unzipped package folder) and inspect internal XML parts without manually zipping/unzipping.

## Commands

Run with:

```powershell
powershell -ExecutionPolicy Bypass -File tests/docx-harness/docx-xml-harness.ps1 <args>
```

### 1) Summary

```powershell
powershell -ExecutionPolicy Bypass -File tests/docx-harness/docx-xml-harness.ps1 -Action summary -InputPath "tests/Sample NDA.docx" -AsJson
```

Includes:
- part counts
- document paragraph/list stats
- numbering stats
- numbering schema-order check (`abstractNum` before `num`)

### 2) List XML/relationship parts

```powershell
powershell -ExecutionPolicy Bypass -File tests/docx-harness/docx-xml-harness.ps1 -Action list -InputPath "tests/Sample NDA.docx"
```

### 3) Extract package

```powershell
powershell -ExecutionPolicy Bypass -File tests/docx-harness/docx-xml-harness.ps1 -Action extract -InputPath "tests/Sample NDA.docx" -OutDir "tests/word-desktop/.tmp/sample-nda-extracted"
```

### 4) Show a specific part

```powershell
powershell -ExecutionPolicy Bypass -File tests/docx-harness/docx-xml-harness.ps1 -Action show -InputPath "tests/Sample NDA.docx" -Part "word/document.xml"
```

### 5) Grep across XML parts

```powershell
powershell -ExecutionPolicy Bypass -File tests/docx-harness/docx-xml-harness.ps1 -Action grep -InputPath "tests/Sample NDA.docx" -Pattern "<w:numPr>"
```

Regex mode:

```powershell
powershell -ExecutionPolicy Bypass -File tests/docx-harness/docx-xml-harness.ps1 -Action grep -InputPath "tests/Sample NDA.docx" -Pattern "w:numId\\s+w:val=\"\\d+\"" -Regex
```

### 6) XPath query

```powershell
powershell -ExecutionPolicy Bypass -File tests/docx-harness/docx-xml-harness.ps1 -Action query -InputPath "tests/Sample NDA.docx" -Part "word/numbering.xml" -XPath "//w:num"
```

## Notes

- `InputPath` accepts:
  - `.docx` file
  - unzipped docx package folder (must contain `[Content_Types].xml` and `word/document.xml`)
- For `.docx`, the harness extracts to a temp folder and cleans up automatically.
- Add `-KeepExtracted` to keep extracted files for debugging.
