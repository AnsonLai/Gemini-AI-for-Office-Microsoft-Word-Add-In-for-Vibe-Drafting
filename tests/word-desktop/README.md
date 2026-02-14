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
