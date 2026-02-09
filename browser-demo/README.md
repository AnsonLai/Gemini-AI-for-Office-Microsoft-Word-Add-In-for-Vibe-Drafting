# Browser Demo

No-build browser demo for the standalone OOXML reconciliation engine.

It demonstrates end-to-end `.docx` mutation in the browser, including text redlines, formatting, list/table transforms, comments, and highlights.

## Files

- `browser-demo/demo.html`: static UI
- `browser-demo/demo.js`: browser module pipeline

The demo imports reconciliation APIs from:
- `src/taskpane/modules/reconciliation/standalone.js`

## Architecture Alignment

The demo follows the same reconciliation architecture used by the add-in and documented in:
- `src/taskpane/modules/reconciliation/ARCHITECTURE.md`

Operationally:
- Uses OOXML-first reconciliation (`applyRedlineToOxml`) for paragraph edits.
- Preserves track changes via OOXML revision tags (`w:ins`/`w:del`).
- Handles package-level artifacts for numbering/comments when needed.
- Validates resulting package structure before download.

## What The Demo Applies

1. Text rewrite on `DEMO_TEXT_TARGET` (Gemini-backed when key is present; deterministic fallback otherwise)
2. Format-only change on `DEMO FORMAT TARGET` (markdown hints)
3. List generation on `DEMO_LIST_TARGET` (with numbering artifact handling)
4. Table transformation on `DEMO_TABLE_TARGET`
5. Gemini surprise tool action (`comment`, `highlight`, or `redline`)

The demo ensures marker paragraphs exist before running.

## Run

Serve repository root with any static server (required for ES modules), then open:

- `http://localhost:8000/browser-demo/demo.html`

Example:

```bash
python -m http.server 8000
```

Do not use `file://`.

## Gemini API Key

- Enter key in the page and click `Save Gemini Key`
- Stored in browser `localStorage` for that origin
- Used for:
  - text rewrite suggestion
  - surprise tool action selection

If Gemini is unavailable, the demo continues with fallback behavior.

## Pipeline Overview

1. Read uploaded `.docx` using JSZip
2. Parse `word/document.xml`
3. Apply operations by exact target marker paragraph
4. Extract replacement nodes from reconciliation output (`package`, `document`, or fragment)
5. Merge numbering/comments package artifacts if emitted
6. Validate resulting package:
   - `word/document.xml` well-formed
   - `w:sectPr` body ordering valid
   - numbering/comments parts + relationships + content types consistent
7. Download mutated `.docx`

## Important Behavior Notes

- Redlines are generated via OOXML, not Word runtime state.
- List/table operations can emit package output and extra parts.
- Numbering merge is additive when missing; existing numbering part is preserved.
- Comments merge into `word/comments.xml` and related package metadata.

## Known Limits

- Targeting uses exact marker paragraph text, not semantic search.
- Demo applies a fixed operation sequence; it is not a general-purpose editor UI.
- Browser runtime constraints apply (memory/file size/network for Gemini).

## Troubleshooting

- "Target paragraph not found": verify marker paragraphs still exist in the uploaded doc.
- Validation error about numbering/comments: check whether package relationships or content types were removed by prior tooling.
- No Gemini output: verify API key and network access; fallback path should still run.
