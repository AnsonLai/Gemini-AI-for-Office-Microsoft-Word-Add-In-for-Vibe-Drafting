# Browser Demo

This folder contains a no-build browser demo for the standalone OOXML reconciliation engine.

## Files

- `browser-demo/demo.html` - Static page UI.
- `browser-demo/demo.js` - Browser module script that:
  - Reads an uploaded `.docx` with JSZip.
  - Applies kitchen-sink edits using `applyRedlineToOxml`.
  - Writes back `word/document.xml` (and numbering artifacts when needed).
  - Downloads the modified `.docx`.

## What The Demo Applies

1. Text rewrite (Gemini-backed when API key is provided; fallback rewrite otherwise)
2. Format-only change (`bold + underline` markdown hints)
3. Bullets with sub-bullets
4. Markdown table transformation
5. Gemini surprise tool action (model chooses one: `comment`, `highlight`, or `redline`)

## Run It

Use any static web server from the repository root (important for relative module imports), then open:

- `http://localhost:8000/browser-demo/demo.html`

Example with Python:

```bash
python -m http.server 8000
```

## Gemini API Key

- Enter your Gemini API key in the demo page and click `Save Gemini Key`.
- The key is persisted in browser `localStorage` for this page origin.
- On run, Gemini is used twice:
  - Once to rewrite `DEMO_TEXT_TARGET`.
  - Once to choose and execute one surprise tool action (`comment`, `highlight`, or `redline`).

## Notes

- This is browser-only usage (no Node runtime required for the demo itself).
- `demo.html` includes an import map for `diff-match-patch` and `demo.js` imports JSZip from CDN.
- Do not open with `file://`; use a local server so ES modules load correctly.
