# Browser Demo

No-build browser demo for the standalone OOXML reconciliation engine.

It demonstrates end-to-end `.docx` mutation in the browser, including text redlines, formatting, list/table transforms, comments, and highlights.

## Modes

### Chat Mode (Primary)

Interactive **contract redline review** powered by Gemini AI:

1. Upload a `.docx` document
2. Enter your Gemini API key
3. Type a review instruction (e.g., "Review this contract and flag clauses that deviate from market standards")
4. Gemini analyzes the full document text and returns structured operations (redlines, comments, highlights)
5. Operations are applied to the document OOXML
6. Download the marked-up result
7. Continue the conversation for follow-up reviews

The chat supports multi-turn conversation — Gemini retains context from previous turns.

### Kitchen-Sink Mode (Legacy)

One-click demo that applies a fixed set of operations to marker paragraphs:

1. Text rewrite on `DEMO_TEXT_TARGET` (Gemini-backed when key is present; deterministic fallback otherwise)
2. Format-only change on `DEMO FORMAT TARGET` (markdown hints)
3. List generation on `DEMO_LIST_TARGET` (with numbering artifact handling)
4. Table transformation on `DEMO_TABLE_TARGET`
5. Gemini surprise tool action (`comment`, `highlight`, or `redline`)

## Files

- `browser-demo/demo.html`: static UI (chat layout + styles)
- `browser-demo/demo.js`: browser module pipeline (chat engine + OOXML operations)

The demo imports reconciliation APIs from:
- `src/taskpane/modules/reconciliation/standalone.js`
  - including shared paragraph-targeting helpers (`resolveTargetParagraph`, marker parsing, strict/fuzzy matching)

## Architecture Alignment

The demo follows the same reconciliation architecture used by the add-in and documented in:
- `src/taskpane/modules/reconciliation/ARCHITECTURE.md`

Operationally:
- Uses OOXML-first reconciliation (`applyRedlineToOxml`) for paragraph edits.
- Preserves track changes via OOXML revision tags (`w:ins`/`w:del`).
- Handles package-level artifacts for numbering/comments when needed.
- Validates resulting package structure before download.

## Run

Serve repository root with any static server (required for ES modules), then open:

- `http://localhost:8000/browser-demo/demo.html`

Example:

```bash
python -m http.server 8000
```

Do not use `file://`.

## Gemini API Key

- Enter key in the top bar and click `Save Key`
- Stored in browser `localStorage` for that origin
- Used for:
  - Chat mode: multi-turn contract review and analysis
  - Kitchen-Sink mode: text rewrite suggestion + surprise tool action

If Gemini is unavailable, the kitchen-sink demo continues with fallback behavior. Chat mode requires a valid API key.

## Chat Pipeline

1. Upload `.docx` → read with JSZip
2. Extract all paragraph text from `word/document.xml`
3. Build Gemini system instruction with full document listing
4. User types review instruction → sent as multi-turn chat
5. Gemini responds with explanation + JSON array of operations (`redline`/`comment`/`highlight`)
6. Operations applied paragraph-by-paragraph using reconciliation engine
7. Numbering/comments package artifacts merged if emitted
8. Output validated; download button enabled
9. Paragraph listing refreshed for next turn

### Chat Targeting

- Chat operations support `targetRef` (for example `P12`) in addition to `target` text.
- The demo resolves targets in this order:
  1. `targetRef` paragraph index (when provided)
  2. strict text match
  3. fuzzy text match fallback
- Redline diffing uses the resolved paragraph's current text, which reduces failures when model-provided `target` text drifts slightly.
- Target resolution is delegated to shared reconciliation core helpers (exported via `standalone.js`) so non-demo consumers can reuse the same behavior.

## Kitchen-Sink Pipeline

1. Read uploaded `.docx` using JSZip
2. Parse `word/document.xml`
3. Apply operations by exact target marker paragraph
4. Extract replacement nodes from reconciliation output (`package`, `document`, or fragment)
5. Merge numbering/comments package artifacts if emitted
6. Validate resulting package
7. Download mutated `.docx`

## Important Behavior Notes

- Redlines are generated via OOXML, not Word runtime state.
- Chat mode sends the full document text to Gemini for analysis.
- Multi-turn chat history is maintained in-memory (lost on page refresh).
- List/table operations can emit package output and extra parts.
- Numbering merge is additive when missing; existing numbering part is preserved.
- Comments merge into `word/comments.xml` and related package metadata.

## Known Limits

- Chat mode may still skip some format-only operations when the reconciliation engine requests native Word API fallback (`useNativeApi`) and no OOXML payload is returned.
- Kitchen-sink mode targeting uses exact marker paragraph text, not semantic search.
- Browser runtime constraints apply (memory/file size/network for Gemini).
- Document text extraction is plain-text only (no formatting or style metadata sent to Gemini).
- Multi-turn history is in-memory; refreshing the page resets the conversation.

## Troubleshooting

- "Target paragraph not found": Gemini may have slightly modified the paragraph text when referencing it. Check the engine log for details.
- "Format-only fallback requires native Word API": this operation was a pure formatting change where the engine could not safely localize spans in OOXML. The browser demo skips it; use a more specific target (`targetRef` + exact paragraph text) or run through the add-in Word path.
- Demo version in log does not match expected (`v2026-02-12-chat-targeting-core`): force refresh the page (`Ctrl+F5`) to bypass cached module URLs.
- Validation error about numbering/comments: check whether package relationships or content types were removed by prior tooling.
- No Gemini output: verify API key and network access; kitchen-sink fallback path should still run.
- Chat input disabled: upload a `.docx` file first to enable the chat.
