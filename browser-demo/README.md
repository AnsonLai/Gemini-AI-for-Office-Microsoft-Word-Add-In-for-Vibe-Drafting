# Browser Demo

No-build browser demo using the repository's pinned `@ansonlai/docx-redline-js@0.8.2`
package and public DOCX facade.

It demonstrates end-to-end `.docx` mutation in the browser, including text redlines, formatting, list/table transforms, comments, and highlights.
It also renders a live side-by-side preview using `docxjs` so tracked changes can be reviewed immediately.

> [!IMPORTANT]
> The `docxjs` preview is not authoritative for Word list behavior. Validate list/numbering correctness against Microsoft Word desktop when debugging reconciliation output.

## Modes

### Chat Mode (Primary)

Interactive **contract redline review** powered by Gemini AI:

1. Upload a `.docx` document
2. Choose edit mode in the top toggle (`Redlines` or `Direct edits`)
3. (Optional) Drag/drop reference `.docx` files into the left-side library panel
4. Enter your Gemini API key
5. Type a review instruction (e.g., "Review this contract and flag clauses that deviate from market standards")
6. Gemini analyzes the full document text and returns structured operations (redlines, comments, highlights)
7. Operations are applied to the document OOXML
8. Download the marked-up result
9. Continue the conversation for follow-up reviews

As operations are applied, the right-side preview pane is refreshed from the in-memory `.docx` package.

The chat supports multi-turn conversation — Gemini retains context from previous turns.

### Edit Mode Toggle

- `Redlines`: content edits are applied as tracked changes (insertions/deletions).
- `Direct edits`: content edits are applied directly (no tracked insert/delete markup).
- Switching mode resets chat turn history so Gemini instructions stay aligned with the selected mode.

### Reference Library Panel

- Add multiple `.docx` documents via drag/drop or the `Add` button in the left panel.
- Each library file is converted to plain text in-browser and shown in the panel list.
- Library text is included in every chat request as supplemental context.
- Updating/clearing/removing library files resets chat history so context stays consistent.
- Use the panel chevron (`«` / `»`) to collapse or expand the library sidebar.
- Use each source checkbox to include/exclude that document from the next chat message context.

### Gemini Request Shape

- Request payload construction is centralized in `buildGeminiRequestPayload(...)` in `browser-demo/demo.js`.
- System prompt assembly is in `buildSystemInstruction(...)` in `browser-demo/demo.js`.
- Request transport uses `src/taskpane/modules/chat/gemini-client.js`, shared with the add-in and evaluations. It retries transient HTTP/network failures at most three times, with jitter, and bounds each attempt to 90 seconds. It does not execute or replay edits.
- The latest request metadata is exposed in the browser console as:
  - `window.__BROWSER_DEMO_LAST_GEMINI_REQUEST__`
- This debug object contains:
  - masked endpoint
  - request method and headers
  - selected vs total library source counts
  - formatting query and match counts
- Document text, request bodies, system prompts and API keys are excluded from this diagnostic snapshot.

### Kitchen-Sink Mode (Legacy)

One-click demo that applies a fixed set of operations to marker paragraphs:

1. Text rewrite on `DEMO TEXT TARGET` (Gemini-backed when key is present; deterministic fallback otherwise)
2. Format-only change on `DEMO FORMAT TARGET` (markdown hints)
3. List generation on `DEMO LIST TARGET` (with numbering artifact handling)
4. Table transformation on `DEMO TABLE TARGET`
5. Gemini surprise tool action (`comment`, `highlight`, or `redline`)

Missing markers are seeded as separate paragraphs through the public API,
with every source paragraph checked for preservation. Existing underscore
markers such as `DEMO_TEXT_TARGET` remain accepted aliases.

## Files and local startup

- `demo.html`: static UI and import map to the library's published browser bundle.
- `demo.js`: file selection, model prompts, preview and download UI.
- `document-session.js`: public `openDocx` document lifecycle, immutable batch targets and atomic mutation.
- JSZip remains a preview dependency of `docx-preview`; document editing and save use the library facade. The import map resolves the package from the root `node_modules` directory, while `docx-preview` is loaded from its pinned CDN URL.

Install root dependencies, serve the repository root, then open
`http://localhost:8000/browser-demo/demo.html`:

```powershell
npm install
python -m http.server 8000 --bind 127.0.0.1
```

Do not use `file://`. Both root and MCP pin exact library version 0.8.2.

## Gemini API Key

- Enter key in the top bar and click `Save Key`
- Stored in browser `localStorage` for that origin
- Used for:
  - Chat mode: multi-turn contract review and analysis
  - Kitchen-Sink mode: text rewrite suggestion + surprise tool action

If Gemini is unavailable, the kitchen-sink demo continues with fallback behavior. Chat mode requires a valid API key.

## Document pipeline

1. Upload DOCX bytes and open a document session with `openDocx`.
2. Inspect paragraphs and build a read-only Markdown projection for formatting context.
3. Send the user's instruction to Gemini using the shared request client.
4. Pass the returned operation array to one atomic document-facade batch.
5. On failure, retain the original working document and show the error. No partial batch is committed.
6. Serialize with `toUint8Array`, refresh preview and enable download.

`targetRef` uses the initial source of the batch; earlier structural operations do not shift later targets. Localized `replacements` can specify an occurrence. Full paragraph/range redlines remain supported. Targeting and structural capabilities belong to the installed library, including its documented refusals.

Direct mode disables generation of new tracked changes. It does not accept existing revisions from other authors. Comments and highlights are annotations in either edit mode.

Kitchen-sink mode also uses the document lifecycle for marker seeding and its operation batch. Its deterministic fallback needs no Gemini key. Gemini calls remain optional for that mode.

## Verification and limits

The browser preview is not a Word fidelity oracle. List capability gaps and separately reproduced Reject All defects remain in the [agentic list plan](../docs/plans/2026-08-29-agentic-tools-and-list-reliability.md). This migration does not certify additional list shapes.

Offline tests cover session open/inspect, localized edits, direct edits, comment preservation, atomic failure and serialize/reopen. The local browser validation page at `http://localhost:8000/scripts/browser-document-validation.html` exercises the same session in a real Chromium browser without a model request. Its seven grouped checks cover Word-authored source inspection, tracked edit and serialize/reopen, preserved package parts, Accept All/Reject All, mixed-batch rollback, direct mode preserving another author's revisions, and comment insertion with an existing thread.

Cross-host parity compares semantic outcomes and preserved parts through browser, MCP and the production Word adapter with mocked Word transport. It does not substitute for the actual Office.js collector or independent desktop Word oracle. The browser preview and browser pass report are not evidence of Word rendering, native list behavior or whole-document Office.js transport.

See [consumer package boundaries](../docs/package-boundaries.md) for host responsibilities. Whole-document Word binary transport remains deferred.

## Troubleshooting

- Upload a DOCX before entering a chat instruction.
- A target refusal leaves the batch unapplied; inspect the engine log for its error code.
- Refresh after updating the demo or installed package. The import map must resolve the browser bundle.
- Preview failure does not establish that the DOCX edit failed; download/reopen and independent Word checks are separate gates.
- Chat requires a Gemini key. The deterministic kitchen-sink fallback and local validation page do not.
- Conversation history is in memory and resets on refresh or context/mode changes.
