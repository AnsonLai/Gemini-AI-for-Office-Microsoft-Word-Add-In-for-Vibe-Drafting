# Browser Demo

No-build browser demo for the standalone OOXML reconciliation engine.

It demonstrates end-to-end `.docx` mutation in the browser, including text redlines, formatting, list/table transforms, comments, and highlights.
It also renders a live side-by-side preview using `docxjs` so tracked changes can be reviewed immediately.

> [!IMPORTANT]
> The `docxjs` preview is not authoritative for Word list behavior. Validate list/numbering correctness against Microsoft Word desktop when debugging reconciliation output.

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

As operations are applied, the right-side preview pane is refreshed from the in-memory `.docx` package.

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
  - includes right-side `docxjs` (`docx-preview`) live preview pane
- `browser-demo/demo.js`: browser module pipeline (chat engine + OOXML operations)
  - renders preview with `renderChanges` enabled so insertions/deletions are visible

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
  2. when `targetRef` drifts after earlier same-turn structural edits, strict text rematch using the turn-start paragraph snapshot (preferring the original table/body context)
  3. strict text match
  4. fuzzy text match fallback
- Redline diffing uses the resolved paragraph's current text, which reduces failures when model-provided `target` text drifts slightly.
- Target resolution is delegated to shared reconciliation core helpers (exported via `standalone.js`), including turn-snapshot drift correction (`buildTargetReferenceSnapshot`, `resolveTargetParagraphWithSnapshot`), so non-demo consumers can reuse the same behavior.

### Table Structure Edits (Chat)

- When a redline target paragraph is inside a table cell and `modified` is markdown table text, the demo applies reconciliation at **table scope** (the containing `w:tbl`) instead of single-paragraph scope.
- This is required for structural updates like adding/removing/reordering rows.
- For these edits, Gemini should return the **full target table** as markdown in `modified` and include accurate `targetRef`.
- Table-scope detection uses shared reconciliation core helpers (exported via `standalone.js`) so other projects can reuse the same targeting behavior.
- If Gemini returns multiline cell text (for example `Title:\nDate:`) instead of full markdown table, the demo now attempts a shared-core heuristic that synthesizes a full markdown table and applies reconciliation at table scope.
- For symmetric two-column signature rows (same label in both columns, e.g. `Title:`), synthesized insertion rows are mirrored across both columns (e.g. `Date:` on both sides).

### List Structure Edits (Chat)

- When a target paragraph is part of an OOXML numbered/bulleted list and a redline contains multiline list content, the demo can promote the edit to **contiguous list-block scope**.
- This helps "insert item between N and N+1" requests by preserving surrounding list items in the same list block, instead of rewriting only one paragraph.
- For supported middle-insert patterns, the demo now uses an **insertion-only** heuristic first, adding only new list paragraph(s) as redlines and leaving existing items untouched.
- List heuristics use shared reconciliation core helpers exported via `standalone.js` (`planListInsertionOnlyEdit`, `synthesizeExpandedListScopeEdit`).
- List numbering payloads are remapped to fresh `numId`/`abstractNumId` values and merged into existing `word/numbering.xml`, preventing accidental continuation or style collision with distant lists.
- Composite list markers are normalized in list generation (for example `- A. Item` becomes `A. Item`) so ordered list style can be inferred correctly and marker text is not duplicated in content.
- List conversion now bypasses text-only no-op short-circuits when loose list markers are present, so existing marker-prefixed plain text (`A.`, `B.`, `C.`) can still be converted into true Word list structure.
- If a single-paragraph redline is a no-op but `modified` is a one-line list marker (for example `1. DEFINITION`), the demo now applies shared standalone fallback helpers (`buildSingleLineListStructuralFallbackPlan`, `executeSingleLineListStructuralFallback`) to force structural list conversion and strip manual marker text into true list numbering.
- In chat mode, this fallback is intentionally limited to non-list paragraphs (`allowExistingList: false`) so it cannot accidentally rebind existing list items.
- For header conversions, single-line fallback is intended for non-list header paragraphs to avoid disturbing existing list chains.
- For explicit numeric single-line markers (`1.`, `2.`, `3.`, ...), fallback now creates a dedicated multilevel decimal-outline definition per converted paragraph (`%1.`, `%1.%2.`, `%1.%2.%3.`, ...), applies a start override, and uses a fresh `numId`; this prevents accidental continuation into unrelated lists and preserves nested numbering markers in Word.
- Explicit composite ordered markers in multiline insertion edits (for example `2.2.1`) now map to deeper OOXML list levels during insertion-only planning, so requested sub-sub items are inserted at the intended depth instead of becoming same-level siblings.
- When multiline insertion under a nested ordered item is ambiguous (for example model emits bullet marker text), insertion-only heuristics promote inserted lines one level deeper to preserve sub-item intent.
- Redundant manual list prefixes in list-item text (for example `- Item`, `2.1. Item`, or `2.1. - Item`) are stripped during list-item reconciliation so marker text is not duplicated in visible content.
- Browser demo fallback applies explicit starts with num-level `w:lvlOverride/w:startOverride` only (abstract-level start override disabled) to avoid renderer-wide renumber side effects across unrelated lists.
- For repeated single-line conversions without explicit numeric starts, the demo reuses a shared generated `numId` per list style so numbering continues across non-contiguous targets.

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
- Demo version in log does not match expected (`v2026-02-14-chat-docx-preview-16`): force refresh the page (`Ctrl+F5`) to bypass cached module URLs.
- Need Word-grounded list diagnostics: run `tests/word-desktop/list-inspector.ps1` against the generated `.docx` and compare `listId`/`listLevel`/`listValue` for the affected paragraphs.
- Validation error about numbering/comments: check whether package relationships or content types were removed by prior tooling.
- No Gemini output: verify API key and network access; kitchen-sink fallback path should still run.
- Chat input disabled: upload a `.docx` file first to enable the chat.
