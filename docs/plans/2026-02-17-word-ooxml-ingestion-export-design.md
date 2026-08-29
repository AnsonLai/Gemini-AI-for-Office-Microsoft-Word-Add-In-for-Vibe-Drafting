# Word OOXML Ingestion Export Design

## Goal

Add standalone-friendly helpers that ingest Word OOXML and return:

1. Readable plain text (tags stripped, paragraph structure preserved)
2. Basic markdown (headings, bold, italics, bullets/numbered lists)

## Scope

- Target WordprocessingML (`w:*`) only.
- Keep conversion intentionally conservative:
  - Paragraph boundaries preserved.
  - Basic run formatting only (`w:b`, `w:i`).
  - Obvious heading/list detection only.
- No generic XML conversion in v1.
- No advanced markdown features (tables, links, nested list fidelity, tracked-change annotations).

## Placement

- Add a new pipeline helper module:
  - `src/taskpane/modules/reconciliation/pipeline/ingestion-export.js`
- Re-export API from:
  - `src/taskpane/modules/reconciliation/standalone.js`
  - `src/taskpane/modules/reconciliation/index.js`

This keeps logic close to existing ingestion/parsing code while exposing it through the current public surfaces.

## API Proposal

- `ingestWordOoxmlToPlainText(ooxml, options = {})`
  - Returns `{ text, warnings }`
  - `text` contains paragraph-separated readable plain text.
- `ingestWordOoxmlToMarkdown(ooxml, options = {})`
  - Returns `{ markdown, warnings }`
  - `markdown` contains line-based markdown with basic heading/list/run formatting.

Both functions are non-throwing and return warning messages when parse/input issues occur.

## Data Flow

1. Parse OOXML document.
2. Collect paragraph nodes (`w:p`) in document order.
3. For each paragraph:
  - Extract visible run text (`w:t`, `w:tab`, `w:br`, `w:cr`).
  - Detect paragraph-level cues from `w:pPr`:
    - Heading style (`w:pStyle` beginning with `Heading` + level parse)
    - List info (`w:numPr` with `w:ilvl`)
4. Render:
  - Plain text: normalize intra-paragraph spacing and join paragraphs with blank-line separation.
  - Markdown: apply paragraph prefix and run formatting wrappers (`**`, `*`) for obvious cases.

## Error Handling

- Invalid XML: return empty output + warning.
- No paragraphs:
  - Fallback to document text content with normalized whitespace for plain text.
  - For markdown, return plain-text fallback (no markdown markers) + warning.

## Test Strategy

Add `tests/standalone_ingestion_export_tests.mjs` covering:

1. Plain-text extraction removes tags and preserves paragraph readability.
2. Markdown heading conversion from `HeadingN` style.
3. Markdown bullet/numbered prefixes from `w:numPr`.
4. Markdown bold/italic conversion from run properties.
5. Parse failure behavior (no throw, warning present).

## Risks

- List marker type can be ambiguous without numbering.xml; v1 will use simple deterministic markers.
- Heading detection via style name might miss custom style mappings.
- Nested or overlapping run formatting is simplified for v1.
