# Reconciliation Local MCP (`docx`)

Local MCP server for creating and editing `.docx` files using the standalone OOXML reconciliation engine.

This server is intended for local automation and testing without Word JS APIs.

## What It Supports

- `docx_new`: create a minimal valid `.docx` session
- `docx_open`: open an existing `.docx` file into a session
- `docx_list_paragraphs`: inspect paragraph ids + text for targeting
- `docx_edit_paragraph`: edit one paragraph via reconciliation (`applyRedlineToOxml`)
- `docx_add_comment`: add OOXML comments anchored to text
- `docx_save_as`: write session to disk
- `docx_close`: close and release session memory

## Architecture Alignment (GSD Workflow)

This MCP server is built upon the shared **Reconciliation Core** and follows the global project documentation:
- **[SPEC.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/SPEC.md)**: Host-agnostic vision.
- **[ARCHITECTURE.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/ARCHITECTURE.md)**: Deep technical details of the core engine.
- **[ROADMAP.md](file:///c:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/ROADMAP.md)**: The splitting vision for independent sub-projects.

## Architecture Alignment

The MCP server uses the same reconciliation modules documented in:
- `src/taskpane/modules/reconciliation/ARCHITECTURE.md`

Key points:
- Edits are OOXML-first and reconciliation-driven.
- Redlines are emitted as OOXML revision markup (`w:ins`/`w:del`) when enabled.
- Numbering/comment artifacts are merged into package parts when required.
- No Word-native fallback is available in MCP mode.

If reconciliation returns `useNativeApi`, the tool errors because there is no Word runtime in MCP.

## Install

From repository root:

```bash
cd mcp/docx-server
npm install
```

Or:

```bash
npm run mcp:docx:install
```

## Run

```bash
npm start
```

The server uses stdio transport (for MCP clients).

From repository root:

```bash
npm run mcp:docx
```

## Claude Code MCP Config Example

Adjust the path for your machine:

```json
{
  "mcpServers": {
    "docx": {
      "command": "node",
      "args": [
        "[root directory]/mcp/docx-server/src/server.mjs"
      ]
    }
  }
}
```

## Typical Workflow

1. Create or open a session: `docx_new` or `docx_open`
2. Discover targets: `docx_list_paragraphs`
3. Edit by id: `docx_edit_paragraph`
4. Optionally annotate: `docx_add_comment`
5. Persist: `docx_save_as`
6. Cleanup: `docx_close`

## Redline Behavior

Session default:
- `generateRedlines` on `docx_new` / `docx_open` (default `true`)

Per-call override:
- `docx_edit_paragraph.generateRedlines`

When `generateRedlines=true`:
- Text edits are written with OOXML revisions (`w:ins`/`w:del`)
- Output is saved as tracked changes in the document package

When `generateRedlines=false`:
- Content is rewritten without revision wrappers

## Tool Notes

### `docx_edit_paragraph`

Input:
- `paragraphId` must come from `docx_list_paragraphs`
- `newText` accepts plain text and markdown hints supported by reconciliation

Output fields include:
- `changed`
- `generateRedlines`
- `sourceType` (`package`, `document`, or `fragment`)
- `updatedText`

For list-style edits, if numbering definitions are produced, the server merges `word/numbering.xml` and related package metadata automatically.

### `docx_add_comment`

Anchors comments by `textToFind` inside the target paragraph and merges:
- `word/comments.xml`
- content type overrides
- document relationships

## Current Constraints

- Paragraph-scoped editing only (`docx_edit_paragraph` edits one paragraph handle).
- No Word JS features (selection, native comments API, native list API).
- Operations that require native Word fallback are rejected in local MCP mode.

## Troubleshooting

- "Unknown paragraph id": refresh ids using `docx_list_paragraphs` after edits.
- "requires Word native API fallback": the requested transform is not fully OOXML-compatible in standalone mode.
- Save output frequently with `docx_save_as` during iterative edits.
