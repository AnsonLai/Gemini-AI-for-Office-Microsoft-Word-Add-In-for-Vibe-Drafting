# Reconciliation Local MCP (`docx`)

Local MCP server for creating and editing `.docx` files with `@ansonlai/docx-redline-js@0.8.2`.

This server is intended for local automation and testing without Word JS APIs.

## What It Supports

- `docx_new`: create a minimal valid `.docx` session
- `docx_open`: open an existing `.docx` file into a session
- `docx_list_paragraphs`: inspect paragraph ids + text for targeting
- `docx_edit_paragraph`: edit one paragraph via `doc.applyOperations()`
- `docx_apply_operations`: apply an atomic batch of package operations
- `docx_add_comment`: add OOXML comments anchored to text
- `docx_save_as`: write session to disk
- `docx_close`: close and release session memory

## Document Lifecycle

Sessions hold a `DocxDocument` from `openDocx`. The server uses `doc.inspect()` to list and resolve paragraph handles, `doc.applyOperations()` for edits and comments, and `doc.toUint8Array()` when saving. A small blank template supplies `docx_new`; its optional title is inserted without tracked changes. The package manages ZIP, comments, numbering, relationships, and content types.

A failed atomic batch leaves the session document unchanged. Structured engine error codes are returned to the MCP client.

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
3. Edit by id: `docx_edit_paragraph`, or submit a batch with `docx_apply_operations`
4. Optionally annotate: `docx_add_comment`
5. Persist: `docx_save_as`
6. Cleanup: `docx_close`

## Redline Behavior

Session default:
- `generateRedlines` on `docx_new` / `docx_open` (default `true`)

Per-call override:
- `docx_edit_paragraph.generateRedlines` and `docx_apply_operations.generateRedlines`

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
- `sourceType` (`package`)
- `updatedText`

The package updates numbering definitions and related metadata when needed.

### `docx_add_comment`

Anchors comments by `textToFind` inside the target paragraph. The package updates comment parts and relationships.

## Current Constraints

- `docx_edit_paragraph` edits one paragraph handle. Use `docx_apply_operations` for multi-operation batches.
- No Word JS features (selection or native Word APIs).

## Troubleshooting

- "Unknown paragraph id": refresh ids using `docx_list_paragraphs` after edits.
- Save output frequently with `docx_save_as` during iterative edits.
