# Local DOCX MCP server

This stdio MCP server creates and edits `.docx` files through the public
`@ansonlai/docx-redline-js@0.8.3` facade. It runs in Node and does not use Word
JS APIs, a Word selection or a live Office document.

## Tools

- `docx_new`: create a session from the bundled blank DOCX; an optional title is
  inserted without tracked changes.
- `docx_open`: load a DOCX from a path relative to the server process.
- `docx_list_paragraphs`: inspect a page of paragraphs for targeting.
- `docx_edit_paragraph`: replace one inspected paragraph by its returned id.
- `docx_apply_operations`: submit an atomic operation batch.
- `docx_add_comment`: add a comment anchored to text in a paragraph.
- `docx_save_as`: explicitly serialize a session and write it to disk.
- `docx_close`: release the in-memory session.

The server stores sessions in memory. Opening or editing a session does not
change its source file; persist the current state with `docx_save_as`. Paragraph
inspection returns a one-based paragraph `index` and an `id`: the package
`paraId` when available, otherwise `idx:<index>`. Listing uses a zero-based
`start` offset and a `limit` from 1 to 500 (default 50). Refresh ids after an
edit because paragraph identities or indexes can change.

## Document lifecycle and atomicity

The service uses `openDocx`, `doc.inspect()`, `doc.applyOperations()` and
`doc.toUint8Array()`. `docx_new` uses the bundled blank template. The package
manages ZIP, comments, numbering, relationships and content types. A failed
atomic batch does not replace the session document; structured engine error
codes are returned to the MCP client.

This facade replaced four local package/targeting/XML services and removed the
MCP server's direct JSZip and `@xmldom/xmldom` dependencies. Local ownership is
now the MCP tool contract, filesystem access, session registry and blank
template; supported document inspection, edits, package relationships and
serialization are library behavior. See the [library offload
review](../../docs/library-offload-review.md) for the file-level inventory.

## Install and run

From the repository root:

```powershell
npm run mcp:docx:install
npm run mcp:docx
```

Or install and run directly:

```powershell
cd mcp/docx-server
npm install
npm start
```

Configure an MCP client with the absolute path to `mcp/docx-server/src/server.mjs`,
for example:

```json
{
  "mcpServers": {
    "docx": {
      "command": "node",
      "args": ["/absolute/path/to/AIWordPlugin/mcp/docx-server/src/server.mjs"]
    }
  }
}
```

## Typical workflow

1. Create or open a session.
2. List paragraphs and target an item using its returned id.
3. Edit one paragraph, or submit an atomic operations batch; optionally add a
   comment.
4. Save explicitly with `docx_save_as`.
5. Close the session when finished.

`generateRedlines` defaults to `true` for create/open sessions and can be
overridden on edit/batch calls. Tracked text edits use OOXML revision wrappers;
`generateRedlines=false` rewrites content without adding those wrappers. This
is package-level DOCX behavior and is not a claim about Word's live Accept All
or Reject All fidelity.

## Constraints

- `docx_edit_paragraph` edits one paragraph handle; use `docx_apply_operations`
  for a multi-operation batch.
- This server does not provide Word selection, native Word commands or
  Office.js transport.
- Refresh paragraph ids after each edit and save when you need durable output.

## Troubleshooting

- `TARGET_NOT_FOUND` / “Unknown paragraph id”: call `docx_list_paragraphs`
  again and use a current id.
- Use `docx_save_as` to write the current in-memory session; it is not
  auto-saved.
