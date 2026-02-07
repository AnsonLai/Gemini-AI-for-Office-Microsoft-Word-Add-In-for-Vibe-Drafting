# Reconciliation Local MCP (`docx`)

Local MCP server for creating and editing `.docx` files with the reconciliation engine.

## Features

- `docx_new`: create a minimal valid `.docx` session
- `docx_open`: open an existing `.docx`
- `docx_list_paragraphs`: inspect paragraph ids + text
- `docx_edit_paragraph`: edit one paragraph with reconciliation logic
- `docx_add_comment`: add OOXML comments anchored in paragraph text
- `docx_save_as`: write session to disk
- `docx_close`: close session

## Install

From the repo root:

```bash
cd mcp/docx-server
npm install
```

## Run

```bash
npm start
```

The server runs over stdio (for MCP clients).

## Claude Code MCP config example

Adjust the absolute path for your machine:

```json
{
  "mcpServers": {
    "docx": {
      "command": "node",
      "args": [
        "C:/Users/Phara/Desktop/Projects/AIWordPlugin/AIWordPlugin/mcp/docx-server/src/server.mjs"
      ]
    }
  }
}
```

## Notes

- Session default redline mode:
  - set during `docx_new` or `docx_open` with `generateRedlines`
  - defaults to `true`
- Per edit override:
  - `docx_edit_paragraph.generateRedlines` overrides the session default for that call
- `docx_new` minimal package includes:
  - `[Content_Types].xml`
  - `_rels/.rels`
  - `word/document.xml`
  - `word/_rels/document.xml.rels`

