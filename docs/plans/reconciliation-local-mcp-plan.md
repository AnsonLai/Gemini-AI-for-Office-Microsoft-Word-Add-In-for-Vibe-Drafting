# Local MCP Plan: Reconciliation Engine for `.docx` Editing

## Goal

Build a local MCP server (for clients like Claude Code) that can open a `.docx`, apply targeted edits through the reconciliation engine, and write a valid updated `.docx` back to disk.

## Why this is feasible now

The reconciliation module is already close to MCP-ready:

- Standalone entrypoint exists: `src/taskpane/modules/reconciliation/standalone.js`
- No Word API requirement in standalone path
- Node XML provider injection already supported (`configureXmlProvider`)
- Core edit primitives exist:
  - text + formatting redlines: `applyRedlineToOxml`
  - comments: `injectCommentsIntoOoxml`
  - parsing/ingestion helpers for targeting

## Scope (V1)

- Local-only execution over stdio transport
- Single-user session model
- `.docx` create/open/edit/save operations
- Paragraph-targeted edits and comment insertion
- Track changes behavior configurable (enabled/disabled) per session and per edit call

Out of scope for V1:

- Multi-user concurrency
- Full document semantic rewrite in one call
- Cloud storage integrations

## Proposed architecture

1. `mcp/docx-server/src/server.mjs`
- MCP server bootstrap (stdio transport)
- Tool registration + input validation

2. `mcp/docx-server/src/services/docx-package-service.mjs`
- Load/save zip (`word/document.xml`, optional `word/comments.xml`, rels)
- Ensure package part updates are consistent
- Build new minimal valid `.docx` packages for `docx_new`

3. `mcp/docx-server/src/services/docx-session-store.mjs`
- In-memory session map keyed by `sessionId`
- Stores parsed XML, original path, dirty state, and default `generateRedlines`

4. `mcp/docx-server/src/services/reconciliation-service.mjs`
- Thin adapter around existing module exports
- Configures XML provider once for Node runtime
- Normalizes engine output into document-level updates

5. `mcp/docx-server/src/services/paragraph-targeting-service.mjs`
- Build stable paragraph handles (`w14:paraId` when present, fallback generated ids)
- Resolve paragraph by id/index/text signature

## MCP tools (V1)

1. `docx_new`
- Input: optional `outputPath`, optional `title`, optional `generateRedlines` (default `true`)
- Behavior: creates a new in-memory minimal valid `.docx` package and opens a session; optionally writes it immediately
- Output: `sessionId`, paragraph count, optional created path

2. `docx_open`
- Input: `path`
- Input options: optional `generateRedlines` default for the opened session
- Output: `sessionId`, paragraph count, brief preview

3. `docx_list_paragraphs`
- Input: `sessionId`, optional window (`start`, `limit`)
- Output: `{ id, index, text }[]`

4. `docx_edit_paragraph`
- Input: `sessionId`, `paragraphId`, `newText`, optional `author`, optional `generateRedlines`
- Behavior: apply reconciliation to a single paragraph and replace that node in `word/document.xml`
- Output: change status + updated paragraph preview

5. `docx_add_comment`
- Input: `sessionId`, `paragraphId`, `textToFind`, `comment`
- Behavior: inject comment markers + ensure `comments.xml` and relationships wiring
- Output: `commentsApplied`, warnings

6. `docx_save_as`
- Input: `sessionId`, `outputPath`
- Output: saved path, file size, modified flag

7. `docx_close`
- Input: `sessionId`
- Output: closed status

## Minimal valid package template (`docx_new`)

`docx_new` should produce this minimum valid set:

- `[Content_Types].xml`
- `_rels/.rels`
- `word/document.xml`

`word/document.xml` baseline must include `w:sectPr` as the last element in `w:body`:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p/>
    <w:sectPr/>
  </w:body>
</w:document>
```

`[Content_Types].xml` must declare at least:

- default for `.rels`
- default for `.xml`
- override for `/word/document.xml`

`_rels/.rels` must point to `word/document.xml` with officeDocument relationship type.

When comments are added later, package service appends:

- `word/comments.xml`
- `word/_rels/document.xml.rels` comments relationship
- comments override in `[Content_Types].xml`

## Track changes policy (redlines on/off)

- Session default:
  - Set by `docx_new`/`docx_open` via optional `generateRedlines`
  - Defaults to `true`
- Per-edit override:
  - `docx_edit_paragraph.generateRedlines` overrides the session default for that call
- Engine behavior:
  - `true` -> pass `generateRedlines: true` into reconciliation functions to emit `w:ins`/`w:del`/related revision nodes
  - `false` -> pass `generateRedlines: false`; changes are applied without creating new revision markup
- Non-goal for V1:
  - No automatic accept/reject of pre-existing tracked changes already in the source document

## Critical implementation detail

`applyRedlineToOxml` can return package/fragment OOXML in list/table-cell paths.  
For MCP document editing, add a normalization layer:

- Detect package output (`<pkg:package ...>`)
- Extract relevant `/word/document.xml` payload
- Convert fragment output into paragraph/table node replacements in the active document XML
- Never write fragment/package XML directly as `word/document.xml`

## Phase plan

### Phase 1: Server scaffold + runtime wiring

- Add `mcp/docx-server` package
- Add MCP SDK, `jszip`, `@xmldom/xmldom` dependencies
- Initialize server with `docx_new`, `docx_open`, and `docx_close`
- Acceptance: can create/open/close local `.docx` sessions from MCP client

### Phase 2: Document session + paragraph indexing

- Implement session store and paragraph indexing service
- Expose `docx_list_paragraphs`
- Implement session-level redline defaults and per-call override resolution
- Acceptance: client can inspect paragraph ids/text deterministically and observe stable redline mode selection

### Phase 3: Paragraph edit tool

- Implement `docx_edit_paragraph` using reconciliation service
- Replace targeted paragraph node in DOM and mark session dirty
- Acceptance: edited `.docx` opens in Word with valid track changes

### Phase 4: Comment tool + package part updates

- Implement `docx_add_comment`
- Update `comments.xml`, `[Content_Types].xml`, and `word/_rels/document.xml.rels` as needed
- Acceptance: inserted comments render correctly in Word review pane

### Phase 5: Save/export + client integration

- Implement `docx_save_as`
- Add Claude Code MCP config docs and example commands
- Acceptance: end-to-end open -> edit -> save works from Claude Code

### Phase 6: Validation and hardening

- Add XML/package validation checks before save
- Add rollback-on-failure per tool call
- Add test corpus (tables, lists, comments, existing redlines)
- Add explicit regression tests for both `generateRedlines=true` and `generateRedlines=false`
- Acceptance: no corrupted `.docx` in regression suite

## Test strategy

- Unit tests:
  - paragraph id resolution
  - package/fragment normalization
  - session lifecycle
  - redline mode resolution (session default vs per-edit override)
- Integration tests:
  - `docx_new` output opens in Word and validates as package
  - open/edit/save roundtrip against sample docs
  - list/table/comment scenarios
  - same edit in both redline modes (with/without revision markup)
- Smoke tests:
  - invoke tools through MCP transport and verify outputs

## Risks and mitigations

- Ambiguous paragraph targeting in repeated text
  - Mitigation: primary key = `w14:paraId`, fallback with index + text hash

- Fragment/package output from engine modes
  - Mitigation: strict normalization pipeline before DOM replacement

- OOXML relationship drift when adding comments
  - Mitigation: centralize rel/content-type writes in package service with validation

## Deliverables

- `mcp/docx-server/` runnable MCP server
- MCP tool schema + docs for Claude Code setup
- Automated tests for document integrity and edit correctness
- Example scripts for local manual verification
