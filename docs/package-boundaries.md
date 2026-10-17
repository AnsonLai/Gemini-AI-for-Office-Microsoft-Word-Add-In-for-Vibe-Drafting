# Consumer package boundaries

Both manifests exact-pin `@ansonlai/docx-redline-js@0.9.0`. The library owns canonical inspection, targeting, literal matching, OOXML mutation, revisions, numbering and DOCX package semantics. Application prompting, scope, harness and Word transport defects are fixed here; engine defects receive upstream reports and independent reproducers.

| Surface | Responsibility |
| --- | --- |
| Loop, registry and guide | Bounded provider/tool exchanges, scope, paired history, truthful completion and UI events. |
| Inspection session | Bounded opaque handles/cursors tied to strong library descriptors and fresh source identity. |
| Intent compiler | Validate the complete allowlisted proposal and map desired states to public canonical operations. No model call, search engine or native edit algorithm. |
| Document tools | Fresh resolution, exact-source checkpoint, atomic preparation/commit and affected-state verification. |
| `@ansonlai/docx-redline-js/word` (library) | Read Word Flat OPC parts, assemble the insertion package with comment/numbering parts, and Word-aware revision tokens (since 0.10.0; was `consumer-core.js` and `word-source-identity.js`). |
| Word adapter | Source/selection/count reads, sync, one body insertion, tracking state and checkpoint transport. |
| Browser document session | Open/inspect/atomic edit/serialize through the public complete-DOCX facade; JSZip remains for preview. |
| MCP document service | Node file/session I/O and public facade operations; persistence is explicit. |

The Word plugin keeps Flat OPC transport, as approved by the owner. Supported editing has no native paragraph/list/table algorithm or compatibility executor, and no package assembly: the library's `/word` surface builds the one package the plugin writes. Word preparation does not confirm a write; synchronized host evidence plus verification does. Browser/MCP edits happen in memory, with download/save separate. Offline success or preview does not certify Office.js transport.

Ordinary-comment output is a merge delta. Until a supported public merge path
exists, the Word consumer refuses adding comments to sources with existing
comments; it never replaces the whole part with a delta or appends duplicate
parts. Restore uses a derived comparison for proven Word-generated metadata,
guarded by canonical comment relationship/anchor equality, without changing XML
payloads or ordinary freshness checks.

Current behavior: [tool contracts](agentic-tool-contracts.md). Measured changes and host evidence: [reliability plan](./plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md). Unsupported engine shapes remain in [library issues](library-issues/README.md); removal of native fallback does not close them.
