# Package Boundaries, Integrations, and Web Portability Plan

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** MCP facade migration complete; portable consumer boundary and browser-demo migration pending.

**Recommended order:** 3 of 4 August plans.

## Baseline and dependencies

Both consumers pin exact `@ansonlai/docx-redline-js@0.8.2`. The [September upgrade](completed/2026-09-09-docx-redline-js-v0.5.4-upgrade.md) is complete, including actual Office.js transport validation.

Outcome/recovery policies in [reliability](completed/2026-08-29-reliability-and-quality-gates.md) are complete. Preserve the verified mappings and explicit capability limits in [agentic tools](2026-08-29-agentic-tools-and-list-reliability.md) when consolidating interfaces. Its broader list migration remains open on separately recorded library dependencies; that does not prevent portable extraction or browser document-lifecycle migration. The browser demo already uses the shared Gemini request client.

## Current architecture and completed work

| Boundary | Actual current state |
| --- | --- |
| Library | Public OOXML and `openDocx` / inspect / apply / serialize lifecycle |
| MCP | Upgrade WP5 completed `DocxDocument` sessions/services, generic errors and atomic `docx_apply_operations` |
| MCP packaging | Direct JSZip/xmldom dependencies and obsolete package/targeting/XML services removed |
| Add-in integration index | Re-exports library **and Word-specific adapters**; not a host-neutral module |
| Add-in redline transport | `executePureOoxmlBatch(context, targetScope, ...)` performs Word I/O; pure preparation helpers also present |
| Browser demo | Still imports JSZip and several library service subpaths; not migrated to the document facade |

The previous plan incorrectly placed `executePureOoxmlBatch` in a zero-Office.js core and claimed complete web portability/removal of JSZip everywhere. Editing logic can be portable without promising reuse of every host-dependent tool or UI component.

## Remaining work packages

### WP1 — Separate portable consumer logic from Word adapters

- Inventory imports/callers before splitting entry points; avoid a blanket export that pulls Word adapters into non-Word consumers.
- Define a portable entry point for operation mapping, validation and engine result handling without Word globals, host context, filesystem access or browser UI dependencies.
- Keep `getOoxml`, `insertOoxml`, tracking-mode management and `executePureOoxmlBatch` in the Word transport entry point.
- Extract shared pure preparation where consumers need it; preserve supported exports during migration.
- Add meaningful boundary checks and an offline import/execution test for the portable entry point.

### WP2 — Migrate browser-demo document editing

- Inventory `browser-demo/demo.js` ZIP/XML packaging, library subpath imports and preview dependencies.
- Route DOCX open, inspect, mutation and save through the public document lifecycle where supported.
- Preserve editing modes, comments, styles, numbering, relationships and download behavior.
- Distinguish preview/editing dependencies: docx-preview currently expects global JSZip. Retain it until preview requirements change; removal is not a facade prerequisite.
- Share proven mapping/prompt helpers without importing Word transport.
- Verify file open/edit/download/reopen in a browser, including localized edits, comments and atomic failure.

### WP3 — Maintain MCP contracts and cross-host parity

MCP facade migration is already complete; do not recreate its session store/services.

- Preserve tool names, session semantics, target references and generic error propagation.
- Extend stdio tests for changed contracts; retain create/edit/batch/comment/rollback/save/reopen coverage.
- Run identical fixture operations through portable core/applicable host adapters; compare semantic accepted/rejected outcomes and preserved parts rather than byte-identical ZIP archives.
- Document responsibilities and unsupported host capabilities.

## Deferred transport option

Whole-document `getFileAsync(Compressed)` → `openDocx` → `insertFileFromBase64` is a proposed future Word transport. It is **not implemented or certified by the September upgrade**. Evaluate only for a demonstrated requirement with separate host fidelity tests. Range/body Flat-OPC insertion is the implemented upgrade path.

## Acceptance

- [x] MCP uses document sessions without direct JSZip/xmldom editing services.
- [ ] Portable consumer entry point executes without Word/UI globals.
- [ ] Word I/O stays in its adapter, with callers migrated safely.
- [ ] Browser demo uses supported document APIs for editing and passes save/reopen checks.
- [ ] Preview dependencies and cross-host capability differences documented.

**Next:** [OOXML engine and performance](2026-08-29-oxml-engine-and-performance.md), profiling settled interfaces before selecting optimization work.
