# Package Boundaries, Integrations, and Web Portability Plan

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** Complete (2026-09-30).

**Recommended order:** 3 of 4 August plans.

## Baseline and dependencies

Both consumers pin exact `@ansonlai/docx-redline-js@0.8.2`. The [September upgrade](2026-09-09-docx-redline-js-v0.5.4-upgrade.md) is complete, including actual Office.js transport validation.

Outcome/recovery policies in [reliability](2026-08-29-reliability-and-quality-gates.md) are complete. Preserve the verified mappings and explicit capability limits in [agentic tools](../2026-08-29-agentic-tools-and-list-reliability.md) when consolidating interfaces. Its broader list migration remains open on separately recorded library dependencies; that does not prevent portable extraction or browser document-lifecycle migration. The browser demo already uses the shared Gemini request client.

## Architecture before this extraction

| Boundary | State before this extraction |
| --- | --- |
| Library | Public OOXML and `openDocx` / inspect / apply / serialize lifecycle |
| MCP | Upgrade WP5 completed `DocxDocument` sessions/services, generic errors and atomic `docx_apply_operations` |
| MCP packaging | Direct JSZip/xmldom dependencies and obsolete package/targeting/XML services removed |
| Add-in integration index | Re-exports library **and Word-specific adapters**; not a host-neutral module |
| Add-in redline transport | `executePureOoxmlBatch(context, targetScope, ...)` performs Word I/O; pure preparation helpers also present |
| Browser demo | Still imports JSZip and several library service subpaths; not migrated to the document facade |

The previous plan incorrectly placed `executePureOoxmlBatch` in a zero-Office.js core and claimed complete web portability/removal of JSZip everywhere. Editing logic can be portable without promising reuse of every host-dependent tool or UI component.

## Work packages and closure record

### Execution record (2026-09-30)

- Completed upgrade/reliability plans moved to `completed/` in `3daf159`; links repaired.
- WP1 extraction separates preparation from Word I/O in `consumer-core.js` and preserves existing adapter exports.
- Browser baseline failed at startup on an unresolved XML dependency; the import map now targets the published browser bundle.
- WP2 routes browser open, inspect, editing, and serialization through the document-session facade. JSZip remains for docx-preview only. Browser UI opened, edited, downloaded, and reopened a DOCX: reupload exposed 42 prompt paragraphs, inspection found 44 paragraphs including two empty revision-view paragraphs, and all 24 source texts were retained exactly. The downloaded packages had matching SHA-256 hashes; unchanged XML parts matched exactly, with lists, table, and comment content retained.
- WP3 parity checks exercise shared fixtures through the portable core, Word batch adapter, MCP, and browser facade. Accepted and rejected text, comments, style/numbering/hyperlink relationships, and preserved foreign revisions matched semantically; atomic failures left the source unchanged. The tests compare package semantics and unchanged part bytes, not whole ZIP byte identity.
- Final verification: 47 offline suites passed with 0 failures (4 excluded); development and production builds passed; provider calls were false. The actual Office.js adapter run passed 8 checks and the independent Word oracle passed 12; browser document validation passed 7 checks. See the [validation record](../../validation-reports/2026-09-30-package-boundaries-and-integrations.md) and [consumer package boundaries](../../package-boundaries.md).

### WP1 — Separate portable consumer logic from Word adapters

**Status: Complete.** The explicit `consumer-core.js` entry point exports immutable source/baseline capture, redline and list mapping, argument validation, pure OOXML operation preparation and result handling. `prepareCanonicalBatch` returns ready/no-op/refused/package-error states without claiming a host write. Word adapters retain scope I/O, tracking and confirmed mutation outcomes. Existing named exports remain compatible. Boundary/import/execution tests and existing adapter suites pass.

- Inventory imports/callers before splitting entry points; avoid a blanket export that pulls Word adapters into non-Word consumers.
- Define a portable entry point for operation mapping, validation and engine result handling without Word globals, host context, filesystem access or browser UI dependencies.
- Keep `getOoxml`, `insertOoxml`, tracking-mode management and `executePureOoxmlBatch` in the Word transport entry point.
- Extract shared pure preparation where consumers need it; preserve supported exports during migration.
- Add meaningful boundary checks and an offline import/execution test for the portable entry point.

### WP2 — Migrate browser-demo document editing

**Status: Complete.** The browser demo uses the public document-session facade for DOCX lifecycle and atomic editing, while JSZip remains available to docx-preview. Real-browser UI validation confirmed open/edit/download/reopen behavior and package preservation.

- Inventory `browser-demo/demo.js` ZIP/XML packaging, library subpath imports and preview dependencies.
- Route DOCX open, inspect, mutation and save through the public document lifecycle where supported.
- Preserve editing modes, comments, styles, numbering, relationships and download behavior.
- Distinguish preview/editing dependencies: docx-preview currently expects global JSZip. Retain it until preview requirements change; removal is not a facade prerequisite.
- Share proven mapping/prompt helpers without importing Word transport.
- Verify file open/edit/download/reopen in a browser, including localized edits, comments and atomic failure.

### WP3 — Maintain MCP contracts and cross-host parity

**Status: Complete.** MCP contracts and session semantics remain covered, and shared fixture operations pass through the portable core, Word adapter, MCP facade, and browser session with matching semantic outcomes and atomic rollback.

MCP facade migration is already complete; do not recreate its session store/services.

- Preserve tool names, session semantics, target references and generic error propagation.
- Extend stdio tests for changed contracts; retain create/edit/batch/comment/rollback/save/reopen coverage.
- Run identical fixture operations through portable core/applicable host adapters; compare semantic accepted/rejected outcomes and preserved parts rather than byte-identical ZIP archives.
- Document responsibilities and unsupported host capabilities.

## Deferred transport option

Whole-document `getFileAsync(Compressed)` → `openDocx` → `insertFileFromBase64` is a proposed future Word transport. It is **not implemented or certified by the September upgrade**. Evaluate only for a demonstrated requirement with separate host fidelity tests. Range/body Flat-OPC insertion is the implemented upgrade path.

## Acceptance

- [x] MCP uses document sessions without direct JSZip/xmldom editing services.
- [x] Portable consumer entry point executes without Word/UI globals.
- [x] Word I/O stays in its adapter, with callers migrated safely.
- [x] Browser demo uses supported document APIs for editing and passes save/reopen checks.
- [x] Preview dependencies and cross-host capability differences documented.

**Next:** [OOXML engine and performance](2026-08-29-oxml-engine-and-performance.md), profiling settled interfaces before selecting optimization work.
