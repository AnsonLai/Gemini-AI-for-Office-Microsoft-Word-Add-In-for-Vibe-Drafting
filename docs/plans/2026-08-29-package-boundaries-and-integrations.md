# Package Boundaries and Integrations Plan

**Date:** 2026-08-29  
**Status:** Active — consolidated plan; no work has been executed from this document.

## Objective

Finish the remaining delivery work around the extracted package, add-in
verification, npm publication, and the local MCP document workflow.

## Remaining work

1. **Word add-in rollout verification.** Run and record the Desktop/Web manual
   matrix for the shared operation adapter, including empty-target behavior,
   tracked edits, comments, highlights, undo, and known platform differences.
2. **Package publication.** Resolve npm authentication/2FA or use the documented
   release workflow, publish `@ansonlai/docx-redline-js`, then switch consumers
   from the temporary/local dependency arrangement to the published version and
   verify CDN/browser loading.
3. **Local MCP V1.** Implement the server scaffold, package/session services,
   paragraph indexing, paragraph edit, comment/package-part updates, save/export,
   rollback-on-failure, and regression coverage for redlines on and off.
4. **Repository boundary follow-through.** Keep add-in-only integration code out
   of the core package and decide/execute the later split of word-addin,
   browser-demo, and MCP repositories after the package publication path is
   stable.

## Explicitly not carried forward as open work

- The reconciliation-core extraction is complete; its duplicate `copy` plan is
  retained only as a migrated historical record.
- OOXML ingestion/export implementation is complete; its design record is
  retained alongside the completed implementation plan.

## Acceptance criteria

- Manual add-in results are recorded and reproducible.
- The package can be published and consumed through the intended release path.
- MCP open/edit/comment/save flows preserve valid `.docx` packages and pass the
  regression corpus.
- Core and host-specific boundaries remain explicit.

## Migrated source plans

- `docs/plans/migrated/2026-02-17-word-addin-standalone-engine-consolidation.md`
- `docs/plans/migrated/2026-02-22-reconciliation-repo-extraction-and-publish.md`
- `docs/plans/migrated/reconciliation-local-mcp-plan.md`
- `docs/plans/migrated/2026-02-17-word-ooxml-ingestion-export-design.md`
- `docs/plans/migrated/2026-02-21-reconciliation-core-extraction copy.md`

