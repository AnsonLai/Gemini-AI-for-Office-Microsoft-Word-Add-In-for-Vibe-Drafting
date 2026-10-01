# Canonical List Migration Follow-up

**Date:** 2026-09-30  
**Status:** Planned — waiting for upstream canonical conversion/remapping capabilities.  
**Baseline:** Exact `@ansonlai/docx-redline-js@0.8.3` in root and MCP consumers.

## Origin and scope

At the user's explicit request, the completed consumer reliability and fidelity
work in the [August agentic plan](completed/2026-08-29-agentic-tools-and-list-reliability.md)
is closed and archived. Its unfinished full canonical list migration is
transferred here. This transfer does not claim that all commands already use
canonical operations.

The 0.8.3 upgrade resolves the reported rejection, historical-inspection,
all-empty-range and public-facade numbering defects. [Validation](../validation-reports/2026-09-30-docx-redline-v083.md)
passes 51 offline suites, 68 actual Office.js checks, 92 independent Word checks
on their exports, and 12 additional facade Word checks. These are the starting
evidence, not validation of the future capabilities below.

## Current consumer behavior

The [library offload review](../library-offload-review.md) records the completed
execution and packaging migration across the five original plans. The remaining
native/legacy list routes below are the next offload opportunity, subject to the
upstream prerequisites and independent host checks in this plan.

- `insert_list_item` uses canonical body batches for the verified active
  bullet/decimal subset at source/resolved levels 0–1 with redlining enabled;
  root outdent needs a following same-list root sibling.
- Plain anchors, deeper levels, alternate numbering styles, tracking-off
  requests and unsupported outdent contexts retain verified native routes.
- General `edit_list` uses the established range reconciliation helper;
  header conversion uses native Word. Neither is fully migrated.
- Request validation, source-baseline guards and no replay after uncertain
  host writes remain required. Do not remove active helper modules.

## Upstream prerequisites

Library plan: `C:/Users/Phara/Desktop/Projects/Docx Redline JS/docs/plans/2026-09-30-canonical-list-operations.md`.
The published 0.8.3 release explicitly excludes canonical plain-to-list
conversion and list-format changes. [Capability report](../library-issues/2026-09-30-canonical-list-operations.md).

Required behavior includes:

1. Convert plain/manual headers to native lists while preserving paragraph
   identity, untouched formatting and original accepted/rejected views.
2. Change list kind or numbering style with safe numbering allocation,
   remapping, continuation/restart and range-boundary semantics.
3. Support changed header text together with conversion in an atomic batch.
4. Return explicit refusal for unsupported shapes; never report literal
   markers or unapplied formatting as successful native list conversion.
5. Resolve or explicitly constrain the bare-XML unchanged-header
   `RECEIPT_RECONCILIATION_FAILED` case before routing that form.

Keep library implementation separate from this repository. Negotiate the actual
released operation/schema/capabilities; the proposed `list-format` contract is
not an existing API to code against yet.

## Work packages

### WP0 — Validate the upstream release and define supported shapes

- Read final declarations, schema, receipts, capability changes and migration
  notes; update exact root/MCP pins together.
- Add source-bound regressions for conversion, format change and changed-header
  batches. Check unsupported shapes refuse before writes.
- Preserve single-empty/all-empty rejection and facade numbering regressions.
- Record supported levels, formats, scope and tracking behavior before routing.

### WP1 — Migrate consumer command paths

- Move supported `edit_list` and `convert_headers_to_list` requests to one
  immutable-source canonical batch and one confirmed Word insertion.
- Expand insertion routing only for shapes whose text, paragraph marks,
  numbering and tracking behavior have corresponding host evidence.
- Bind unsorted header indexes to their replacement texts before sorting.
- Select unsupported native paths before preparation; never replay after
  library failure or an uncertain host write.
- Preserve structured outcomes and source-baseline safeguards.

### WP2 — Independent Word and cross-consumer verification

- Extend frozen Word-authored sources for decimal/letter/Roman headers, deeper
  lists, continuation/restart, range splits and style-linked unsupported cases.
- Assert exact accepted/rejected paragraph arrays, list labels/values/levels,
  logical numbering identity and untouched formatting/package parts.
- Exercise actual production commands through Office.js, save DOCX exports,
  then independently open/Accept All/Reject All/reopen in desktop Word.
- Exercise complete DOCX facade operations through browser/MCP parity tests;
  preview and mocked Word proxies are not native fidelity oracles.
- Keep paid model evaluations and optional PDF rendering separate.

### WP3 — Cleanup and documentation

- Remove legacy helpers only after no active callers remain and coverage
  establishes equivalent behavior.
- Update tool contracts, architecture, README, state and roadmap with actual
  supported behavior, remaining native paths and final evidence.
- Once all recorded library issues are resolved, retain needed validation
  evidence in `docs/validation-reports/`, delete the entire `docs/library-issues/`
  folder, and remove or update every link to it. This is completion cleanup;
  keep the folder while issues remain open.
- Archive this follow-up only when its acceptance criteria are satisfied.

## Acceptance

- [ ] Released canonical operations support the required conversion and
      numbering-remapping behaviors, with explicit unsupported errors.
- [ ] Remaining supported list commands use source-bound canonical batches.
- [ ] Current/native fallback boundaries and tracking semantics are documented.
- [ ] Offline, actual Office.js and independent Word accepted/rejected gates pass.
- [ ] Browser/MCP facade fidelity and atomic batch behavior pass.
- [ ] Retired helpers have no active callers; docs and durable reports are updated.
- [ ] All recorded library issues are resolved, `docs/library-issues/` is deleted,
      and its incoming links are removed or updated.

**Next action:** implement and release the upstream capabilities, then start
WP0 here. No consumer-side OOXML workaround is authorized by this plan.
