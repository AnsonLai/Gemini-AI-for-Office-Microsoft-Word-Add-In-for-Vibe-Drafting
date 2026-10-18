# Canonical List Migration Follow-up

> Scope update (2026-10-02): the owner approved the [library-first lightweight
> harness plan](./2026-10-02-library-first-lightweight-agent-reliability.md).
> Unsupported native editing commands and their migration queue are retired.
> The smaller retained subset uses direct library batches; broader list/table
> structure is refused. This is an intentional capability retirement, not
> evidence that upstream defects were fixed. Preserve the dated incident and
> independent library reproducers below, including issue cleanup conditions.

**Date:** 2026-09-30  
**Status:** Complete and archived (2026-10-17). WP1–WP3 done for the exposed list actions; browser/MCP facade parity was dropped when those spike consumers were removed.
**Baseline:** Exact `@ansonlai/docx-redline-js@0.9.0` in root and MCP consumers (verified 2026-10-12).

## Origin and scope

At the user's explicit request, the completed consumer reliability and fidelity
work in the [August agentic plan](./2026-08-29-agentic-tools-and-list-reliability.md)
is closed and archived. Its unfinished full canonical list migration is
transferred here. This transfer does not claim that all commands already use
canonical operations.

The 0.8.3 upgrade resolved the reported rejection, historical-inspection,
all-empty-range and public-facade numbering defects. [Validation](../../validation-reports/2026-09-30-docx-redline-v083.md)
passes 51 offline suites, 68 actual Office.js checks, 92 independent Word checks
on their exports, and 12 additional facade Word checks. These are the starting
evidence, not validation of the future capabilities below.

**2026-10-02 note:** 0.8.4 adds `textOccurrence` and `applyRedlineToOxml`
`numberingXml` (see the [0.8.4 validation](../../validation-reports/2026-10-02-docx-redline-v084.md)),
but still lacks plain-to-list conversion and list-format changes. It also
introduced a numbering-continuity regression ([separate list operations restart
numbering](../../validation-reports/2026-10-17-library-issues-closeout.md))
that blocks canonical `convert_headers_to_list`. The acceptance criteria and the
no-consumer-side-workaround stance below are unchanged.

**2026-10-12 note:** 0.9.0 satisfies upstream prerequisites 1 to 5 below
([0.9.0 validation](../../validation-reports/2026-10-12-docx-redline-v090.md)):

1. `list-format` `convert-plain` converts plain headers in place with `w14:paraId` preserved.
2. `list-format` `change-format` and `join` allocate new numbering definitions.
3. `redline` plus `list-format` compose for changed-text headers.
4. Unsupported shapes return explicit refusal codes.
5. The bare-XML `RECEIPT_RECONCILIATION_FAILED` case is fixed.

It also adds `paragraph-format.numbering`, and the numbering-continuity
regression is fixed (separate list operations now share one list). The agent's
typed intents expose none of this yet: header-to-list conversion and
numbering-style changes are still refused before writing (KB-4 in
[`KNOWN_BUGS.md`](../KNOWN_BUGS.md)). The old `reproduce-plain-header-conversion.mjs` reproducer is
obsolete (now refused, `BATCH_OPERATION_FAILED`); the capability moved to
`list-format` `convert-plain`. All library issues are resolved
([index](../../validation-reports/2026-10-17-library-issues-closeout.md)); the no-consumer-side-workaround stance is
unchanged.

## Current consumer behavior

The [library offload review](../../library-offload-review.md) records the completed
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
The published 0.8.3 release explicitly excluded canonical plain-to-list
conversion and list-format changes. [Capability report](../../validation-reports/2026-10-17-library-issues-closeout.md).

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
- **2026-10-12 status:** the root and MCP pins are updated to exact 0.9.0, and the offline and host gates are green (46 of 46 suites; 14 default Office.js cases; independent Word oracle 68 PASS, 0 FAIL). No `list-format` intent is exposed yet, so the source-bound regressions for conversion, format change and changed-header batches, and the supported-shape record, remain to do.

### WP1 — Migrate consumer command paths

> **2026-10-14 status (offline done, host evidence pending):** the typed intents
> `convert_to_list`, `change_list_format` and `continue_list` map one-to-one onto the
> library's `list-format` (see [tool contracts](../../agentic-tool-contracts.md)).
> Header conversion, changed-text header plus conversion, bullet-to-decimal and
> list continuation are covered offline by `tests/document_list_intents_tests.mjs`,
> and golden step 12 is a replay instead of a scripted refusal. Library
> refusals pass through (`error.failedOperations`) with no write. Not done:
> real-Word golden, Office.js and Word-oracle evidence (KB-4), range scope,
> `paragraph-format.numbering` and the retirement of the legacy list helpers.
>
> **2026-10-14 host evidence (WP2):** five list-action capability cases plus one no-write
> refusal control run in the default Office.js manifest (20 cases; 26 host checks) from
> Word-authored sources (`scripts/create-agentic-list-action-fixtures.ps1`). The independent
> Word oracle asserts `ListFormat.ListString` after `Revisions.AcceptAll/RejectAll`: 88 PASS,
> 0 FAIL (the 14 original cases unchanged). Artifacts: `.cache/reliability/list-officejs`, `list-oracle`.

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
- ~~Exercise complete DOCX facade operations through browser/MCP parity tests;
  preview and mocked Word proxies are not native fidelity oracles.~~ Dropped with the browser demo and MCP server (2026-10-17).
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

- [x] Released canonical operations support the required conversion and
      numbering-remapping behaviors, with explicit unsupported errors (0.9.0).
- [x] Remaining supported list commands use source-bound canonical batches
      (`convert_to_list`, `change_list_format`, `continue_list`, `insert_list_item`; 2026-10-14).
- [x] Current/native fallback boundaries and tracking semantics are documented
      (no native fallback; [tool contract](../../agentic-tool-contracts.md)).
- [x] Offline, actual Office.js and independent Word accepted/rejected gates pass
      ([2026-10-14 list-format validation](../../validation-reports/2026-10-14-list-format-intents.md)).
- [x] ~~Browser/MCP facade fidelity and atomic batch behavior pass.~~ Dropped 2026-10-17: the owner removed the browser demo and MCP server as obsolete spikes. Parity of the Word Flat OPC path with the library's DOCX facade, including atomic refusal, is still checked by `tests/docx_cross_host_parity_tests.mjs`.
- [x] Retired helpers have no active callers; docs and durable reports are updated (2026-10-17: the retired list modules `list-operation-plan.js`, `list-level-utils.js`, `word-structured-list.js` and `redline-operation-converter.js` are absent, which `tests/no_legacy_shared_operation_bridge_tests.mjs` asserts, and no unused list helper remains in `src/`).
- [x] All recorded library issues are resolved, `docs/library-issues/` is deleted,
      and its incoming links are removed or updated.

**Closed (2026-10-17):** every acceptance item is met or dropped. Still unrun, and tracked in the phase-2 plan's release track rather than here: a live-provider check of golden step 12 and Word Online. `paragraph-format.numbering` and range-scoped changes stay unexposed until there is a product need.
