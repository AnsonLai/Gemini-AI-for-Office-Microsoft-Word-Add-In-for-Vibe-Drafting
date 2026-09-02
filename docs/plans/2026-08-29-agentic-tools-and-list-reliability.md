# Agentic Tools and List Reliability Plan

**Date:** 2026-08-29  
**Status:** Active — consolidated plan; no work has been executed from this document.

## Objective

Finish moving reusable operation and list behavior out of command-local code,
then make list conversion and nested insertion deterministic across the add-in
and browser demo.

## Already completed

- The route planner, shared list parsing, package construction, Word OOXML
  helpers, list-markdown builders, and paragraph-identity parsing have already
  been extracted for the scoped candidates.

## Remaining work

1. Establish deterministic browser-demo reproductions for non-contiguous header
   conversion, nested markers such as `2.2.1`, and untouched neighboring lists.
   Capture document and numbering XML before and after each case.
2. Add a shared context-first list-application helper that resolves `numId`,
   `ilvl`, and an explicit strategy (`reuse_existing_list`,
   `attach_to_dedicated_sequence`, or `plain_redline`).
3. Route browser-demo and add-in list operations through the same context-first
   strategy. Reuse one dedicated numbering sequence per conversion turn.
4. Make marker depth the primary nesting signal and indentation only a fallback.
   Ensure insertion-only and synthesized list-block paths use the same rule.
5. Remove renumbering risks from fallbacks: do not mutate abstract-level starts
   and do not emit independent numbering payloads once a sequence exists.
6. Finish the remaining command-layer orchestration/OOXML-builder cleanup,
   especially list and header-conversion routines still identified in the
   source plan.
7. Add regression coverage for dedicated numbering chains, nested insertion,
   and preservation of unrelated list regions; update architecture and demo
   documentation after the tests stabilize.

## Acceptance criteria

- Header conversion creates one intended chain without collateral renumbering.
- `2.2.1` remains at nesting level 2 and does not become `2.3`.
- Existing unrelated bullets and ordered lists are byte/structure stable except
  for intended edits.

## Migrated source plans

- `docs/plans/migrated/2026-02-09-agentic-tools-move-out-plan.md`
- `docs/plans/migrated/agentic-tools-alignment-list-fix-plan.md`

