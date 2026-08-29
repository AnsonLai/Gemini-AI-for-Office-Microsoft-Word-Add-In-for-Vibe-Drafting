# Agentic-Tools Alignment Plan (List/Bullet Reliability)

## Objective

Mirror `agentic-tools.js` list behavior more closely in reconciliation/browser-demo so:

1. Non-contiguous section header conversion (`1.`, `2.`, `3.`...) does not merge into unrelated existing lists.
2. Sub-sub insertions like `2.2.1` are applied at the correct nested level (not promoted to `2.3`).
3. Existing bullets/ordered lists outside edited targets are not renumbered or damaged.

## Implementation Plan

1. Establish clean baseline and lock current behavior.
   1. Reproduce both failures in `browser-demo` with deterministic prompts.
   2. Capture before/after `word/document.xml` + `word/numbering.xml` snapshots per turn.
   3. Add temporary list debug trace so routing decisions can be compared to `agentic-tools.js`.

2. Mirror `agentic-tools` list strategy in shared reconciliation helper.
   1. Create a shared helper (not demo-only) for list-context apply, modeled after `executeInsertListItem` and list-preserve paths.
   2. Inputs: target paragraph OOXML, optional adjacent paragraph OOXML, marker (`2.2.1`, etc.), operation type.
   3. Outputs: resolved `numId`, `ilvl`, and apply strategy (`reuse_existing_list`, `attach_to_dedicated_sequence`, `plain_redline`).

3. Replace browser-demo list routing with context-first routing.
   1. First try context reuse: if target/adjacent paragraph is list-bound, reuse that `numId` and derive `ilvl` from marker depth.
   2. For non-list section headers (`1.`, `2.`, `3.` across document), allocate one dedicated sequence per turn and reuse it.
   3. Use single-line structural fallback only as last resort when no list context exists.

4. Remove renumber-risk behavior from fallback path.
   1. Keep num-level start override only.
   2. Do not mutate abstract-level starts in browser-demo path.
   3. Do not emit independent numbering payloads per header once a dedicated sequence exists.

5. Strengthen marker-depth handling in shared code.
   1. Use marker depth as primary signal (`2.2.1` => `ilvl=2`).
   2. Keep indentation-based inference only as fallback.
   3. Apply the same logic in insertion-only and list-block synthesis paths.

6. Add regression tests before finalizing.
   1. Converting non-contiguous headers to `1..9` must not renumber existing body lists.
   2. Inserting `2.2.1` after `2.2` must retain same `numId` and use `ilvl=2`.
   3. Unrelated bullets/lists must preserve original numbering and order.

7. Update docs after tests pass.
   1. `src/taskpane/modules/reconciliation/ARCHITECTURE.md`
   2. `browser-demo/README.md`
   3. Add explicit list routing order and behavioral guarantees.

## Acceptance Criteria

1. Header conversion produces one dedicated chain (`1..9`) without affecting unrelated lists.
2. `2.2.1` is inserted as a sub-sub item, not as `2.3`.
3. No collateral renumbering or structural damage in untouched bullet/list regions.
