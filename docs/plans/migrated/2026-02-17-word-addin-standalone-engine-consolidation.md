# Word Add-In Standalone Engine Consolidation Implementation Plan

> **Migrated on 2026-08-29:** Remaining rollout verification was consolidated
> into [`2026-08-29-package-boundaries-and-integrations.md`](../2026-08-29-package-boundaries-and-integrations.md).
> This document is retained in `migrated/` as historical detail.

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Consolidate Word add-in mutation paths so `apply_redlines`, `insert_comment`, and `highlight_text` use the same shared reconciliation operation engine already used by standalone/browser flows.

**Architecture:** Keep `taskpane.js` focused on tool orchestration and move operation-application behavior into shared reconciliation services. Add a Word-specific adapter that uses `standalone-operation-runner` core logic on scoped OOXML (`paragraph` or `range`) and writes results back with existing Word integration helpers. Use direct cutover to the shared operation path for migrated tools (no legacy runtime fallback branches).

**Tech Stack:** Office.js (Word), ES modules, reconciliation core (`standalone.js`, `services/standalone-operation-runner.js`), integration helpers (`integration/word-ooxml.js`), Node-based regression tests for pure modules.

---

## Status

- Completed: highlight/comment migration to shared standalone operation bridge in `src/taskpane/modules/commands/agentic-tools.js` (no legacy runtime fallback path).
- Completed: redline/operation routing cutover in `executeRedline` now routes canonical operations through reconciliation integration adapter `applyWordOperation(...)` (sequential, pure OOXML, strict out-of-range no-op, no implicit append).
- Completed: dedicated Word adapter module `src/taskpane/modules/reconciliation/integration/word-operation-runner.js` for shared standalone operation application on paragraph/range scopes.
- Added: command-layer redline converter seam in `src/taskpane/modules/commands/redline-operation-converter.js` for mapping add-in payloads into canonical shared redline operations.
- Added: shared OOXML format-only no-span handling now routes through reconstruction fallback in `engine/oxml-engine.js` so migrated paths do not stall on native API deferral when span extraction is empty.
- Added: add-in redline routing includes a targeted `modify_text` nearest-paragraph rebase guard when the indexed paragraph is empty, using `originalText` matching to avoid broad paragraph drift.
- Next focus: finish cleanup of remaining non-redline overlap surfaces and align any residual tool prompt language with canonical redline contract.

### Current Constraints

- Converter (`redline-operation-converter`) is intentionally strict: it requires non-empty scoped start text and returns no-op when scope text cannot be resolved.
- The add-in command layer currently applies empty-target mitigation only for `modify_text` (via nearby paragraph matching on `originalText`).
- For `edit_paragraph` / `replace_paragraph` / `replace_range`, empty-target paragraph indices still no-op instead of speculative remapping.

## Execution Progress (2026-02-19)

- `Task 1` converter seam: completed.
  - Added `src/taskpane/modules/commands/redline-operation-converter.js`.
  - Added `tests/redline_operation_converter_tests.mjs`.
  - Wired converter into `executeRedline` in `src/taskpane/modules/commands/agentic-tools.js`.
- `Task 3` highlight migration: completed.
  - `executeHighlight` now routes through shared scope bridge only (`applySharedOperationToWordScope`).
- `Task 4` comment migration: completed.
  - `executeComment` now routes through shared scope bridge only (`applySharedOperationToWordScope`).
- `Task 2` Word adapter module: completed.
  - Added `src/taskpane/modules/reconciliation/integration/word-operation-runner.js`.
  - Re-exported adapter surface in `src/taskpane/modules/reconciliation/index.js`.
  - Kept command compatibility shim in `src/taskpane/modules/commands/shared-operation-bridge.js`.
  - Added regression coverage in `tests/word_operation_runner_adapter_tests.mjs`.
- `Task 5` redline migration: completed.
  - `executeRedline` now delegates directly to `applyWordOperation(...)` using canonical operation conversion.
  - Preserved strict out-of-range no-op behavior and sequential application semantics.
  - Added adapter-level regression cases for paragraph rewrite, range list insertion, single-paragraph concatenation insertion-shape, and table replace-range.
- `Task 6` docs/refactor cleanup: completed.
  - Removed legacy command compatibility shim `src/taskpane/modules/commands/shared-operation-bridge.js`.
  - Removed dead helper in command layer (`validateToolPrerequisites`) and kept migrated tools on shared adapter-only paths.
  - Updated architecture/reconciliation/root docs to reflect direct adapter usage and no bridge shim.
- `Task 7` verification/rollout: in progress (automated verification complete; manual add-in matrix pending).
  - Verified repeatedly in-session:
    - `node tests/redline_operation_converter_tests.mjs`
    - `node tests/shared_operation_bridge_tests.mjs`
    - `node tests/standalone_operation_runner_tests.mjs`
  - Added adapter-focused coverage:
    - `node tests/word_operation_runner_adapter_tests.mjs` (includes redline/highlight/comment paragraph/range cases)
    - `node tests/no_legacy_shared_operation_bridge_tests.mjs` (enforces removal of legacy command bridge)
    - `node tests/migrated_tool_cutover_tests.mjs` (enforces command-level redline/comment/highlight shared-engine cutover and no legacy fallback imports)
  - Added manual runbook template:
    - `tests/word-desktop/README.md` -> `Shared-Engine Manual Matrix (Desktop + Web)`
  - Outstanding: run the Desktop/Web manual matrix locally and record results using the runbook template.
- `Execution update (Task 2 + Task 5 pass)`: completed.
  - Added reconciliation Word adapter: `src/taskpane/modules/reconciliation/integration/word-operation-runner.js`.
  - Added adapter regression suite: `tests/word_operation_runner_adapter_tests.mjs`.
  - Cut `executeRedline` over to direct adapter delegation (`applyWordOperation(...)`) in `src/taskpane/modules/commands/agentic-tools.js`.
  - Converted command bridge to compatibility re-export: `src/taskpane/modules/commands/shared-operation-bridge.js`.
  - Exported adapter API in `src/taskpane/modules/reconciliation/index.js`.
  - Verification run:
    - `node tests/word_operation_runner_adapter_tests.mjs`
    - `node tests/shared_operation_bridge_tests.mjs`
    - `node tests/redline_operation_converter_tests.mjs`
    - `node tests/standalone_operation_runner_tests.mjs`
- `Execution update (Task 6 + Task 7 automation pass)`: completed on 2026-02-22.
  - Removed legacy command compatibility shim `src/taskpane/modules/commands/shared-operation-bridge.js`.
  - Added cleanup guard `tests/no_legacy_shared_operation_bridge_tests.mjs`.
  - Expanded adapter regression coverage for highlight/comment operations in `tests/word_operation_runner_adapter_tests.mjs`.
  - Fixed Word adapter package selection for single-paragraph list insertion-shape outputs so multi-node runner output is preserved.
  - Verification rerun:
    - `node tests/standalone_operation_runner_tests.mjs`
    - `node tests/standalone_docx_plumbing_tests.mjs`
    - `node tests/standalone_smoke.mjs`
    - `node tests/include_numbering_behavior.mjs`
    - `node tests/no_word_api_standalone_check.mjs`
    - `node tests/redline_operation_converter_tests.mjs`
    - `node tests/shared_operation_bridge_tests.mjs`
    - `node tests/word_operation_runner_adapter_tests.mjs`
    - `node tests/no_legacy_shared_operation_bridge_tests.mjs`
    - `node tests/migrated_tool_cutover_tests.mjs`
- `Execution update (Task 7 continuation pass)`: completed on 2026-02-22.
  - Added command cutover guard test: `tests/migrated_tool_cutover_tests.mjs`.
  - Added Desktop/Web manual matrix runbook template to `tests/word-desktop/README.md`.
  - Updated root verification docs to include the new cutover regression command.
## Decision Log (2026-02-17)

1. Canonical redline model (Q1)
- Recommendation: keep current AI/tool payloads for now, and normalize them at the command boundary into one internal redline operation object before execution.
- Clarification: "canonical redline" means the single internal operation contract used by the shared runner (not a new AI prompt format immediately). The converter maps `edit_paragraph` / `replace_paragraph` / `replace_range` / `modify_text` inputs into this one execution shape.
- Naming transition note: keep shared redline field name `modified` for this consolidation pass; plan a follow-up migration to align add-in terminology/language with the shared contract after routing stabilization.

2. Text mutation policy
- Chosen: substring search/replace behavior (`modify_text` style) for redline text edits.

3. Execution substrate
- Chosen: pure OOXML mutation path through shared operation runner (no mixed Word-API text mutation path for migrated redline operations).

4. Multi-operation handling
- Chosen: sequential handling.

5. Append/out-of-range target handling
- Chosen: explicit append operation only.
- Policy: when a resolved target is out of range, do strict no-op with explicit warning/error.
- No implicit append behavior in redline operations.
- Append is permitted only through a dedicated append operation shape so behavior is intentional and auditable.

## Overlap Inventory (Current State)

### 1) Redline/Operation Routing Overlap
- Add-in path:
  - `src/taskpane/modules/commands/agentic-tools.js:61` (`executeRedline`)
  - Large local branching for `edit_paragraph`, `replace_paragraph`, `replace_range`, `modify_text`
  - Local list/table handling and OOXML insertion logic
- Shared standalone path:
  - `src/taskpane/modules/reconciliation/services/standalone-operation-runner.js:901` (`applyOperationToDocumentXml`)
  - Shared target resolution, list/table heuristics, insertion-only list logic, comments/highlights/redlines

### 2) Highlight/Comment Overlap
- Add-in path:
  - `executeHighlight` uses per-paragraph OOXML patching in command layer (`agentic-tools.js:1250`)
  - `executeComment` uses Word `search + insertComment` directly (`agentic-tools.js:1164`)
- Shared standalone path:
  - `applyOperationToDocumentXml` already supports `highlight` and `comment` with shared target logic.

### 3) Table/List Heuristic Overlap
- Add-in path:
  - Local table reconcile wrapper `tryApplyMarkdownTableWithOxmlEngine` (`agentic-tools.js:1651+`)
  - Local list tools (`executeInsertListItem`, `executeEditList`, `executeConvertHeadersToList`)
- Shared standalone path:
  - List/table targeting and structural fallback logic in shared reconciliation core + operation runner.

### 4) Existing Consolidation Already Present
- Add-in already uses shared integration route planner for some `edit_paragraph` flow:
  - `routeChangeOperation(...)` -> `routeWordParagraphChange(...)`
  - `src/taskpane/modules/reconciliation/integration/word-route-change.js`
- This is a strong foundation for deeper consolidation.

## Target Consolidation Design

1. Introduce one canonical operation shape for mutation execution:
- `redline`: `{ type, targetRef, targetEndRef?, target, modified }`
- `highlight`: `{ type, targetRef, target, textToHighlight, color }`
- `comment`: `{ type, targetRef, target, textToComment, commentContent }`

2. Add a Word adapter layer that:
- Reads scoped OOXML (paragraph or expanded range),
- Executes shared `applyOperationToDocumentXml(...)`,
- Writes back via `insertOoxmlWithRangeFallback(...)` and `withNativeTrackingDisabled(...)`,
- Applies numbering payload when returned.

3. Keep `taskpane.js` and tool prompts mostly unchanged initially; add a converter from existing AI tool payloads into canonical operations.

4. Roll out as direct cutover per tool:
- Migrated tools execute the shared operation path only.
- No legacy runtime fallback branch remains for migrated tools.

---

### Task 1: Add operation-conversion seam (no behavior change) [Completed]

**Status:** Implemented and wired.

**Files:**
- Create: `src/taskpane/modules/commands/redline-operation-converter.js`
- Modify: `src/taskpane/modules/commands/agentic-tools.js`
- Test: `tests/redline_operation_converter_tests.mjs`

**Step 1: Write failing tests for conversion**
- Cover conversion from current tool outputs (`edit_paragraph`, `replace_paragraph`, `replace_range`, comment/highlight payloads) to canonical operations.

**Step 2: Implement minimal converter**
- Implemented pure helpers:
  - `applySubstringSearchReplace(...)`
  - `toScopedSharedRedlineOperation(...)`

**Step 3: Wire converter into command layer (no execution change yet)**
- Generate canonical operation objects in parallel with legacy logic for logging/validation.

**Step 4: Run tests**
- `node tests/redline_operation_converter_tests.mjs`

**Step 5: Commit**
- `feat(commands): add canonical operation converter for add-in tools`

---

### Task 2: Build Word adapter for shared standalone operation runner [Completed]

**Status:** Completed.

**Files:**
- Create: `src/taskpane/modules/reconciliation/integration/word-operation-runner.js`
- Modify: `src/taskpane/modules/reconciliation/index.js`
- Test: `tests/word_operation_runner_adapter_tests.mjs`

**Step 1: Write failing adapter tests (pure/mocked)**
- Implemented in `tests/word_operation_runner_adapter_tests.mjs` with mocked Word context/scope:
  - scope selection (`single paragraph` vs `expanded range`)
  - call-through to shared runner (injectable `runner`)
  - insertion mode and error behavior (no legacy runtime fallback branch)
  - result handling (`hasChanges` false path)

**Step 2: Implement adapter**
- Public API:
  - `applyWordOperation(context, operation, scope, options)`
- Use:
  - `applyOperationToDocumentXml` from `services/standalone-operation-runner.js`
  - `insertOoxmlWithRangeFallback`, `withNativeTrackingDisabled` from integration helpers
 - Added shared OOXML bridge exports in adapter module:
   - `applySharedOperationToParagraphOoxml(...)`
   - `applySharedOperationToScopeOoxml(...)`

**Step 3: Export adapter**
- Re-export via `src/taskpane/modules/reconciliation/index.js`.

**Step 4: Run tests**
- `node tests/word_operation_runner_adapter_tests.mjs`

**Step 5: Commit**
- `feat(reconciliation): add Word adapter for standalone operation runner`

---

### Task 3: Migrate `executeHighlight` to shared engine path [Completed]

**Status:** Implemented via `applySharedOperationToWordScope` in command layer.

**Files:**
- Modify: `src/taskpane/modules/commands/agentic-tools.js`
- Test: `tests/redline_operation_converter_tests.mjs` (and add highlight-specific wiring tests as needed)

**Step 1: Write failing test for highlight operation wiring**
- Verify canonical highlight ops are passed to Word adapter.

**Step 2: Implement migration**
- Replace local `applyHighlightToOoxml` mutation block with `applyWordOperation(...)`.
- Remove legacy local highlight mutation fallback branch.

**Step 3: Run tests**
- `node tests/redline_operation_converter_tests.mjs`

**Step 4: Commit**
- `refactor(commands): route highlight tool through shared standalone operation engine`

---

### Task 4: Migrate `executeComment` to shared engine path [Completed]

**Status:** Implemented via `applySharedOperationToWordScope` in command layer.

**Files:**
- Modify: `src/taskpane/modules/commands/agentic-tools.js`
- Test: `tests/redline_operation_converter_tests.mjs` (and add comment-specific wiring tests as needed)

**Step 1: Write failing test for comment operation wiring**
- Verify canonical comment ops are produced and delegated.

**Step 2: Implement migration**
- Route comment ops via `applyWordOperation(...)`.
- Remove legacy `search + insertComment` fallback branch.

**Step 3: Run tests**
- `node tests/redline_operation_converter_tests.mjs`

**Step 4: Commit**
- `refactor(commands): route comment tool through shared standalone operation engine`

---

### Task 5: Migrate `executeRedline` primary path to shared engine adapter [Completed]

**Status:** Completed for primary execution path.

**Files:**
- Modify: `src/taskpane/modules/commands/agentic-tools.js`
- Modify: `src/taskpane/modules/reconciliation/integration/word-route-change.js` (only if needed for shared helper reuse)
- Test: `tests/word_operation_runner_adapter_tests.mjs`

**Step 1: Write failing regression tests**
- Added adapter-level regression coverage in `tests/word_operation_runner_adapter_tests.mjs`:
  - simple paragraph rewrite
  - explicit range list insertion between existing items
  - single-paragraph concatenation insertion shape
  - markdown table replace_range

**Step 2: Implement redline delegation via shared engine**
- In `executeRedline`, convert AI changes to canonical redline operations.
- Use substring search/replace as the default redline text mutation strategy.
- Apply via `applyWordOperation(...)` from reconciliation integration adapter.
- Redline path no longer calls command-local OOXML bridge helpers directly.

**Step 3: Preserve tracking semantics**
- Preserved existing `setChangeTrackingForAi` lifecycle and disabled native tracking during shared adapter insertions.

**Step 4: Run tests**
- `node tests/word_operation_runner_adapter_tests.mjs`

**Step 5: Commit**
- `refactor(commands): route redline execution through shared standalone operation engine`

---

### Task 6: Remove duplicated legacy branches after cutover [Completed]

**Status:** Completed (legacy command bridge removed; docs aligned with direct reconciliation adapter path).

**Files:**
- Modify: `src/taskpane/modules/commands/agentic-tools.js`
- Modify: `src/taskpane/modules/reconciliation/ARCHITECTURE.md`
- Modify: `src/taskpane/modules/reconciliation/README.md`
- Modify: `README.md`

**Step 1: Delete migrated legacy branches**
- Remove local migrated operation branches (no fallback helpers for migrated tools).

**Step 2: Remove direct command-layer engine calls where superseded**
- Reduce local `applyRedlineToOxml`/list/table decision logic in command layer.

**Step 3: Update architecture/docs**
- Document canonical operation pipeline and strict no-fallback policy for migrated tools.

**Step 4: Commit**
- `docs+refactor: document add-in shared-engine cutover and no-fallback policy`

---

### Task 7: Verification matrix and rollout [In Progress]

**Status:** Automated verification matrix complete in-session; manual Word Desktop/Web matrix remains pending.

**Files:**
- Modify: `tests/` (add scenario fixtures if needed)
- Modify: `src/taskpane/taskpane.js` (feature flag wiring if runtime-config driven)

**Step 1: Automated verification**
- `node tests/standalone_operation_runner_tests.mjs`
- `node tests/standalone_docx_plumbing_tests.mjs`
- `node tests/standalone_smoke.mjs`
- `node tests/include_numbering_behavior.mjs`
- `node tests/no_word_api_standalone_check.mjs`
- New:
  - `node tests/redline_operation_converter_tests.mjs`
  - `node tests/word_operation_runner_adapter_tests.mjs`
  - `node tests/migrated_tool_cutover_tests.mjs`

**Step 2: Manual Word add-in verification checklist (Desktop + Web)**
- Redline single paragraph
- Redline explicit range list insertion
- Redline table conversion
- Highlight op
- Comment op
- Track changes on/off permutations
- Use runbook template: `tests/word-desktop/README.md` (`Shared-Engine Manual Matrix (Desktop + Web)`)

**Step 3: Rollout**
- Enable migrated path directly for target tool set.
- Monitor logs and error rate.
- Remove old path immediately for migrated tools once verification passes.

**Step 4: Commit**
- `chore: enable shared standalone operation engine for Word add-in after validation`

---

## Notes / Non-Goals (Initial Pass)

- `edit_table`, `edit_section`, and `convert_headers_to_list` remain specialized tools in initial consolidation and are not fully reimplemented in standalone operation runner in this pass.
- Zip/package-level standalone docx plumbing (`standalone-docx-plumbing.js`) is not directly applicable to Word add-in runtime mutation; Word adapter integration focuses on OOXML scope mutation + insertion.
