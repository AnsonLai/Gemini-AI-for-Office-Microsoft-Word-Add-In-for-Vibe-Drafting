# Reliability Hardening Plan (Model-Variance Reduction)

> **Migrated on 2026-08-29:** Remaining verification and quality-gate work was
> consolidated into [`2026-08-29-reliability-and-quality-gates.md`](../2026-08-29-reliability-and-quality-gates.md).
> This document is retained in `migrated/` as historical detail.

## Objective

Reduce brittleness and per-model performance variance in the Word add-in by moving
reliability enforcement out of prompts and into code: verified targeting, mechanical
validation, per-model configuration, richer failure feedback, history invariants, and
a model-in-the-loop eval harness.

## Status (last updated 2026-06-12)

| WP | Title | Status |
|----|-------|--------|
| WP1 | Verified content anchors | ✅ Done |
| WP2 | Mechanical change-set validator | ✅ Done |
| WP3 | Per-model capability/quirk registry | ✅ Done |
| WP4 | Informative tool-failure feedback | ✅ Done |
| WP5 | History invariant enforcement | ✅ Done (manual in-Word smoke test still recommended) |
| WP6 | Auto-checkpoint + IndexedDB | ✅ Done (manual in-Word save/revert smoke test still recommended) |
| WP7 | Model-in-the-loop eval harness | ✅ Done (real-model run needs a GEMINI_API_KEY; only the live inference step is unverified here) |
| WP8 | Documentation refresh | ✅ Done |

**All work packages complete.** Remaining before merge: the manual in-Word smoke tests called out under WP5/WP6 and the "Overall acceptance" item 3, plus an optional real-model eval run (WP7). Nothing has been committed yet.

### Post-plan follow-up fixes (2026-06-13, from live testing)
Surfaced while the user tested in real Word; not part of the original 8 WPs:
- **Engine skip reasons now reach the model.** `applyRedlineChangesToWordContext`
  (bridge) collects a `skipped[]` array (`{paragraphIndex, operation, reason}`) and
  returns it; `applyRedlineChangeSet` → `executeRedline` merge it into the
  `TOOL_FAILURE` detail via `formatRejections`. Fixes loops where the engine silently
  skipped a change (e.g. empty target) and the model retried blindly.
- **Empty-paragraph insertion now works — via NATIVE Word APIs.** The package's
  `toScopedSharedRedlineOperation` hard-rejects an empty target. First attempt routed the
  empty case through the OOXML reconstruction engine; that generated valid-looking OOXML in
  Node but Word's `insertOoxml` rejected it with `InvalidArgument` on a real empty/last
  paragraph. **Replaced with native insertion** (`insertContentAsNativeParagraphs` in
  `word-redline-runner.js`): `paragraph.insertText(firstLine,'Replace')` then
  `insertParagraph(line,'After')` for the rest. This relies on the document's
  change-tracking mode (`setChangeTrackingForAi` sets `trackAll` when redlines are on), so
  insertions are tracked automatically and revert cleanly — the proven pattern from the
  legacy `word-route-change.js:155`. Plain text only (markdown formatting not applied in
  this path — a known limitation). Logic covered by `tests/empty_paragraph_insertion_tests.mjs`
  (mock-based; the real Word `insertText`/`insertParagraph` can't run in Node).
- **Prompt guidance added** (`redline-prompt.js`): you can only target existing
  paragraphs; to add new content, use ONE replace_paragraph on an existing (ideally
  blank) paragraph with `\n`-separated content — don't invent paragraph numbers past
  the last [P#]. (This is the model's actual failure mode: it spread new content across
  non-existent P6–P10.)
- **Append-at-end-of-document now supported.** Targeting `paragraphIndex = paragraphCount+1`
  with a content-bearing op appends new paragraph(s) after the last paragraph:
  - `sanitizeChangeSet` allows `count+1` for `replace_paragraph`/`replace_range`/`edit_paragraph`
    (rejects `count+2`+ and `modify_text` at `count+1`).
  - The bridge detects `startIndex === paragraphCount` and appends via native
    `insertContentAsNativeParagraphs` (insert each `\n`-separated line as a new paragraph
    after the last one). Same native-API rationale as empty-paragraph insertion above —
    avoids the `InvalidArgument` Word throws on OOXML injection.
  - Prompt (`redline-prompt.js`) tells the model: one change with all new paragraphs in
    `content` (`\n`-separated); target a blank paragraph if present, else `count+1` to append.
  - **Not verified in real Word:** the `'After'` OOXML insertion path needs Office.js; the
    generation step is covered by `tests/empty_paragraph_insertion_tests.mjs`, but confirm
    the actual append renders correctly in Word (new paragraphs appear after the last one,
    tracked, and reject-all fully reverts).
- **Gemini 3.5 Flash thought-leakage fix (observed live 2026-07-04):** the thinking model
  emitted a change with its chain-of-thought in `replacementText` and NO `content`. WP2's
  `empty_content` guard blocked it — and crucially prevented the reasoning text being
  written into the document via the engine's `content ?? newContent ?? replacementText`
  fallback — but the task failed and burned a loop-guard strike. Three-part fix:
  1. **Sanitizer step 3b** (`change-validation.js`): repairs wrong-field payloads
     (`content`<->`newContent` mix-ups), then DELETES every field that does not belong to
     the chosen operation so junk in unused fields can never reach the engine fallback.
     Also: `edit_paragraph` missing `newContent` now rejects as `empty_content`.
  2. **In-tool corrective retry** (`executeRedline`, agentic-tools.js): when a change set
     yields 0 applied (document untouched), it re-asks the diff generator ONCE via
     `buildCorrectiveRetryPrompt(basePrompt, previousChanges, rejectionDetail)` (exported
     from `redline-prompt.js`) — showing the model its own invalid JSON + machine reasons —
     before reporting TOOL_FAILURE to the chat model. Max 2 diff attempts; a legitimate
     `[]` on attempt 1 still returns "no changes to suggest" without retrying.
  3. **Prompt/schema hardening** (`redline-prompt.js`): `originalText`/`replacementText`
     schema descriptions now say "OMIT entirely for other operations; never put
     notes/reasoning here", and the prompt gained a rule against writing reasoning into
     any field value.
  Regression tests: `testThoughtLeakageIntoUnusedFieldRejected` (the exact live payload),
  `testWrongFieldContentRepair`, `testInapplicableFieldsStripped`, `testCorrectiveRetryPrompt`.
- **Gemini 3.5 Flash repetition-loop fix (observed live 2026-07-05):** the diff model
  emitted the SAME empty `replace_range` object dozens of times until it exhausted the
  full 48k `maxOutputTokens`, producing ~149KB of JSON truncated mid-string →
  `JSON.parse` failed → null → and because `callGeminiForDiffs` had NO timeout/abort,
  the UI hung until the chat-level total timeout. Four-part fix:
  1. `repairTruncatedJsonArray(text)` (change-validation.js): string/escape/depth-aware
     scanner that salvages the complete leading objects from a truncated JSON array.
     Used by `callGeminiForDiffs` in place of bare `JSON.parse`; the sanitizer's
     `duplicate_target` dedupe then collapses the repeated objects.
  2. `callGeminiForDiffs` now has an `AbortController` + 90s timeout
     (`DIFF_CALL_TIMEOUT_MS`) so a degenerate generation can't hang the UI.
  3. Model profiles gained `diffMaxOutputTokens: 16384` (chat loop keeps 48000);
     bounds the cost/latency of a repetition loop before salvage kicks in.
  4. Prompt-size guards: `formatRejections` caps at 8 lines (+"…and N more");
     `buildCorrectiveRetryPrompt` truncates the echoed previous JSON at 4000 chars.
  Regression tests: `testRepairTruncatedJsonArray` (incl. the exact truncation shape,
  escaped quotes/braces in strings, unsalvageable inputs), `testFormatRejectionsCap`,
  retry-prompt truncation case in `testCorrectiveRetryPrompt`, profile value test.
  **Bigger picture:** gemini-3.5-flash has now shown two distinct structured-output
  pathologies (thought leakage, repetition loop) that gemini-2.5-flash handles cleanly.
  Consider running the WP7 eval harness against both and, if 3.5 keeps failing, marking
  its profile `toolCallReliability: "low"` and/or steering users to 2.5 as default.

### Notes for the next implementer
- **Test files referenced in this plan/README do not all exist.** The actual test
  suite lives in `tests/*.mjs` and is run as `node tests/<name>.mjs`. The README's
  `standalone_docx_plumbing_tests.mjs`, `standalone_operation_runner_tests.mjs`,
  `word_operation_runner_adapter_tests.mjs`, and `migrated_tool_cutover_tests.mjs`
  are **not present**. Use the real ones, e.g. `word_redline_runner_table_normalization_tests.mjs`,
  `insert_list_item_level_tests.mjs`, `agentic_tools_table_intent_tests.mjs`.
- **Pre-existing broken test (not your fault):** `targeting_helpers_extraction_tests.mjs`
  fails with `ERR_PACKAGE_PATH_NOT_EXPORTED` because it imports
  `@ansonlai/docx-redline-js/index.js`, a subpath the package no longer exports.
- **Lint reality:** `agentic-tools.js` has ~243 pre-existing prettier/no-undef
  violations (single quotes, missing trailing commas, `console` not declared global)
  and does NOT pass `office-addin-lint check` today. New code added to it intentionally
  matches the file's existing single-quote idiom rather than the prettier config, to
  avoid a 3300-line reformat that would bury real changes. New standalone modules
  (e.g. `change-validation.js`) DO pass lint cleanly. Do not run `--fix` on the whole
  `agentic-tools.js` file.
- `.js` source files use ESM syntax with no `"type"` in package.json; Node auto-detects
  and reparses them as ES modules (a `MODULE_TYPELESS_PACKAGE_JSON` warning is normal).
- Plain Node is v24 in this environment.

## How to use this plan

- Work packages (WP1–WP8) are ordered by priority but are **independent** unless a
  dependency is called out. Complete one WP fully (including its tests) before starting
  the next. Commit after each WP.
- Run the existing regression suites after every WP to confirm no regressions:
  ```bash
  node tests/standalone_docx_plumbing_tests.mjs
  node tests/standalone_operation_runner_tests.mjs
  node tests/word_operation_runner_adapter_tests.mjs
  node tests/migrated_tool_cutover_tests.mjs
  npm run lint
  ```
- Known pre-existing failure (NOT caused by your work, do not try to fix it):
  `formatting_tests.mjs` "Middle Paragraph Formatting" test fails on bold/italic.
- Key files:
  - `src/taskpane/taskpane.js` — chat UI, agentic loop, Gemini API calls, tool schemas.
  - `src/taskpane/modules/commands/agentic-tools.js` — tool implementations
    (`executeRedline`, `applyRedlineChangeSet`, `callGeminiForDiffs`, etc.).
  - `src/taskpane/modules/chat/chat-history.js` — history repair helpers
    (`validateHistoryPairs`, `removeAllFunctionPairs`, `createFreshStartWithContext`).
- Line numbers below are approximate anchors from the current codebase; if they have
  drifted, search for the quoted symbol names instead.

## Background: how the pipeline works today

1. User message goes into `chatHistory` and the agentic loop in `taskpane.js`
   (payload built around line ~1662) calls Gemini with native function-calling tools
   (`apply_redlines`, `edit_list`, `edit_table`, `insert_comment`, etc.).
2. When the model calls `apply_redlines`, `executeRedline(instruction, fullDocumentText)`
   in `agentic-tools.js` (line ~103) makes a **second** Gemini call
   (`callGeminiForDiffs`, line ~240, already uses `responseMimeType: "application/json"`
   + `responseSchema`) with a ~90-line prompt asking for a JSON array of changes,
   each targeting paragraphs by integer index (`[P#]` anchors).
3. `applyRedlineChangeSet(aiChanges, instruction)` (line ~68) applies the changes.
4. Failure handling today: many tool failures return `{ message, showToUser: false }`
   with generic text; malformed function calls are recovered with regex parsers in
   `taskpane.js` (`tryParseArgs` ~line 1780, `parseMalformedEditListArgs` ~line 1804);
   history corruption is repaired after the fact via a 4-tier ladder
   (`taskpane.js` ~lines 1686–1727).

Brittleness sources, in order of impact: (a) unverified integer paragraph targeting,
(b) prose rules in prompts that models follow inconsistently, (c) model quirks
hardcoded inline across files, (d) blind retries on uninformative errors.

---

## WP1 — Verified content anchors for paragraph targeting

**Goal:** an edit can never silently land on the wrong paragraph.

> ✅ **COMPLETED 2026-06-12.** What was built:
> - `src/taskpane/modules/commands/change-validation.js` (new, lint-clean) exporting
>   `normalizeForAnchor`, `parseAnchoredParagraphs`, and `verifyAnchor`. Note: the
>   document text reaching `executeRedline` is the `[P#|meta] text` format from
>   `extractEnhancedDocumentContext` (taskpane.js ~line 113), so `parseAnchoredParagraphs`
>   strips the `[P#|meta]` header to recover each paragraph's text. `verifyAnchor`
>   normalizes both sides (trim, collapse whitespace, strip a stray leading `[P#]`),
>   matches via `startsWith`/`includes`, searches a ±2 window, and only auto-corrects
>   on a single unambiguous neighbor match.
> - `agentic-tools.js`: imported the two helpers; added required `anchorText` to the
>   prompt spec, the `responseSchema` properties, and the schema `required` array;
>   `applyRedlineChangeSet` now takes a 3rd arg `paragraphTexts`, verifies each change,
>   collects `rejectedChanges`, applies `correctedIndex` (shifting `endParagraphIndex`
>   by the same offset), and returns `rejectedChanges`; `executeRedline` builds
>   `paragraphTexts` and surfaces rejections via the new local `formatAnchorRejections`
>   helper (WP4 should fold this into a shared `formatRejections`).
> - `tests/change_validation_tests.mjs` (new) — 7 test fns covering parse, exact match,
>   off-by-one correction, no-match-reports-actual, normalization tolerance, compat
>   mode (missing anchor), and ambiguous-neighbor handling. Passes.
> - Verified non-broken existing suites still pass (`word_redline_runner_table_normalization_tests`,
>   `insert_list_item_level_tests`, `agentic_tools_table_intent_tests`).

### Steps

1. In `agentic-tools.js`, edit the diff-generation prompt inside `executeRedline`
   (the `fullPrompt` template, ~line 115). Add a new required field to the change
   object spec:
   - `"anchorText"`: REQUIRED for every operation. The first 30–60 characters of the
     CURRENT text of the paragraph at `paragraphIndex`, copied VERBATIM from the
     document content provided. For `replace_range`, anchor the START paragraph.
     For empty paragraphs use `""`.
2. Update the `responseSchema` passed by `callGeminiForDiffs` (the `jsonSchema`
   object near line ~279) to include `anchorText` as a required string property.
3. Create a new module `src/taskpane/modules/commands/change-validation.js` exporting:

   ```js
   /**
    * Verify that a change's anchorText matches the actual paragraph text.
    * @param {object} change - one change object (has paragraphIndex, anchorText)
    * @param {string[]} paragraphTexts - actual paragraph texts, index 0 = [P1]
    * @returns {{ ok: boolean, correctedIndex?: number, reason?: string,
    *             actualTextSnippet?: string }}
    */
   function verifyAnchor(change, paragraphTexts) { ... }
   ```

   Behavior:
   - Normalize both sides before comparing: trim, collapse runs of whitespace to a
     single space, strip a leading `[P#]` marker if the model included one.
   - If `anchorText` is missing/empty: return `{ ok: true }` (backwards compatible —
     log a console warning).
   - If the paragraph at `paragraphIndex - 1` starts with (or contains) the
     normalized anchor: `{ ok: true }`.
   - Otherwise search a window of ±2 paragraph indexes. If exactly one neighbor
     matches, return `{ ok: true, correctedIndex }` (off-by-one tolerance).
   - Otherwise return
     `{ ok: false, reason: "anchor_mismatch", actualTextSnippet: <first 60 chars of the paragraph at the claimed index> }`.
4. In `applyRedlineChangeSet` (`agentic-tools.js` ~line 68), before applying each
   change: call `verifyAnchor`. Apply `correctedIndex` when present (also correct
   `endParagraphIndex` by the same offset for `replace_range`). Skip changes that
   fail, and collect them into a `rejectedChanges` array.
5. Change the return value of `executeRedline` so rejected changes are reported to
   the calling model (see WP4 message format): include for each rejection the
   claimed index, the anchor the model gave, and `actualTextSnippet`.

### Tests

Create `tests/change_validation_tests.mjs` (plain Node, no Word required — follow the
style of `tests/standalone_operation_runner_tests.mjs`: assert + console output,
exit code 1 on failure). Cases:
1. Exact anchor match at claimed index → ok.
2. Anchor matches index+1 only → ok with `correctedIndex`.
3. Anchor matches nowhere in window → `ok: false` with `actualTextSnippet`.
4. Whitespace/`[P#]`-prefix differences still match after normalization.
5. Missing anchorText → ok (compat mode).

### Acceptance criteria

- All new tests pass; existing suites still pass.
- A change with a wrong index and non-matching anchor is rejected, not applied.
- A change with an off-by-one index but matching neighbor anchor is applied to the
  correct paragraph.

---

## WP2 — Mechanical change-set validator (enforce prompt rules in code)

**Goal:** every rule currently stated as prose in the `executeRedline` prompt is also
enforced or auto-repaired in code, so weaker models that ignore prose still produce
valid change sets.

**Depends on:** nothing (but lives in the same new module as WP1:
`change-validation.js`).

> ✅ **COMPLETED 2026-06-12.** What was built:
> - `change-validation.js`: `sanitizeChangeSet(rawChanges, paragraphCount)` returning
>   `{ changes, rejected }`. Rejection entries are `{ change, paragraphIndex, operation, reason }`
>   (a superset of the planned `{change, reason}`, so they share a shape with WP1's
>   anchor rejections and can be formatted uniformly). All 9 rules implemented in the
>   documented order: `invalid_operation`, `index_out_of_range` (upper bound skipped
>   when `paragraphCount` is 0/unknown), `[P#]` strip (repair), `empty_content`,
>   `original_text_too_long`, `modify_text_structural_content`, `malformed_table`,
>   `schema_text_in_content`, `duplicate_target`.
> - `change-validation.js` ALSO now contains `formatRejections(rejected)` + an internal
>   `REJECTION_HINTS` reason→guidance map (this is the helper WP4 planned to create — it
>   was the natural home next to the reason codes, so it was built here. WP4's remaining
>   job shrinks to wiring it into edit_list/edit_table).
> - `agentic-tools.js`: imported `sanitizeChangeSet` + `formatRejections`; removed the
>   temporary local `formatAnchorRejections`; `executeRedline` now calls
>   `sanitizeChangeSet(aiChanges, paragraphTexts.length)` first, passes the sanitized
>   changes to `applyRedlineChangeSet`, merges sanitizer + anchor rejections, and reports
>   them in the `TOOL_FAILURE` / partial-success messages.
> - `tests/change_validation_tests.mjs`: +12 test fns (clean passthrough, invalid_operation,
>   index_out_of_range incl. unknown-count, [P#] strip, empty_content, original_text_too_long,
>   modify_text_structural_content, malformed_table, schema_text_leak, dedupe, formatRejections).
>   All pass; module is lint-clean.
> - **Implementation note on the pseudo-table guard:** follows the plan literally
>   (replace_* content that is a single line containing `|` with no newline and no `|---`
>   separator is rejected). This can in theory false-positive on legit single-line plain
>   text containing a pipe character; acceptable given the model is instructed not to emit
>   single-line tables. Revisit if users report it.

### Steps

1. In `src/taskpane/modules/commands/change-validation.js`, add:

   ```js
   /**
    * Sanitize and validate a raw AI change array.
    * @param {Array} rawChanges
    * @param {number} paragraphCount - total paragraphs in document
    * @returns {{ changes: Array, rejected: Array<{change, reason}> }}
    */
   function sanitizeChangeSet(rawChanges, paragraphCount) { ... }
   ```

2. Implement these rules, in this order, per change:
   - **Drop non-objects** and entries without a recognized `operation`
     (`edit_paragraph`, `replace_paragraph`, `modify_text`, `replace_range`).
     Reason: `invalid_operation`.
   - **Clamp/reject indexes:** `paragraphIndex` must be an integer in
     `[1, paragraphCount]`; `endParagraphIndex` (replace_range) must be ≥
     `paragraphIndex` and ≤ `paragraphCount`. Reject otherwise. Reason:
     `index_out_of_range`.
   - **Strip `[P#]` markers** from `content`, `newContent`, `replacementText`
     (regex: `/\[P\d+\]\s*/g`). This is a repair, not a rejection.
   - **Empty-content guard:** `replace_paragraph` / `replace_range` with missing or
     empty `content` → reject, reason `empty_content`.
   - **modify_text length:** if `originalText` is longer than 80 chars, do NOT
     reject — convert the change to `edit_paragraph` is not possible without the
     full paragraph text, so instead truncate is wrong too; therefore reject with
     reason `original_text_too_long` so the model can retry with `edit_paragraph`.
   - **modify_text content shape:** reject `modify_text` whose `replacementText`
     contains `\n`, a markdown table pipe row, or list markers (`/^\s*([-*]|\d+\.)\s/m`).
     Reason: `modify_text_structural_content`.
   - **Pseudo-table guard:** if `operation` is `replace_*` and `content` matches a
     single-line pipe string (contains `|`, no `\n`, no `|---`), reject with reason
     `malformed_table` (the prompt's "A|B|C" rule).
   - **Schema-text leak guard:** reject if `content`/`newContent`/`replacementText`
     contains substrings like `"paragraphIndex"`, `"operation":`, or
     `endParagraphIndex` (the model leaked schema into document text). Reason:
     `schema_text_in_content`.
   - **Dedupe:** if two changes target the same `paragraphIndex` with the same
     `operation`, keep the first, reject the rest with reason `duplicate_target`.
3. Call `sanitizeChangeSet` in `executeRedline` immediately after the
   `Array.isArray(aiChanges)` check (~line 201), BEFORE `applyRedlineChangeSet`.
   Merge its `rejected` list with WP1 anchor rejections in the tool result message.
4. Do NOT shorten the prompt in this WP (prompt slimming is a follow-up once the
   validator is proven in production).

### Tests

Extend `tests/change_validation_tests.mjs` with one test per rule above (≥9 cases),
plus one "clean change set passes through unmodified" case.

### Acceptance criteria

- Each rule has a passing test demonstrating rejection/repair.
- Valid change sets are not altered except for `[P#]` stripping.

---

## WP3 — Per-model capability/quirk registry

**Goal:** one place declares how each model behaves; no inline model-name
conditionals scattered through the code.

> ✅ **COMPLETED 2026-06-12.** What was built:
> - `src/taskpane/modules/config/model-profiles.js` (new, lint-clean, pure/Node-importable):
>   `getModelProfile(modelName)` + exported `MODEL_PROFILES` / `DEFAULT_PROFILE`. Profiles
>   cover the **actual** model lineup from the settings dropdown (`taskpane.html`):
>   `gemini-2.5-pro`, `gemini-2.5-flash`, `gemini-flash-latest`, `gemini-flash-lite-latest`,
>   `gemini-3.5-flash`, `gemini-3.1-pro-preview`. Prefix match uses the **longest** matching
>   key so `gemini-3.5-flash-preview` resolves to `gemini-3.5-flash`.
> - Wired into `taskpane.js`: `getModelProfile` imported; `modelProfile` resolved once at
>   the top of `sendChatMessage` (before the `try`, so it's in scope in the outer `catch`);
>   agentic-loop payload `maxOutputTokens` now from `modelProfile.maxOutputTokens`;
>   `callGeminiWithRetry(apiUrl, payload, modelProfile.retries)`; both throttle messages
>   (timeout-in-loop and the catch-block timeout override) now gated on
>   `modelProfile.previewThrottleWarning`.
> - Wired into `agentic-tools.js`: `callGeminiForDiffs` uses `modelProfile.temperature`
>   and `modelProfile.maxOutputTokens`.
> - `tests/model_profiles_tests.mjs` (new): exact lookup, prefix/longest-prefix match,
>   unknown→default, behavior-preserving defaults (48000/0.1/3), preview flag. All pass.
>
> **Deviations from the written spec (intentional, to satisfy the "no behavior change"
> acceptance criterion):**
> 1. **Values:** the illustrative spec used `maxOutputTokens: 65536` / `temperature: 0.2`.
>    Actual current behavior is `48000` (old `API_LIMITS.MAX_OUTPUT_TOKENS`) and `0.1`
>    (the diff call's temperature). Profiles use the real current values so nothing changes.
> 2. **Chat-loop temperature NOT set.** The spec said to set `generationConfig.temperature`
>    in the agentic loop from the profile. The loop historically set NO temperature; adding
>    one (e.g. 0.1) would materially change chat behavior. So the profile's `temperature`
>    drives only the deterministic diff/structured call (`callGeminiForDiffs`); the chat
>    loop still omits temperature. The profile field is documented accordingly.
> 3. **Throttle message reworded** to "This model is in preview…" (model-gated) instead of
>    the old always-on "If you're using Gemini 3…" string. Only `gemini-3.5-flash` and
>    `gemini-3.1-pro-preview` carry `previewThrottleWarning: true`.
> 4. **`API_LIMITS` is now effectively unused** (its two consumers were converted to the
>    profile). Left in place (definition + DI injection into agentic-tools) to avoid
>    touching dependency-injection plumbing; a future cleanup WP can remove it.
> 5. No inline model-name branching needed folding in: the "Fix tables for Gemini 3.5 Flash
>    quirks" commit applied table normalization to all models (not model-gated), so the only
>    model-dependent behavior was the throttle warning.

### Steps

1. Create `src/taskpane/modules/config/model-profiles.js`:

   ```js
   const MODEL_PROFILES = {
     "gemini-2.5-pro":     { maxOutputTokens: 65536, toolCallReliability: "high",
                             temperature: 0.2, retries: 3, supportsResponseSchema: true },
     "gemini-2.5-flash":   { maxOutputTokens: 65536, toolCallReliability: "high",
                             temperature: 0.2, retries: 3, supportsResponseSchema: true },
     "gemini-flash-latest":{ maxOutputTokens: 65536, toolCallReliability: "high",
                             temperature: 0.2, retries: 3, supportsResponseSchema: true },
     // Gemini 3.x preview models: throttling observed; table quirks (see git
     // commit "Fix tables for Gemini 3.5 Flash quirks")
   };
   const DEFAULT_PROFILE = { maxOutputTokens: 65536, toolCallReliability: "medium",
                             temperature: 0.2, retries: 3, supportsResponseSchema: true };
   function getModelProfile(modelName) {
     if (MODEL_PROFILES[modelName]) return MODEL_PROFILES[modelName];
     // prefix match so versioned names like gemini-2.5-flash-002 resolve
     const key = Object.keys(MODEL_PROFILES).find(k => modelName.startsWith(k));
     return key ? MODEL_PROFILES[key] : DEFAULT_PROFILE;
   }
   ```

   Before finalizing values, grep the codebase for existing per-model logic and fold
   it in: search for `gemini-` and `Gemini 3` in `src/taskpane/` (e.g. the throttle
   warning at `taskpane.js` ~line 1646, table quirk handling from the
   "Fix tables for Gemini 3.5 Flash quirks" commit — run
   `git show afe9525 --stat` and `git show 449439f` to find those sites).
2. In `taskpane.js`:
   - Import `getModelProfile`.
   - In the agentic loop payload (~line 1662) set
     `generationConfig.maxOutputTokens` and `generationConfig.temperature` from the
     profile of the active model (the model name is available where `apiUrl` is
     built, ~line 1301).
   - Pass `profile.retries` into `callGeminiWithRetry` (~line 2426) instead of the
     hardcoded `retries = 3` default.
   - Replace the hardcoded "Gemini 3 ... throttled" message logic (~line 1646) with
     a profile flag (e.g. `previewThrottleWarning: true` on 3.x profiles) so the
     message only shows for models that actually have the flag.
3. In `agentic-tools.js`, `callGeminiForDiffs` (~line 240): use the profile for
   `temperature` and `maxOutputTokens` instead of literals, keeping current values
   as the profile defaults so behavior is unchanged for existing models.
4. Do not change any model defaults in `loadModel` (`taskpane.js` ~line 576).

### Tests

Create `tests/model_profiles_tests.mjs`:
1. Exact name lookup returns its profile.
2. Prefix match (`gemini-2.5-flash-002` → `gemini-2.5-flash` profile).
3. Unknown model returns `DEFAULT_PROFILE`.

The module must be importable in plain Node (no Office.js, no DOM, no localStorage
at module top level).

### Acceptance criteria

- No behavior change for current default models (same tokens/temperature/retries).
- `grep -rn "Gemini 3" src/taskpane` shows the throttle message driven by the
  profile flag, not an inline model-name check.

---

## WP4 — Informative tool-failure feedback to the model

**Goal:** when a tool fails, the model's next attempt is informed, not blind.

**Depends on:** WP1/WP2 (uses their rejection data) — but can be done standalone for
the generic cases.

> ✅ **COMPLETED 2026-06-12.** What was built (note: the `formatRejections` helper and
> the redline zero-applied path were already built during WP2; this WP finished the rest):
> - `executeRedline` (agentic-tools.js): the "no valid array" return is now
>   `TOOL_FAILURE invalid_response: ...`; the partial-success path reports
>   `Applied X of Y edits ... Z rejected:\n<detail>` (Y = applied + rejected). The clean
>   success path (no rejections) message is **unchanged** per the acceptance criterion.
>   The `aiChanges.length === 0` ("no changes to suggest") path is intentionally left as a
>   non-failure — it means the model decided no edit was needed, not a failure to apply.
> - `executeEditList`: the no-items and catch returns now emit
>   `TOOL_FAILURE edit_list ... P{start}-P{end}: ...` with the index range.
> - `executeEditTable`: the catch return now emits
>   `TOOL_FAILURE edit_table at P{idx} ({action}): ...` (inner errors already carried the
>   failing stage). Per the plan, these functions were NOT restructured — only the
>   existing failure return messages were enriched with the index/action already in scope.
> - Loop guard (taskpane.js ~line 2337, `MAX_NO_PROGRESS_TOOL_LOOPS: 2`): **verified, not
>   changed.** A failed redline counts as attempted-but-not-successful; its detailed
>   `TOOL_FAILURE` text is included in `toolResult`, which is pushed into the
>   functionResponse content (taskpane.js ~line 2317), so the next retry is informed.
> - `tests/change_validation_tests.mjs`: `testFormatRejections` strengthened to assert each
>   line carries index + operation + reason (+ anchor/snippet when present). All pass.
>
> **Note for WP5+:** the three redline taskpane dispatch branches still push
> function-call/response pairs via raw `chatHistory.push` (taskpane.js ~lines 2327/2332);
> WP5 converts these to the atomic `appendFunctionExchange` helper.

### Steps

1. In `agentic-tools.js`, `executeRedline`: replace the three generic failure
   returns (~lines 201–223) with structured messages. Keep `showToUser: false`.
   Format (plain text, since it goes into a functionResponse):
   - No valid array: `"TOOL_FAILURE invalid_response: The diff generator did not return a JSON array. Retry with a simpler instruction or use edit_paragraph operations only."`
   - Zero applied: include per-change reasons, e.g.
     ```
     TOOL_FAILURE no_changes_applied: 3 changes were rejected.
     - P12 (edit_paragraph): anchor_mismatch. You claimed the paragraph starts with
       "The Receiving Party shall" but it actually starts with "ARTICLE 4 — TERM".
     - P15 (replace_range): empty_content.
     Re-read the paragraph numbers in the document content and retry with corrected
     indexes and anchorText copied verbatim.
     ```
   - Partial success: report both counts:
     `"Applied 4 of 6 edits. 2 rejected: <same per-change detail>."` with
     `showToUser: true`.
2. Build the rejection detail string in one helper in `change-validation.js`:
   `formatRejections(rejected)` → string. Reuse it for redline, and where
   straightforward, for `edit_list` / `edit_table` failure paths in
   `agentic-tools.js` (search for `return` statements with `showToUser: false`
   inside `executeEditList` / `executeEditTable` and add the failing index +
   reason where the information is already available — do not restructure those
   functions).
3. In `taskpane.js`, verify the loop guard (~line 2338, "Stopped to prevent a retry
   loop") still works: it should now trip less often because retries are informed.
   Do not change its threshold.

### Tests

Extend `tests/change_validation_tests.mjs`: `formatRejections` produces one line per
rejection containing the index, operation, reason, and (when present) the actual
text snippet.

### Acceptance criteria

- A rejected change produces a functionResponse string containing the actual
  paragraph snippet so the model can self-correct.
- User-visible chat messages are unchanged for the success path.

---

## WP5 — History invariant enforcement at write time

**Goal:** the 4-tier history repair ladder almost never fires because history can no
longer become invalid.

> ✅ **COMPLETED 2026-06-12.** What was built:
> - `chat-history.js`: added + exported `appendFunctionExchange(history, modelTurn, userTurn)`.
>   Validates shapes, requires ≥1 functionCall in the model turn, and enforces per-tool-name
>   call/response count parity (mirrors `validateHistoryPairs`); allows non-functionCall parts
>   (text/thought) in the model turn. Throws and leaves `history` untouched on any mismatch.
> - `taskpane.js`: imported the helper; the one genuine function-exchange push pair (the
>   model `parts` + user `functionResponses` turns, formerly two `chatHistory.push` calls)
>   now goes through `appendFunctionExchange`. Added a `console.warn` at the top of the
>   tier-1 recovery branch noting that reaching it means an invariant escaped the helper.
> - **Which pushes were intentionally left as raw `chatHistory.push`** (they are NOT function
>   exchanges, so the spec's "leave plain pushes" applies): the user-text prompt push
>   (~line 1328); the model-text + user-text recovery pair shown on
>   `MALFORMED_FUNCTION_CALL`/`UNEXPECTED_TOOL_CALL` (~lines 1973/1977); and the single
>   model-text turn that ends the loop on a normal response (~line 2376). Verified that
>   `functionResponses` is now consumed only inside `appendFunctionExchange`.
> - `tests/chat_history_invariant_tests.mjs` (new): 8 test fns — valid pair appends both,
>   mixed model parts allowed, multiple tools matched, name mismatch throws + leaves history
>   unchanged, count mismatch throws, no-functionCall throws, bad shapes throw, and
>   `validateHistoryPairs` leaves an appendFunctionExchange-built history unchanged. All pass.
> - **Lint note:** `chat-history.js` (like `agentic-tools.js`) predates the prettier config
>   and already had 27 cosmetic violations; the new block adds 2 of the same class
>   (indent/line-wrap). Left matching the file's existing idiom rather than reformatting the
>   whole file. Logic is correct and tests pass.
> - **Not run here:** the in-Word manual smoke test (one chat request that runs a tool).
>   taskpane.js syntax-checks and the unit tests pass, but a real Word run is still advised.

### Steps

1. In `src/taskpane/modules/chat/chat-history.js`, add and export:

   ```js
   /**
    * Append a model functionCall turn and its functionResponse turn atomically.
    * @param {Array} history
    * @param {object} modelTurn  - { role:"model", parts:[{functionCall:...}, ...] }
    * @param {object} userTurn   - { role:"user", parts:[{functionResponse:...}, ...] }
    * @returns {Array} the same history array (mutated)
    * @throws {Error} if the pair is malformed (mismatched names/counts)
    */
   function appendFunctionExchange(history, modelTurn, userTurn) { ... }
   ```

   Validation inside: every `functionCall` part in `modelTurn` must have a
   corresponding `functionResponse` part in `userTurn` with the same `name`, and
   counts must match. On mismatch, throw — the caller's existing catch paths handle
   it (and the repair ladder remains as the net).
2. In `taskpane.js`, find all `chatHistory.push` sites that push function-call or
   function-response turns (~lines 1957, 1961, 2316, 2321, 2362). Where a call turn
   and response turn are pushed in sequence, replace the two pushes with one
   `appendFunctionExchange` call. Leave the plain user-text push (~line 1322)
   untouched.
3. Keep `validateHistoryPairs` / `removeAllFunctionPairs` /
   `createFreshStartWithContext` and the tier ladder exactly as they are.
4. Add a `console.warn` inside the tier-1 recovery branch (`taskpane.js` ~line
   1690) noting that reaching this point means an invariant escaped
   `appendFunctionExchange` — useful signal during manual testing.

### Tests

Create `tests/chat_history_invariant_tests.mjs`:
1. Valid call/response pair appends both turns.
2. Mismatched function name throws and leaves history unchanged (length check).
3. Call turn with 2 functionCalls + response turn with 1 functionResponse throws.
4. `validateHistoryPairs` on a history built solely via `appendFunctionExchange`
   returns it unchanged.

Note: `chat-history.js` must remain importable in plain Node for these tests. If it
currently touches DOM/Office at import time, move those references inside functions.

### Acceptance criteria

- All pushes of function exchanges in `taskpane.js` go through
  `appendFunctionExchange`.
- Tests pass; manual smoke test (one chat request that runs a tool) still works.

---

## WP6 — Auto-checkpoint before mutating tools + IndexedDB storage

**Goal:** every document mutation is automatically recoverable; checkpoints stop
hitting the localStorage ~5MB quota.

> ✅ **COMPLETED 2026-06-12.** What was built:
> - `src/taskpane/modules/storage/checkpoint-store.js` (new, lint-clean). Async IndexedDB
>   API: `saveCheckpoint(label, ooxml)→id`, `getCheckpoint(id)`, `getLastCheckpoint()`,
>   `popLastCheckpoint()`, `listCheckpoints()`, `clearCheckpoints()`, plus
>   `importCheckpoints(records)` for migration. DB `aiwordplugin-checkpoints`, store
>   `checkpoints`, keyPath `id` autoIncrement, fields `{id, timestamp, label, ooxml}`,
>   cap 10 (oldest evicted). `indexedDB` is referenced only inside function bodies, so the
>   module is importable in plain Node. **Pure helpers** (exported, unit-tested):
>   `formatAutoCheckpointLabel`, `makeCheckpointRecord`, `idsToEvict`,
>   `migrateLegacyCheckpoints`, `MAX_CHECKPOINTS`.
> - `taskpane.js`: removed the localStorage `getCheckpoints`/`saveCheckpoints` and the
>   `STORAGE_LIMITS` constant; rewrote `createCheckpoint(silent, toolName)` and
>   `restoreCheckpoint(id)` to use the store; added one-time `ensureCheckpointMigration()`
>   that imports legacy `docCheckpoints` from localStorage then deletes the key.
> - **Auto-checkpoint:** all 8 mutating-tool dispatch sites already called
>   `createCheckpoint(true)`; changed to `createCheckpoint(true, functionCall.name)` so each
>   is labeled `auto:<tool>:<ISO>`. Added a 15s throttle (`CHECKPOINT_THROTTLE_MS`); a
>   throttled call returns the last auto id so per-message revert buttons still point at a
>   valid recent pre-edit state. Checkpoint failures return -1 and never block the edit.
> - `tests/checkpoint_store_tests.mjs` (new): label format, record shape/defaults, cap
>   enforcement (under/over/default), and the legacy migration transform. All pass.
> - README "Managing Checkpoints" + "Checkpoints not saving" updated to IndexedDB.
>
> **Design decisions / deviations:**
> 1. **Stable ids instead of array indices.** The old code returned a 0-based array index
>    that shifted on prune (a latent bug); `createCheckpoint` now returns the IndexedDB
>    autoincrement `id` (always ≥1, stable). This flows unchanged through
>    `updateSystemMessage`/`addUndoButton`/`onRestoreCheckpoint` (chat-ui treats it
>    opaquely and only special-cases `-1`), so **no chat-ui.js changes were needed** and
>    per-message revert is now more correct.
> 2. **Added `getCheckpoint(id)`** beyond the spec's export list — the existing UI restores
>    a *specific* per-message checkpoint, not just "last", so this accessor is required.
> 3. **No manual "Save Checkpoint" button exists** in the current UI (it was removed; the
>    `!silent` branches in `createCheckpoint` are retained but currently unused). All
>    checkpoints are auto-checkpoints taken before tools. The README was updated to reflect
>    the actual automatic-before-each-edit behavior rather than the old Save/Revert/Clear
>    button trio.
> 4. **Not run here:** IndexedDB needs a browser, so the DB read/write paths are exercised
>    only via the pure helpers in Node; a real in-Word save→edit→revert→clear smoke test is
>    still advised.

### Steps

1. Create `src/taskpane/modules/storage/checkpoint-store.js` wrapping IndexedDB:
   - DB name `aiwordplugin-checkpoints`, object store `checkpoints`, keyPath `id`
     (auto-increment), fields: `{ id, timestamp, label, ooxml }`.
   - Exports: `saveCheckpoint(label, ooxml)`, `getLastCheckpoint()`,
     `popLastCheckpoint()`, `clearCheckpoints()`, `listCheckpoints()` — all
     async/Promise-based.
   - Cap: keep at most 10 checkpoints; on save beyond 10, delete the oldest.
2. Find the existing checkpoint implementation in `taskpane.js` (search
   `localStorage` near `checkpoint`, see `getCheckpoints` usage around line ~1004)
   and switch it to the new store. Provide a one-time migration: on first use, if
   localStorage checkpoints exist, import them into IndexedDB then remove the
   localStorage key.
3. Auto-checkpoint: in the tool-dispatch section of the agentic loop in
   `taskpane.js` (the `if (functionCall.name === "apply_redlines")` block ~line
   2045 and the sibling branches for `edit_list`, `edit_table`, `edit_section`,
   `insert_list_item`, `convert_headers_to_list`), save a checkpoint labeled
   `auto:<toolName>:<ISO timestamp>` BEFORE executing the tool. Reuse whatever
   existing function captures full document OOXML for manual checkpoints.
   - Throttle: skip the auto-save if the last auto-checkpoint is <15 seconds old
     (multi-tool turns shouldn't spam snapshots).
   - If the checkpoint save fails, log a warning and proceed with the tool (do not
     block edits on checkpoint failure).
4. Update the README "Managing Checkpoints" and "Checkpoints not saving"
   troubleshooting sections to say IndexedDB instead of localStorage.

### Tests

IndexedDB isn't available in plain Node ≤ v22 without flags; structure
`checkpoint-store.js` so the pure logic (cap enforcement, label format, migration
transform) is in exported pure functions, and test those in
`tests/checkpoint_store_tests.mjs`. Manual test in Word: save → revert → clear via
the existing UI buttons.

### Acceptance criteria

- Manual checkpoint UI behaves as before (save/revert/clear).
- An `apply_redlines` run creates an `auto:` checkpoint visible in `listCheckpoints()`.
- A document larger than 5MB OOXML can be checkpointed (this fails today on
  localStorage).

---

## WP7 — Model-in-the-loop eval harness

**Goal:** a repeatable, scoreable way to measure tool reliability per model so prompt
or validator changes can be compared across models before shipping.

**Depends on:** WP1 + WP2 merged (the harness should exercise the validator path).

> ✅ **COMPLETED 2026-06-12.** What was built:
> - **Prompt extraction (the must-not-change-behavior part):**
>   `src/taskpane/modules/commands/redline-prompt.js` (new, lint-clean) exports
>   `buildRedlineDiffPrompt(instruction, fullDocumentText)` and `REDLINE_DIFF_SCHEMA`.
>   The template was extracted **programmatically** (sliced from the source, not retyped)
>   and verified **byte-identical** to the previous inline template when evaluated
>   (proven against the working tree before removal). `agentic-tools.js` now imports both;
>   its inline `fullPrompt` template and `jsonSchema` were removed.
> - `tests/redline_prompt_tests.mjs` (new): asserts key sentinel lines (incl. the literal
>   `\n` table example, the `anchorText` line, `diff-match-patch`), interpolation, and the
>   schema shape/required array. Passes; runs in the default suite (no API).
> - **Eval harness:** `tests/evals/run-evals.mjs` + 5 cases in `tests/evals/cases/`
>   (rename-title, parties→table, bullet-list, targeted-modify, no-change). The runner
>   reads `GEMINI_API_KEY`, accepts repeatable `--model`/`--case`, builds `[P#] text`
>   anchored content from the fixture `.docx` via `adm-zip` + the package's `parseOoxml` /
>   `getDocumentParagraphNodes` / `getParagraphText` (reusing `tests/setup-xml-provider.mjs`),
>   calls the model with the shared prompt+schema, and prints a per-model pass/fail table.
> - README "Development" gained an "AI Eval Harness (Manual)" section; the harness is
>   **manual-only** (gated on `GEMINI_API_KEY`, in no npm script).
>
> **Design decisions / deviations:**
> 1. **Scoring is done against the validated change set, not a re-applied document.** The
>    plan said "apply the resulting change set with the standalone operation runner" — but
>    the package exposes **no host-free function that applies the index-based change set**
>    (`edit_paragraph`/`replace_range`/etc.) to a docx; that logic is Word-API-bound in the
>    add-in. So `scoreChangeSet` runs the model output through the SAME gate production uses
>    (`sanitizeChangeSet` + `verifyAnchor`) to get `changesApplied`/`rejected`, and checks
>    `mustContainText`/`mustNotContainText` against the applied changes' proposed content.
>    This is deterministic, structural, exercises the WP1/WP2 validators, and directly
>    measures the per-model variance the harness exists to catch. `scoreChangeSet` and
>    `buildAnchoredText` are exported for unit-testability.
> 2. **The referenced `tests/standalone_operation_runner_tests.mjs` does not exist** (one of
>    the README's phantom test files); I used the real `setup-xml-provider.mjs` + package
>    parse functions instead.
> 3. **Windows libuv quirk:** `process.exit()` raced undici socket teardown and crashed
>    after printing results; switched to `process.exitCode` + natural drain.
> 4. **Eval/case files use single quotes**, matching the existing `tests/*.mjs` convention
>    (all of which predate the prettier config); `redline-prompt.js` in `src/` is lint-clean.
> 5. **Not verified here:** an actual run against a real model — it needs a paid
>    `GEMINI_API_KEY`. Everything up to and including the API call + error handling + clean
>    exit was smoke-tested (with the no-key guard and a dummy-key 400); only live inference
>    is unconfirmed.

### Steps

1. Create `tests/evals/` with:
   - `tests/evals/run-evals.mjs` — the runner (plain Node).
   - `tests/evals/cases/` — one JSON file per case:
     ```json
     {
       "name": "grammar-fix-single-paragraph",
       "documentFixture": "../../Sample NDA.docx",
       "instruction": "Fix the grammar in the paragraph about confidential information",
       "expect": {
         "minChangesApplied": 1,
         "maxChangesApplied": 3,
         "mustContainText": ["its obligations"],
         "mustNotContainText": ["[P"]
       }
     }
     ```
   - Start with 5 cases: single-paragraph grammar fix; multi-paragraph
     replace_range → markdown table; bullet-list insertion; targeted modify_text;
     and an instruction that should produce NO changes (expect 0 applied, 0
     rejected-as-applied).
2. Runner behavior:
   - Read `GEMINI_API_KEY` from env; accept `--model <name>` (repeatable) and
     `--case <name>` filters.
   - For each (model, case): load the fixture docx, extract paragraph text the same
     way the add-in builds `fullDocumentText` with `[P#]` anchors (reuse/extract the
     shared logic — if it currently lives only in Word-API code, replicate the
     anchor format `[P1] text...` from the docx XML via the existing standalone
     package `@ansonlai/docx-redline-js`).
   - Call the same prompt builder + `callGeminiForDiffs`-equivalent path. If those
     functions are not importable outside the add-in bundle, extract the prompt
     template and schema into a shared module
     `src/taskpane/modules/commands/redline-prompt.js` first (pure string/JSON, no
     Office.js) and have both the add-in and the runner import it. **This
     extraction must not change the prompt text.**
   - Apply the resulting change set with the standalone operation runner (see
     `tests/standalone_operation_runner_tests.mjs` for how to apply ops to document
     XML without Word) and score against `expect`.
   - Output a per-model table: case | pass/fail | changesApplied | rejected |
     failure reason. Exit code 0 only if all cases pass for all requested models.
3. Add to README under Development:
   ```bash
   GEMINI_API_KEY=... node tests/evals/run-evals.mjs --model gemini-2.5-pro --model gemini-flash-latest
   ```
4. This suite calls a paid API — it must NOT run as part of the default regression
   commands. Keep it manual-only.

### Acceptance criteria

- Runner completes against at least one real model with all 5 cases scored.
- Deterministic scoring: re-running compares only structural expectations, not
  exact model wording.
- The add-in builds and behaves identically after the prompt-template extraction
  (diff the generated prompt string in a unit test:
  `tests/redline_prompt_tests.mjs` asserts the template contains key sentinel
  lines like `CRITICAL: Return ONLY valid JSON`).

---

## WP8 — Documentation refresh

**Goal:** README/ARCHITECTURE match the actual system so contributors (and AI agents)
don't act on stale information.

> ✅ **COMPLETED 2026-06-13.** What was changed:
> - `README.md`: replaced the `gemini-1.5-*` model section with the real `loadModel`
>   defaults (`gemini-flash-latest` fast / `gemini-2.5-pro` slow), noted Settings-based
>   selection, and pointed to `model-profiles.js`; rewrote the Project Structure tree to the
>   actual `src/taskpane/modules/*` layout (commands, chat, config, storage, context, glance,
>   ui, utils, docx-redline-js-integration) plus `browser-demo/`, `mcp/docx-server/`,
>   `tests/`, `plans/`. Checkpoint docs (IndexedDB) and the eval-harness command were already
>   added in WP6/WP7.
> - `README.md`: **replaced the three phantom "regression" subsections** (which documented
>   `standalone_docx_plumbing_tests.mjs`, `standalone_operation_runner_tests.mjs`,
>   `word_operation_runner_adapter_tests.mjs`, `migrated_tool_cutover_tests.mjs` — none of
>   which exist) with an accurate "Regression Tests" section listing only suites that
>   actually pass, and a note about the pre-existing `ERR_PACKAGE_PATH_NOT_EXPORTED` failures.
> - `ARCHITECTURE.md`: added the **"Two-stage LLM pipeline"** and **"Reliability layers"**
>   sections (model profiles → schema → sanitizer → anchor verification → auto-checkpoints →
>   retry feedback → history invariants → loop guard).
> - Verified `STATE.md` and `ROADMAP.md` **do exist** at the repo root, so the ARCHITECTURE
>   "Operational Guidance" reference to them is valid (no correction needed).
> - **Acceptance verified:** audited every path/command in both docs against the repo (no
>   stale `gemini-1.5`; all 11 documented `tests/*.mjs` commands resolve and were run to
>   confirm they pass; the `reconciliation/` mention is an explicit "was removed" note).

### Steps

1. `README.md`:
   - Replace the `gemini-1.5-flash` / `gemini-1.5-pro` model section (~lines
     277–288) with the actual defaults from `loadModel` in `taskpane.js`
     (`gemini-2.5-pro` slow / `gemini-flash-latest` fast) and note models are
     user-configurable in Settings.
   - Update the Project Structure tree to the real layout: include
     `src/taskpane/modules/` (subfolders: `commands/`, `chat/`,
     `docx-redline-js-integration/`, `utils/`, plus any added by WP3/WP6),
     `browser-demo/`, `mcp/docx-server/`, `plans/`, `tests/`.
   - Update checkpoint docs per WP6 (IndexedDB).
   - Add the eval harness command per WP7.
2. `ARCHITECTURE.md`:
   - Add a section "Two-stage LLM pipeline" documenting: chat model with function
     calling selects a tool → tool makes a second Gemini call with structured
     output (`responseSchema`) to generate a change set → change set is sanitized
     (`change-validation.js`), anchor-verified, then applied via
     `applyRedlineChangeSet` → failures are returned to the chat model as
     structured `TOOL_FAILURE` messages.
   - Add a section "Reliability layers" listing, in order: model profiles (WP3),
     structured output schemas, change-set sanitizer (WP2), anchor verification
     (WP1), auto-checkpoints (WP6), informative retry feedback (WP4), history
     invariants (WP5), loop guard.
   - Verify `STATE.md` and `ROADMAP.md` exist at repo root; if they don't, remove
     or correct that reference in Operational Guidance item 4.
3. Keep marketing copy in README unchanged; only correct technical content.

### Acceptance criteria

- Every file path and command in both docs resolves against the repo
  (spot-check with `Test-Path` / running each documented test command).

---

## Out of scope (do not do in this plan)

- Shortening or rewriting the `executeRedline` prompt rules (only ADD `anchorText`).
- Changing default models, tool names, or tool schemas beyond adding `anchorText`.
- Touching `@ansonlai/docx-redline-js` package internals.
- Fixing the pre-existing `formatting_tests.mjs` "Middle Paragraph Formatting" failure.
- Streaming, UI redesign, or new tools.

## Overall acceptance

1. All new test files pass: `change_validation_tests.mjs`, `model_profiles_tests.mjs`,
   `chat_history_invariant_tests.mjs`, `checkpoint_store_tests.mjs`,
   `redline_prompt_tests.mjs`.
2. All four pre-existing regression suites still pass.
3. Manual smoke test in Word: a chat request that triggers `apply_redlines`
   applies edits with track changes, creates an auto-checkpoint, and a deliberately
   wrong instruction (e.g. referencing nonexistent content) produces an informed
   retry rather than a wrong edit.
