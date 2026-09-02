# Web Performance Optimization Plan

> **Migrated on 2026-08-29:** Remaining performance work was consolidated into
> [`2026-08-29-oxml-engine-and-performance.md`](../2026-08-29-oxml-engine-and-performance.md).
> This document is retained in `migrated/` as historical detail.

Date: 2026-02-08 (updated)

## Problem Statement

The OOXML reconciliation engine performs well on Word Desktop but shows noticeable slowness on Word for the Web. Word Online's JavaScript runtime has tighter memory constraints, slower DOM operations, and higher `context.sync()` latency than the desktop rich client. The goal is to reduce memory consumption, accelerate edit application, and improve perceived responsiveness on the web platform without changing any user-facing features.

## Architectural Context

The current system flows:

```
AI response (text + markdown)
  -> preprocessMarkdown() strips formatting, extracts hints
  -> applyRedlineToOxml() parses OOXML, selects mode:
       FORMAT-ONLY  -> surgical rPr edits on live DOM
       SURGICAL     -> in-place run-level edits (tables, complex structures)
       RECONSTRUCTION -> rebuild paragraphs from diffed RunModel
       LIST EXPANSION -> pipeline: ingest -> diff -> patch -> serialize
  -> serializer.serializeToString(xmlDoc) produces final OOXML
  -> insertOoxml("Replace") pushes to Word
```

Key bottlenecks on the web:
1. **XML parsing/serialization** is far more expensive in browser JS than in desktop's native runtime.
2. **DOM node creation/cloning** is slower in the web sandbox.
3. **`context.sync()` round-trips** to the Word backend are high-latency on the web.
4. **String concatenation and regex** in hot paths create GC pressure.
5. **Repeated work** (re-parsing, redundant traversals, duplicate allocations) is tolerable on desktop but compounds on web.

## Assessment Update (2026-02-08)

After code review of the current reconciliation stack, these priorities are adjusted:

- **Top priority (highest ROI):** `P5.8` traversal reduction, `P5.2` surgical-mode allocation cuts, `P5.9` span splitting algorithm.
- **High priority (targeted):** `P5.4` `context.sync()` reduction in the `agentic-tools` flows, and `P5.1` parse/serialize reductions where double-parse actually occurs (mostly pipeline/list paths).
- **Medium priority:** `P5.6` diff unification/singleton.
- **Selective priority:** `P5.3` string optimizations (focus on structural wins; avoid micro-optimizations without benchmark proof).
- **Conservative priority:** `P5.5` memory changes that mutate object lifecycles (higher regression risk than initially estimated).
- **Already largely addressed:** `P5.7.4` static OOXML package constants are already centralized in `package-builder.js`.

## Implementation Progress Snapshot (2026-02-08)

This section records what has already been implemented in this branch so the work can be continued in a new conversation without re-discovery.

### Status by Workstream

| Workstream | Status | Notes |
|-----------|--------|-------|
| P5.1 Parse/Serialize | Partial | Implemented parsed-DOM threading (`xmlDoc`) into pipeline ingestion/execute path; list path now reuses the already parsed document from `oxml-engine`. Added production/web-gated `validateBasic`, lazy `pPr` serialization (serialize-on-demand in patch/serialize paths), and removed dead serializer-heavy helpers from `surgical-mode.js`. Remaining: broader serializer-hoisting across all edit paths. |
| P5.2 DOM Allocations (Surgical) | Substantially Complete | Single-pass run extraction for surgical full text + spans; removed hot `Array.from` usage in surgical loops; added span indexing for overlap queries to avoid repeated linear scans in insert/delete operations. |
| P5.3 String Ops | Not Started | Deferred pending benchmark-gated proof. |
| P5.4 `context.sync()` Minimization | Partial | Added batched reconciliation API (`applyReconciliationToParagraphBatch`) and integrated batch edit path in `executeRedline` for eligible multi-paragraph `edit_paragraph` operations. Added batched surgical search+insert in OOXML hybrid route (with sequential fallback), removed per-item anchor sync in structured-list direct OOXML insertion, and removed an extra trailing no-op sync in `executeRedline`. More sync-collapsing still needed in complex routes (`modify_text`, `replace_range`, and fallback-heavy branches). |
| P5.5 Memory Footprint | Not Started | No targeted RunModel shape/lifecycle work yet. |
| P5.6 Diff Optimization | Complete (for planned scope) | Added module-level singleton DMP, introduced shared `computeWordDiffs`, switched reconstruction + surgical callers to unified diff path, optimized `charsToWords` to push+join. |
| P5.7 Web Runtime | Partial | Added web-aware event-loop yielding for large pipeline stages, platform-aware diff semantic-cleanup gating for large payloads, production log-level gating in logger adapter, and lazy dynamic import for table reconciliation mode in `oxml-engine`. Remaining: broader lazy-module coverage and additional web-only threshold tuning. |
| P5.8 Format Traversal | Partial | Improved shared extraction usage (`extractFormattingFromOoxml` now returns spans + hints + paragraphs reused by format flows). Full `DocumentIndex` single-pass architecture is still pending. |
| P5.9 `splitSpans` Algorithm | Complete | Replaced iterative convergence loop with sorted single-pass boundary splitting in `format-span-application.js` (O(s+b)-style walk). |

### Additional Tooling Notes (Current Branch State)

- `taskpane.js` includes stronger malformed-function-call recovery for tool invocation payloads (including malformed `edit_list` / `convert_headers_to_list` argument shapes).
- `executeEditList` was intentionally restored to the legacy OOXML surgical implementation that was previously working, per latest direction.
- Some list-tool reliability work is intentionally separated from this performance plan and should be tracked independently.
- Reconciliation logger now supports level gating (`silent`/`error`/`warn`/`info`), defaulting to lower verbosity in production builds.
- `routeChangeOperation` now batches surgical search + OOXML replacements with sequential fallback on batch failure.
- `surgical-mode.js` removed dead serializer-heavy helper paths that were no longer referenced.
- `oxml-engine.js` now lazy-loads `table-mode.js` only when table transformation/reconciliation routes are actually used.

### Validation Completed So Far

- `npm run build:dev` (passes)
- `node tests/list_tests.mjs` (passes)
- Earlier in this optimization run, `tests/phase4/golden-guardrail.mjs` and `tests/phase4/perf-harness.mjs` were exercised; rerun both after the next batch of changes for a clean checkpoint.
- Re-ran after latest optimizations:
  - `node tests/phase4/golden-guardrail.mjs` (passes)
  - `node tests/phase4/perf-harness.mjs` (passes)
- Re-ran again after additional sync/lazy-load changes:
  - `npm run build:dev` (passes)
  - `node tests/phase4/golden-guardrail.mjs` (passes)
  - `node tests/phase4/perf-harness.mjs` (passes)
- Re-ran after the latest `agentic-tools` batching/sync cleanup:
  - `npm run build:dev` (passes)
  - `node tests/phase4/golden-guardrail.mjs` (passes)
  - `node tests/list_tests.mjs` (passes)

### Remaining Work (Current Gap List)

This is the concrete set of items still missing to reach the plan's Definition of Done.

| Workstream | Missing Work |
|-----------|--------------|
| P5.1 Parse/Serialize | Hoist/reuse serializer instances more consistently in non-pipeline paths (format/table/hybrid helpers); verify no avoidable intermediate serialization remains in hot routes. |
| P5.2 DOM Allocations (Surgical) | Run one final targeted pass for residual allocation hotspots under very large table-cell edits (profiling-driven only). |
| P5.3 String Ops | Not started: centralize/trim remaining namespace-stripping and string rewrite hotspots only if perf harness shows measurable wins. |
| P5.4 `context.sync()` Minimization | Collapse remaining sync clusters in `modify_text` range-expansion/retry logic and `replace_range` fallback flows; add turn-scoped OOXML prefetch/cache with explicit invalidation. |
| P5.5 Memory Footprint | Not started: reduce RunModel object-copy churn and lazy-allocate reconstruction maps/hint lookup structures. |
| P5.6 Diff Optimization | Planned scope complete; optional follow-up is gated diff-result caching only if benchmarked hit-rate is meaningful. |
| P5.7 Web Runtime | Expand lazy-loading beyond table mode to other less-common reconciliation branches; tune web thresholds (yield/diff cleanup/validation) with browser-side measurements. |
| P5.8 Format Traversal | Build the full shared `DocumentIndex` single-pass architecture and remove remaining duplicate paragraph/span traversals in format flows. |
| P5.9 `splitSpans` Algorithm | Core algorithm complete; optional deferred work is batched/deferred DOM mutation if profiling shows additional gain. |

## Phase 5 Workstreams

### P5.1 Reduce XML Parse/Serialize Round-Trips

**Impact: Medium-High | Risk: Medium**

XML parse/serialize overhead is real, but its biggest avoidable double-parse cost is concentrated in the pipeline/list path (not every edit route). Current state:

- `applyRedlineToOxml()` parses the full OOXML once (line 48, `oxml-engine.js`).
- `ReconciliationPipeline.execute()` can now accept caller-supplied `xmlDoc`; redundant parse remains only where callers do not pass pre-parsed DOM.
- Dead serializer-heavy helpers formerly in `surgical-mode.js` were removed; serializer hoisting is still incomplete in other reconciliation helpers.
- `extractFormattingFromOoxml()` traverses the already-parsed DOM but `format-application.js` methods then re-traverse `getDocumentParagraphs()` redundantly.
- Validation in `pipeline.js:validateBasic()` is now environment-gated; extra parse is skipped in production web hot paths (`validationMode: auto`).
- Paragraph properties (`pPr`) are lazily serialized instead of eagerly stringified per paragraph in ingestion.

**Actions:**

1. **Thread the parsed `xmlDoc` through the pipeline** instead of passing the raw OOXML string. `ReconciliationPipeline.execute()` should accept an optional pre-parsed `Document` so the engine router can pass its already-parsed DOM directly, eliminating the redundant parse in `ingestOoxml()`.

2. **Hoist serializer creation to the call-site level.** `applyRedlineToOxml()` already creates a serializer (line 43). Pass it down to `applySurgicalMode`, `applyReconstructionMode`, format-application, and table-mode functions rather than having each create their own. This saves ~4-6 serializer allocations per operation.

3. **Rework `validateBasic()` usage by environment.** Keep strict string re-parse validation in test/debug builds, but avoid full re-parse in production web hot paths unless a failure fallback is triggered.

4. **Lazy-serialize `pPr` in ingestion.** `ingestion-paragraph.js:108` calls `serializeXml(pPr)` for every paragraph to produce `pPrXml`. In many paths (format-only, surgical), this string is never consumed. Store the DOM `pPr` reference and serialize on demand via a getter or at serialization time only.

5. **Defer final serialization.** In surgical and reconstruction modes, `serializer.serializeToString(xmlDoc)` serializes the entire document at the end. If the modified DOM is immediately passed to `insertOoxml`, consider whether the Word API can accept a DOM node directly. If not (likely), this serialization is unavoidable -- but at minimum, ensure we're not serializing intermediate states unnecessarily.

**Estimated impact:** ~5-15% overall; can be ~15-30% in pipeline/list-heavy scenarios where duplicate parse paths are exercised repeatedly.

---

### P5.2 Reduce DOM Node Allocations in Surgical Mode

**Impact: High | Risk: Medium**

Surgical mode (`surgical-mode.js`) is the workhorse for table cell edits and any document with `w:tbl`. It creates many DOM nodes:

- `processRunElement()` (line 252) calls `Array.from(r.childNodes)` for every run, allocating an intermediate array.
- `getUpdatedFullText()` (line 316) does the same traversal again, duplicating work.
- `reconcileFormattingForTextSpan()` (line 189) calls `Array.from(rPr.childNodes)` to check existing formats.
- `processDelete()` and `processInsert()` each use `textSpans.filter(...)` (lines 344, 413-427) which is O(n) per diff segment.

**Actions:**

1. **Merge `processRunElement` and `getUpdatedFullText` into a single pass.** These two functions iterate the same child nodes separately. Combine them to build both the text span metadata and the full text string in one traversal. This halves the childNode iteration count.

2. **Replace `Array.from(childNodes)` with direct `for` loops.** `Array.from()` allocates a new array on each call. In hot loops (every run in every paragraph), iterate `childNodes` directly using `child.firstChild`/`child.nextSibling` or `for (let i = 0; i < node.childNodes.length; i++)`.

3. **Index textSpans for range queries.** Currently `textSpans.filter(s => s.charEnd > startPos && s.charStart < endPos)` runs a full linear scan per diff operation. Pre-sort spans by `charStart` and use binary search to find the relevant range, then iterate only the overlapping subset. This reduces `processDelete` and `processInsert` from O(n*d) to O(d * log(n) + k) where k is the overlap count.

4. **Replace `hasElement()` array allocation with direct child traversal.** Keep semantics tight to direct `w:rPr` children (do not widen with deep `getElementsByTagName` unless intended). Use a `firstChild/nextSibling` loop or indexed child iteration.

5. **Pool `diff_match_patch` instances.** Both surgical mode (line 130) and reconstruction mode (line 28) create a `new diff_match_patch()` per call. Create a single module-level instance and reuse it.

**Estimated impact:** 20-40% reduction in DOM allocations during surgical edits. Span indexing provides the biggest single win for documents with many runs.

---

### P5.3 Optimize String Operations in Hot Paths

**Impact: Low-Medium | Risk: Low**

Several hot paths rely on regex replacements and string concatenation that create GC pressure:

- Namespace stripping: `.replace(/\s+xmlns:[^=]+="[^"]*"/g, '')` appears in at least 8 locations across `serialization.js`, `surgical-mode.js`, `format-application.js`, and `oxml-engine.js`. Each call runs a regex on potentially large XML strings.
- `escapeXml()` in `serialization.js` is called per-run (once in `buildSimpleRun`, once in each `buildDeletionXml`/`buildInsertionXml`). Each call creates intermediate strings.
- `sanitizeAiResponse()` runs 4 sequential regex replacements on the entire modified text.
- The `injectFormatting()` and `applyFont()` functions in `serialization.js` do repeated `includes()` checks and string concatenation per run.
- `serializeToOoxml()` joins paragraphs via `paragraphs.join('')` and accumulates runs via `currentRuns.join('')`, creating intermediate arrays and strings.

**Actions:**

1. **Strip namespaces once at serialization time, not per-fragment.** Instead of stripping `xmlns:` attributes from every small XML fragment during construction, strip them once from the final serialized OOXML string before returning. Add a single `stripNamespaces()` call at the output boundary of `applyRedlineToOxml()` and remove the ~8 inline `.replace()` calls.

2. **Use array-push + single join for OOXML construction.** In `serializeToOoxml()`, replace the pattern of `currentRuns.push(string)` then `currentRuns.join('')` with a single shared output array. Push all XML fragments (paragraph opens, runs, paragraph closes) to one flat array and join once at the end.

3. **Precompile and centralize regex usage.** Hoist namespace/sanitization regexes to module-level constants and avoid re-declaring ad hoc patterns in hot paths.

4. **Reduce repeated run-property string rewrites.** In `applyFont()` and `injectFormatting()`, avoid repeated `includes()` scans and repeated full-string rebuilds when no formatting/font changes are required.

5. **Defer micro-optimizations unless benchmarked.** Template-literal rewrites and `escapeXml` caching should only ship if perf harness shows measurable wins.

**Estimated impact:** ~3-8% in realistic end-to-end runs when combined with structural serialization improvements.

---

### P5.4 Minimize `context.sync()` Round-Trips

**Impact: High | Risk: Low**

Each `context.sync()` call in Word Online is a network round-trip to the Word backend. The integration layer currently uses:

- `integration.js:30-31`: `paragraph.getOoxml()` + `await context.sync()` to fetch OOXML.
- `integration.js:52-53`: `paragraph.insertOoxml(...)` + `await context.sync()` to apply changes.

The `agentic-tools` command flow has many additional sync calls for context extraction, paragraph location, track-mode toggles, and apply-redline fallback paths. Each sync adds 50-200ms of latency on web.

**Actions:**

1. **Batch read + write into a single `Word.run` where possible.** If the caller already has a `Word.run` context, avoid creating nested contexts. Ensure `applyReconciliationToParagraph()` receives and reuses the caller's context rather than potentially nesting.

2. **Batch multiple paragraph edits.** When the AI applies edits to multiple paragraphs in a single tool call, batch all `paragraph.getOoxml()` requests before the first `context.sync()`, process all reconciliation work locally, then batch all `paragraph.insertOoxml()` calls before a final `context.sync()`. This reduces 2N syncs to 2.

3. **Prioritize high-volume `agentic-tools` paths first.** Profile and collapse sync clusters in `applyModifyTextOperation` and multi-paragraph edit flows before touching lower-frequency helpers.

4. **Prefetch/caching with invalidation.** During context extraction, opportunistically cache OOXML for likely targets. If the same paragraph is edited repeatedly in one AI turn, reuse OOXML unless document state changes invalidate it.

5. **Explore `Range`-based operations.** For multi-paragraph edits, fetching a range's OOXML once (covering all target paragraphs) and processing it as a unit may be faster than per-paragraph fetches, reducing sync count and XML overhead.

**Estimated impact:** Reducing sync calls from 2N to 2 for batch operations could save 100-400ms per multi-paragraph edit on web.

---

### P5.5 Reduce Memory Footprint of RunModel and Intermediate Structures

**Impact: Medium | Risk: Medium**

The RunModel abstraction creates an object per run with spread copies (`{...run}`) at multiple pipeline stages. For a large document with hundreds of runs, this creates significant memory pressure:

- `splitRunsAtDiffBoundaries()` creates new run objects via `{...run}` spread (lines 43-48, 61-66 in `patching.js`).
- `applyPatches()` creates another copy of every run via `{...run}` (lines 98, 104, 139, 147).
- `buildReconstructionMapping()` creates `propertyMap`, `paragraphMap`, `sentinelMap`, `containerFragments`, and `replacementContainers` -- all allocated upfront even when most entries are never accessed.
- Format hint processing in `getApplicableFormatHints()` runs per-run, iterating all hints each time.

**Actions:**

1. **Reduce copies only where ownership is local and explicit.** Avoid broad in-place mutation of shared run objects unless invariants are documented and guarded by tests.

2. **Replace high-frequency spread copies with targeted construction.** Build compact new entries with only required fields on hot paths, rather than cloning entire objects for one-field changes.

3. **Shrink RunModel entry shape.** Several fields are optional (e.g., `author`, `containerContext`, `containerId`, `propertiesXml`). Avoid attaching undefined fields. Consider using a class with defined slots rather than plain objects so the engine can optimize property access.

4. **Lazy-allocate reconstruction maps.** In `buildReconstructionMapping()`, `containerFragments`, `replacementContainers`, and `sentinelMapByStart` are allocated upfront. If most edits are single-paragraph with no sentinels, many of these are empty. Allocate lazily -- create the Map/fragment only when the first entry is added.

5. **Pre-filter format hints once per pipeline run.** Instead of calling `getApplicableFormatHints()` per-run (which scans all hints), pre-bucket hints by offset range at the start of serialization, then look up per-run via a cursor. This is analogous to the `createRangeCursorLookup` already used for diff ops.

**Estimated impact:** ~8-15% peak memory reduction with lower regression risk than broad in-place mutation.

---

### P5.6 Optimize Diff Computation

**Impact: Medium | Risk: Low**

The diff engine is invoked in multiple code paths, sometimes redundantly:

- `computeWordLevelDiffOps()` in `diff-engine.js` creates a new `diff_match_patch()` instance per call (line 120).
- `wordsToChars()` builds a `wordHash` Map and `wordArray` from scratch each time.
- The same `wordsToChars` + `diff_main` + `diff_cleanupSemantic` + `charsToWords` sequence is repeated identically in surgical mode (lines 130-134, `surgical-mode.js`), reconstruction mode (lines 28-32, `reconstruction-mode.js`), and the pipeline path (lines 78-80, `pipeline.js` via `computeWordLevelDiffOps`).
- `charsToWords()` iterates characters via `charCodeAt()` per character, building strings by concatenation.

**Actions:**

1. **Share a single `diff_match_patch` instance module-wide.** Create it once at module load. The library is stateless between calls, so a singleton is safe.

2. **Add diff-result caching only behind a measured gate.** For the same `(originalText, modifiedText)` pair, diff is deterministic, but cache hit rate may be low in normal edits. Ship only if perf harness shows repeat-hit benefit.

3. **Optimize `charsToWords` string building.** Replace character-by-character concatenation with an array-push-then-join approach:
   ```js
   const parts = [];
   for (let i = 0; i < chars.length; i++) {
     parts.push(wordArray[chars.charCodeAt(i)]);
   }
   return parts.join('');
   ```

4. **Unify the diff call pattern.** Create a single `computeWordDiffs(text1, text2)` utility that returns the `[op, text][]` tuples. All three callers (surgical, reconstruction, pipeline) should use this instead of duplicating the `wordsToChars -> diff_main -> cleanupSemantic -> charsToWords` sequence. This also makes it trivial to add caching or profiling at a single point.

**Estimated impact:** 5-10% in diff-heavy operations, mostly from singleton + call-path unification.

---

### P5.7 Web-Specific Runtime Optimizations

**Impact: Medium | Risk: Low**

Word Online runs in a constrained browser sandbox. Several adaptations can improve behavior specifically on this platform:

**Actions:**

1. **Yield to the event loop for large operations.** On the web, long-running synchronous JavaScript blocks the UI thread, causing the add-in to appear frozen. For operations touching >50 runs or >5000 characters, insert periodic `await new Promise(r => setTimeout(r, 0))` yields at natural pipeline stage boundaries (between ingestion, diff, patch, serialize). This keeps the UI responsive without changing the logic.

2. **Detect platform and adjust thresholds.** Add a lightweight platform check (`Office.context.platform === 'OfficeOnline'` or equivalent). On web:
   - Lower the validation threshold (or skip `validateBasic()` re-parse entirely, per P5.1.3).
   - Reduce logging verbosity (each `log()` call has overhead in browser console).
   - Consider disabling `diff_cleanupSemantic()` for very large diffs, as its quadratic worst case is more problematic on web.

3. **Lazy-load reconciliation modules.** The reconciliation engine imports `diff-match-patch`, table reconciliation, comment engine, list generation, and more at module load time. On web, this inflates initial load time. Use dynamic `import()` for less-common paths (table-to-text transformation, comment injection) so they load only when needed.

4. **No-op unless regression is found:** package boilerplate constants are already centralized in `package-builder.js`; retain this item as verification, not implementation.

5. **Minimize `console.log` in production.** The logger adapter (`logger.js`) appears to pass through to `console.log`. On web, excessive console output can measurably slow execution. Add a log-level check so `log()` calls are no-ops in production builds.

**Estimated impact:** 10-20% improvement in perceived responsiveness on web. The yield-to-event-loop change alone prevents the "not responding" experience for large edits.

---

### P5.8 Eliminate Redundant Traversals in Format Application

**Impact: High | Risk: Medium**

Format-only and surgical format paths traverse the document multiple times:

- `applyRedlineToOxml()` calls `extractFormattingFromOoxml()` (line 76) which traverses all paragraphs/runs to find existing formatting.
- It then calls `detectTableCellContext()` (lines 60, 91, 133) which traverses for table wrappers.
- Format-only path calls `getDocumentParagraphs()` again (line 172, `format-application.js`), then `buildTextSpansFromParagraphs()` which re-traverses runs.
- `buildParagraphInfos()` traverses spans again.
- `findTargetParagraphInfo()` or `findMatchingParagraphInfo()` searches through paragraph infos.

This results in 4-5 full traversals of the same DOM for a format-only change.

**Actions:**

1. **Build a shared document index once.** Create a `DocumentIndex` structure at the start of `applyRedlineToOxml()` that captures in a single pass:
   - All paragraphs with their text content
   - All text spans with offset mappings
   - Table context (is this inside a table cell?)
   - Existing formatting hints
   - Paragraph properties

   Pass this index to all downstream functions instead of having each re-traverse the DOM.

2. **Merge `extractFormattingFromOoxml` and `buildTextSpansFromParagraphs`.** These functions do nearly identical work (iterate runs, extract text, note formatting). They should be a single pass that returns both `textSpans` and `existingFormatHints`.

3. **Cache `getDocumentParagraphs()` result.** This function is called multiple times within a single `applyRedlineToOxml()` invocation. Cache the result on the `xmlDoc` or pass it as a parameter.

4. **Pre-compute paragraph text for matching.** `findMatchingParagraphInfo()` and `findTargetParagraphInfo()` compare paragraph text to `originalText`. Build a text -> paragraph index upfront rather than checking each paragraph sequentially.

**Estimated impact:** 20-35% improvement for format-only operations, which are among the most common edit types.

---

### P5.9 Optimize `splitSpansAtBoundaries` Algorithm

**Impact: Medium | Risk: Low**

The current implementation in `format-span-application.js` uses a multi-pass approach:

```js
while (splitsOccurred) {
    splitsOccurred = false;
    for (const span of currentSpans) {
        for (const boundary of sortedBoundaries) { ... }
    }
}
```

This is O(s * b * p) where s=spans, b=boundaries, p=passes. For a paragraph with 20 runs and 10 format hints (20 boundaries), this is manageable. But for complex formatting, it can spike.

**Actions:**

1. **Single-pass merge-sort split.** Since both spans and boundaries are sorted, walk them together in a single pass:
   ```
   for each span in sorted spans:
     while next boundary falls within span:
       split span at boundary
       advance to right half
     emit span
   ```
   This is O(s + b) and eliminates the iterative convergence loop entirely.

2. **Avoid DOM manipulation during splitting.** Currently `splitSpanAtOffset()` creates new DOM nodes (`createTextRun`) and modifies the parent. Defer DOM modifications until all splits are computed, then apply them in batch. This reduces DOM mutation count and avoids layout thrashing.

**Estimated impact:** 5-15% for heavily-formatted content. Low-effort since the algorithm change is straightforward.

---

## Sequencing and Dependencies

```
Phase │ Workstream                        │ Depends On │ Risk
──────┼───────────────────────────────────┼────────────┼────────
  1   │ P5.8 Format traversal reduction   │ None       │ Medium
  2   │ P5.2 DOM allocations surgical     │ None       │ Medium
  3   │ P5.9 splitSpans algorithm         │ P5.8       │ Low
  4   │ P5.1 Reduce parse/serialize       │ None       │ Medium
  5   │ P5.4 context.sync minimization    │ None       │ Low
  6   │ P5.6 Diff optimizations           │ None       │ Low
  7   │ P5.3 String optimizations         │ None       │ Low
  8   │ P5.5 Memory footprint             │ P5.6       │ Medium
  9   │ P5.7 Web-specific runtime opts    │ P5.3       │ Low
```

`P5.8 + P5.2 + P5.9` are the fastest path to meaningful web wins in current code paths. `P5.1` should be targeted to true duplicate parse locations (mainly pipeline/list flows). `P5.4` can proceed in parallel at the integration/command layer.

## Validation Strategy

1. **Golden guardrail regression** (`tests/phase4/golden-guardrail.mjs`): All existing baselines must continue to pass after each workstream.

2. **Performance harness** (`tests/phase4/perf-harness.mjs`): Extended with web-relevant scenarios:
   - Add a "simulated web" mode that inserts `performance.now()` timing around each pipeline stage.
   - Add memory snapshot tracking (via `process.memoryUsage()` in Node, `performance.memory` in browser).
   - New test cases:
     - Format-only change on 50-run paragraph (targets P5.8)
     - Batch 5-paragraph edit (targets P5.4)
     - Large table cell edit (targets P5.2)

3. **Existing test suites**: All tests from Phase 4 validation must pass:
   - `standalone_smoke`, `no_word_api_standalone_check`, `include_numbering_behavior`
   - `comment_tests`, `table_tests`, `list_tests`, `integration_tests`
   - `highlight_tests`, `formatting_tests`

4. **Web-specific validation**: Manual testing in Word Online to verify:
   - No "add-in not responding" dialogs during edits
   - Perceived edit latency < 2 seconds for single paragraph changes
   - No visible jank or UI freezes during multi-paragraph operations

5. **Budget gates**:
   - No regression > 5% on any existing perf baseline
   - Target aggregate 20-40% improvement in end-to-end edit latency on web
   - Target 15-25% reduction in peak memory during reconciliation

## Summary of Expected Gains

| Workstream | CPU Improvement | Memory Improvement | Primary Benefit |
|-----------|----------------|-------------------|-----------------|
| P5.1 Parse/Serialize | 5-15% overall (up to 30% in list-heavy paths) | 5-10% | Fewer XML parses |
| P5.2 DOM Allocations | 10-20% | 20-40% | Less node creation |
| P5.3 String Ops | 3-8% | 5-10% | Less GC pressure |
| P5.4 context.sync | 20-50% latency | -- | Fewer round-trips |
| P5.5 Memory Footprint | 3-8% | 8-15% | Smaller RunModel |
| P5.6 Diff Optimization | 5-10% | 5% | Singleton + cache |
| P5.7 Web Runtime | 10-20% perceived | 5-10% | UI responsiveness |
| P5.8 Format Traversal | 20-35% (format) | 10% | Single-pass index |
| P5.9 splitSpans | 5-15% (format) | 5% | O(s+b) algorithm |

**Aggregate target:** 20-40% reduction in end-to-end edit latency on Word Online, with 15-25% memory reduction during reconciliation operations.

## Definition of Done

1. All 9 workstreams implemented with no feature regressions.
2. Golden guardrail and perf harness pass at all stages.
3. Measurable improvement confirmed via perf harness and manual web testing.
4. No new `context.sync()` calls added; existing ones reduced where batching applies.
5. Memory profiling shows reduced peak allocation during representative operations.
