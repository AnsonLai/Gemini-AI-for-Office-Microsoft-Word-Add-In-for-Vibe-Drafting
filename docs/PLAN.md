# Robustness & Reliability Plan — Phase 2

> **Audience note:** this plan is written to be executed by a smaller/cheaper model,
> one work package at a time. It is deliberately prescriptive: exact file paths,
> code skeletons, search anchors, test names, and per-WP acceptance criteria.
> When a line number has drifted, search for the quoted symbol name instead.
> When the plan and the code disagree, the code wins — note the deviation in the
> WP's status entry (see how `plans/reliability-hardening-plan.md` recorded
> deviations) rather than forcing the plan's literal text.

## Objective

Phase 1 (`plans/reliability-hardening-plan.md`, complete) hardened the **redline
pipeline**: anchors, sanitizer, model profiles, checkpoints, history invariants,
truncation salvage, corrective retry. Yet the project still feels fragile.
Phase 2 attacks the reasons *why* incidents keep happening despite those defenses:

1. **Regressions land silently** — there is no `npm test`, no CI; verification is
   ad hoc and the live user is the test harness. (Proof: `tests/addin/migrated_tool_cutover_tests.mjs`
   is failing in the tree *right now* and nobody noticed — see WP1 step 0.)
2. **Hardening was applied unevenly** — the non-redline tools still run on the
   pre-hardening fragility profile (see the call-site table in WP2).
3. **When a defense fires, nobody knows** — everything logs to a console nobody is
   watching, so every incident is a fresh from-scratch investigation.
4. **The core loop is untestable** — the whole agentic loop lives inside one
   ~1,280-line function coupled to Office.js and the DOM.
5. **Model behavior changes are discovered live** — the eval harness is manual,
   costs money, and is effectively never run.
6. **The Word bridge — where the real failures happen — has no repeatable
   verification** ("not verified in real Word" recurs across every plan).
7. **Entropy is unchecked** — versions, docs, and lint drift with no gate.

## Status (created 2026-07-05; expanded same day for smaller-model execution)

| WP | Title | Status |
|----|-------|--------|
| WP1 | Single-command test runner + CI gate | ☐ Todo |
| WP2 | One Gemini client: shared error taxonomy + retry policy | ☐ Todo |
| WP3 | Reliability event log (observability) | ☐ Todo |
| WP4 | Extract testable logic out of `taskpane.js` | ☐ Todo |
| WP5 | Record/replay evals + incident regression cases | ☐ Todo |
| WP6 | Word-bridge verification kit (mock harness + smoke checklist) | ☐ Todo |
| WP7 | Package-boundary contract tests | ☐ Todo |
| WP8 | Drift control (version, docs, lint ratchet) | ☐ Todo |

When you finish a WP, change its row to ✅ Done and add a short indented note
under this table recording what was actually built and any deviations.

## Global rules for the implementing model

**Read this section before every WP. These rules override your defaults.**

1. **One WP per session/commit.** Complete the WP fully (code + tests green),
   run the full regression suite, then commit with the message suggested in the
   WP. Do not start the next WP in the same commit.
2. **Regression check** = `npm test` once WP1 lands. Before WP1 exists, run each
   file under `tests/` individually with `node tests/<name>.mjs` (all pass today
   except the one WP1 step 0 fixes) plus `npm run build:dev`.
3. **New modules must be importable in plain Node** (v24): no `Office`, `Word`,
   `document`, `localStorage`, or `fetch` referenced at module top level. Read
   them lazily inside functions or accept them as injected parameters. Follow
   the existing pattern in `src/taskpane/modules/commands/change-validation.js`
   and `src/taskpane/modules/storage/checkpoint-store.js`.
4. **New modules must be lint-clean** (`npm run lint` reports nothing for them):
   double quotes per prettier config, trailing commas, `/* global ... */`
   headers for any browser globals used inside functions.
5. **Existing big files are NOT lint-clean** (`taskpane.js`,
   `agentic-tools.js`, `chat-history.js` carry hundreds of pre-existing
   violations). When editing them, match their existing idiom (single quotes)
   and NEVER run `lint:fix` / prettier on them — that happens once, deliberately,
   in WP8 step 3. A reformat before then will bury your real diff.
6. **When extracting code, port it verbatim** — copy the exact source text, do
   not retype or "improve" it. Behavior changes hide in retyped regexes. After
   extraction, the old call site imports the new module and the old inline copy
   is deleted (never leave two copies).
7. **Tests are plain Node scripts**: `import assert from 'assert'`, functions
   named `testXxx()`, a `run()` that calls each and `console.log('PASS: ...')`,
   exit code 1 on failure (assert throws). No test framework. Copy the structure
   of `tests/change_validation_tests.mjs`.
8. **`src/**/*.js` files are ESM without `"type": "module"`** in package.json.
   Node prints a `MODULE_TYPELESS_PACKAGE_JSON` warning when importing them —
   that warning is normal, not an error.
9. **Do not touch** `@ansonlai/docx-redline-js` internals, the prompt text in
   `redline-prompt.js`, tool schemas, or default models — unless a WP explicitly
   says so.
10. **If a step is impossible as written** (file moved, symbol renamed), find the
    real location with Grep, do the equivalent, and record the deviation. Do not
    skip the step silently.

## Verified facts about the current code (2026-07-05)

You can rely on these without re-deriving them:

- `src/taskpane/taskpane.js` is 2,519 lines. `sendChatMessage` spans lines
  1148–2428 and contains: context extraction, the agentic loop, the 4-tier
  history recovery ladder (lines ~1686–1731), malformed-function-call recovery
  parsers (lines ~1784–1938), the 10-branch tool dispatch (lines ~2049–2269),
  the loop guard (lines ~2328–2355), and the outer error handler (~2382–2427).
- `src/taskpane/modules/commands/agentic-tools.js` is 1,677 lines. Its Word/
  settings dependencies are injected via `initAgenticTools(deps)` (line 48):
  `loadApiKey, loadModel, loadSystemMessage, loadRedlineSetting,
  loadRedlineAuthor, setChangeTrackingForAi, restoreChangeTracking,
  SAFETY_SETTINGS_BLOCK_NONE, API_LIMITS`.
- There are **four** independently implemented Gemini fetch paths (see WP2).
- Tests: 18 files directly under `tests/` (all pass), 4 under `tests/addin/`
  (3 pass, `migrated_tool_cutover_tests.mjs` FAILS — fixed in WP1 step 0),
  1 Node-runnable file under `tests/word-desktop/` (`list-regression.mjs`,
  passes), and the manual eval harness under `tests/evals/`.
  `tests/setup-xml-provider.mjs` is a shared helper, not a suite.
- `tests/evals/run-evals.mjs` already exports `buildAnchoredText` and
  `scoreChangeSet` and guards its `main()` behind an
  `process.argv[1]` check, so it is import-safe.
- `webpack.config.js` does not currently `require("webpack")` (needed for
  DefinePlugin in WP8). `npm run build:dev` works headless: the dev-cert
  code path only runs for `webpack serve`.
- Tool-result shape conventions in the dispatch (needed for WP4):
  `apply_redlines` / `insert_comment` / `highlight_text` return
  `{ message, showToUser }`; `edit_list` / `insert_list_item` / `edit_table` /
  `edit_section` / `convert_headers_to_list` return `{ message, success }`;
  `perform_research` / `navigate_to_section` return a plain string.

---

## WP1 — Single-command test runner + CI gate

**Goal:** one command runs every check; a machine runs it on every push. A
regression can no longer land silently.

**Why:** 23 runnable test files exist but nothing runs them automatically — and
one (`tests/addin/migrated_tool_cutover_tests.mjs`) is already failing without
anyone noticing. Past drift: README once documented four phantom suites; a
broken package import sat in the tree for weeks. Every later WP adds tests;
without a runner + CI they rot the same way.

**Suggested commit message:** `WP1: add npm test runner + GitHub Actions CI`

### Step 0 — fix the currently-failing suite (pre-existing, 5 minutes)

`tests/addin/migrated_tool_cutover_tests.mjs` fails because Phase 1 moved the
bridge call out of `executeRedline` into `applyRedlineChangeSet`
(`agentic-tools.js:83`; the call to `applyRedlineChangesToWordContext(` is now
at line ~133, and `executeRedline` calls `applyRedlineChangeSet(` at ~219).
The test still asserts the bridge call is inside `executeRedline`. Replace its
`testRedlineCutover` with:

```js
function testRedlineCutover() {
    const executeBody = extractFunctionBody('executeRedline');
    const applyBody = extractFunctionBody('applyRedlineChangeSet');
    assertContains(
        executeBody,
        'applyRedlineChangeSet(',
        'executeRedline should route through applyRedlineChangeSet'
    );
    assertContains(
        applyBody,
        'applyRedlineChangesToWordContext(',
        'applyRedlineChangeSet should route through shared redline runner'
    );
    assertNotContains(
        executeBody,
        'routeChangeOperation(',
        'executeRedline should not route through legacy command-level routeChangeOperation logic'
    );
}
```

Run `node tests/addin/migrated_tool_cutover_tests.mjs` — must print
`PASS: migrated tool cutover tests`.

### Step 1 — create `tests/run-all.mjs`

```js
/**
 * run-all.mjs — discovers and runs every regression suite; one PASS/FAIL table.
 * Usage: npm test   (alias for: node tests/run-all.mjs)
 *
 * Discovery rules:
 *   - every .mjs directly under tests/ except this file and setup-xml-provider.mjs
 *   - every .mjs directly under tests/addin/
 *   - tests/word-desktop/list-regression.mjs (explicit; the other files there
 *     are PowerShell tooling)
 *   - tests/evals/ is EXCLUDED here (paid API); WP5 adds a --replay entry.
 */
import { spawnSync } from "child_process";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const EXCLUDED = new Set(["run-all.mjs", "setup-xml-provider.mjs"]);

function discover(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs
    .readdirSync(dir)
    .filter((f) => f.endsWith(".mjs") && !EXCLUDED.has(f))
    .sort()
    .map((f) => path.join(dir, f));
}

const files = [
  ...discover(__dirname),
  ...discover(path.join(__dirname, "addin")),
  path.join(__dirname, "word-desktop", "list-regression.mjs"),
];

let failed = 0;
for (const file of files) {
  const rel = path.relative(path.join(__dirname, ".."), file);
  const res = spawnSync(process.execPath, [file], {
    encoding: "utf8",
    timeout: 120000,
  });
  const pass = res.status === 0;
  console.log(`${pass ? "PASS" : "FAIL"} ${rel}`);
  if (!pass) {
    failed += 1;
    console.log("--- output ---");
    console.log((res.stdout || "") + (res.stderr || ""));
    console.log("--------------");
  }
}

console.log(`\n${files.length - failed}/${files.length} suites passed.`);
if (failed > 0) process.exitCode = 1;
```

### Step 2 — add the npm script

In `package.json` `"scripts"`, add (keep alphabetical-ish placement near
`"stop"`):

```json
"test": "node tests/run-all.mjs",
```

### Step 3 — create `.github/workflows/ci.yml`

```yaml
name: CI
on:
  push:
  pull_request:
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 24
          cache: npm
      - run: npm ci
      - run: npm test
      - run: npm run build:dev
```

Why the build step matters: `taskpane.js` is imported by **no test** (it is
Office.js-coupled), so the webpack build is currently the only automated check
that its 2,519 lines parse and its imports resolve.

Do NOT add `npm run lint` to CI yet — it fails on pre-existing debt in the
three legacy files. That gate arrives in WP8.

### Tests / verification

1. `npm test` → all suites PASS, exit code 0 (`echo $LASTEXITCODE` /
   `echo $?`).
2. Temporarily change one assertion in any suite to be false → `npm test`
   exits 1 and prints the failure output. Revert.
3. Create an empty `tests/zz_dummy_tests.mjs` containing
   `console.log("PASS")` → it appears in the table with no runner edits.
   Delete it.
4. Push (or open a PR) and confirm the CI run is green.

### Acceptance criteria

- `npm test` runs 23 suites (18 root + 4 addin + 1 word-desktop) and exits 0.
- CI is green on main; a deliberately broken test turns it red.
- New `tests/*_tests.mjs` files are picked up automatically.

---

## WP2 — One Gemini client: shared error taxonomy + retry policy

**Goal:** every Gemini call goes through one client with one tested
error-classification and retry policy. Transient failures (429/5xx/network
blips) stop surfacing as hard tool failures.

**Why:** there are **four** separately implemented fetch paths with four
different (and individually incomplete) policies:

| # | Call site | Timeout | Retry | Profile-driven | Truncation salvage |
|---|---|---|---|---|---|
| 1 | `callGeminiWithRetry` (`taskpane.js:2431`, used by the chat loop at :1676) | 60s/attempt | 5xx, timeouts — **and, by accident, 4xx** (the generic catch retries everything except the history-invariant string) | yes | n/a |
| 2 | Glance checks (`runGlanceChecks`, `taskpane.js:936` builds its own URL, fetch at :966) | none | none | model only | n/a |
| 3 | `callGeminiForDiffs` (`agentic-tools.js:275`) | 90s | **none** | yes | yes |
| 4 | `callGeminiForJSON` (`agentic-tools.js:680`) | **none** | **none** | **no** (hardcoded `temperature: 0.2`, `maxOutputTokens: 48000`) | **no** (bare `JSON.parse`) |

Concrete consequences today:

- A single transient 5xx during a diff call fails the whole `apply_redlines`
  tool (`callGeminiForDiffs` returns `null` → `TOOL_FAILURE invalid_response`).
- A 429 (routine on preview models — the exact models that carry
  `previewThrottleWarning`) is never retried anywhere, even though Google sends
  `Retry-After`.
- An invalid API key is *retried 3 times with backoff* by path 1 (pointless
  30s of waiting) because non-5xx errors fall into the generic retry catch.
- Comment/highlight/navigate/list/table secondary calls (path 4) can hang
  indefinitely on a stalled connection and die on truncated JSON — the exact
  failure modes Phase 1 already solved for the redline path.
- Error classification is string-matching Google's prose in two places
  (`taskpane.js:1681–1684`, `taskpane.js:2470`); a wording change breaks the
  recovery ladder silently.
- The API key is in the URL query string (`?key=...`) in all four paths, so it
  leaks into any logged error text.

**Suggested commit message:** `WP2: unify Gemini calls behind gemini-client.js with typed errors + retry policy`

### Step 1 — create `src/taskpane/modules/api/gemini-client.js`

Lint-clean, Node-importable. Skeleton (fill in the bodies; keep the exported
names exactly as shown — tests and later WPs depend on them):

```js
/*
 * gemini-client.js — the ONLY module that talks to the Gemini API.
 * Pure/Node-importable: fetch is injectable for tests and read lazily
 * from globalThis otherwise.
 */

export const GEMINI_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/models";

export class GeminiClientError extends Error {
  constructor(kind, detail, { status, retryAfterMs } = {}) {
    super(`Gemini ${kind}: ${detail}`);
    this.name = "GeminiClientError";
    this.kind = kind; // one of ERROR_KINDS
    this.status = status;
    this.retryAfterMs = retryAfterMs;
  }
}

export const ERROR_KINDS = Object.freeze([
  "auth",              // bad/expired key: 400 API_KEY_INVALID, 401, 403
  "quota",             // 429
  "server",            // 5xx
  "history_invariant", // the function-call/response-turn 400s
  "timeout",           // our AbortController fired
  "cancelled",         // caller's cancelSignal fired (user pressed cancel)
  "network",           // fetch threw TypeError (offline, DNS, CORS)
  "http",              // other non-OK statuses
  "unknown",
]);

/**
 * Classify a failed call. Pure function; unit-tested.
 * @param {{status?: number, bodyText?: string, error?: Error}} input
 * @returns {{kind: string, retryable: boolean, retryAfterMs?: number, detail: string}}
 */
export function classifyGeminiError({ status, bodyText = "", error }) {
  // history_invariant — BOTH strings currently checked at taskpane.js:1682-1683:
  //   "function response turn comes immediately after a function call turn"
  //   "function call turn comes immediately after a user turn or after a function response turn"
  // auth — status 401/403, or status 400 with "API_KEY_INVALID" in body
  // quota — status 429 (retryable; parse Retry-After if present)
  // server — 500-599 (retryable)
  // network — error instanceof TypeError (retryable)
  // timeout/cancelled — distinguished by the CALLER (it knows which signal
  //   fired); classify plain AbortError as "timeout" here.
  // everything else with a status — "http" (not retryable)
}

/** Exponential backoff with full jitter. attempt is 0-based. */
export function backoffMs(attempt, baseMs = 1000, randomFn = Math.random) {
  return Math.floor(randomFn() * baseMs * Math.pow(2, attempt));
}

/**
 * POST a generateContent payload.
 * @param {object} opts
 *   model, apiKey            - used to build URL + x-goog-api-key header
 *   payload                  - the JSON body
 *   timeoutMs                - per-attempt timeout (required)
 *   maxRetries               - total attempts = maxRetries (match old semantics
 *                              of callGeminiWithRetry's `retries` param)
 *   cancelSignal             - optional AbortSignal for user cancellation
 *   fetchImpl                - optional; defaults to globalThis.fetch
 *   onRetry                  - optional (kind, attempt) => void   [WP3 hook]
 * @returns parsed response JSON
 * @throws GeminiClientError
 */
export async function callGemini(opts) {
  // Per attempt:
  //  - if cancelSignal already aborted -> throw kind "cancelled"
  //  - AbortController + setTimeout(timeoutMs); also abort when cancelSignal fires
  //    (remove the listener in finally — see taskpane.js:2448-2464 for the
  //    existing pattern to preserve)
  //  - fetch(`${GEMINI_BASE_URL}/${model}:generateContent`, { headers:
  //      { "Content-Type": "application/json", "x-goog-api-key": apiKey }, ... })
  //  - !response.ok -> classify({status, bodyText: await response.text()})
  //  - retryable && attempts remain -> wait quota's retryAfterMs (cap 30000ms)
  //    or backoffMs(attempt), call onRetry, continue
  //  - not retryable, or attempts exhausted -> throw GeminiClientError
  //  - fetch threw: AbortError -> cancelled if cancelSignal.aborted else timeout
  //    (timeout is retryable); TypeError -> network (retryable)
}

/** Map an error kind to the user-facing chat message. One place, not five. */
export function userMessageForError(kind, { previewThrottleWarning = false } = {}) {
  // auth    -> "Your Gemini API key was rejected. Please check it in Settings."
  // quota   -> "The AI service is rate-limiting requests. Wait a moment and try
  //             again, or switch models in Settings."
  // timeout -> previewThrottleWarning
  //            ? "This model is in preview and has likely been throttled. Please go into settings and revert to Gemini 2.5."
  //            : "The request timed out. The AI is taking longer than usual. Please try again."
  // server/network/http/unknown -> "Sorry, I couldn't get a response. ..."
  // (timeout strings copied VERBATIM from taskpane.js:2406-2408)
}
```

Intentional behavior changes (record them in the status note): auth errors are
no longer retried (previously retried 3×); 429 and network errors are now
retried (previously fatal); key moves from URL to header.

### Step 2 — migrate the four call sites (one sub-commit each is fine)

1. **Chat loop** (`taskpane.js`): delete `callGeminiWithRetry` (lines
   2431–2518). At the call site (:1676) use
   `callGemini({ model: geminiModel, apiKey: geminiApiKey, payload,
   timeoutMs: TIMEOUT_LIMITS.FETCH_TIMEOUT_MS, maxRetries: modelProfile.retries,
   cancelSignal: currentRequestController?.signal })`.
   Replace the string checks:
   - `taskpane.js:1681-1684` → `apiError instanceof GeminiClientError && apiError.kind === "history_invariant"`.
   - outer catch `taskpane.js:2386` (`'Request cancelled by user'`) →
     `error.kind === "cancelled"` (keep the old string check too, OR'd, since
     other code paths still throw that string).
   - outer catch timeout override `taskpane.js:2405-2409` →
     `error.kind === "timeout"`, message from
     `userMessageForError("timeout", { previewThrottleWarning: modelProfile.previewThrottleWarning })`.
   - Also delete the now-dead URL construction at `taskpane.js:1295` (the
     client builds it).
2. **Glance checks** (`taskpane.js:936-1004`, `runGlanceChecks`): swap the raw
   fetch for `callGemini({..., timeoutMs: 30000, maxRetries: 2})`. Keep its
   existing catch/UI behavior.
3. **`callGeminiForDiffs`** (`agentic-tools.js:275`): replace the fetch block
   with `callGemini({..., timeoutMs: DIFF_CALL_TIMEOUT_MS, maxRetries: 2})`.
   Keep everything after the response parse identical (candidates checks,
   `repairTruncatedJsonArray`, `return null` on error, the timeout
   console.error).
4. **`callGeminiForJSON`** (`agentic-tools.js:680`): replace the fetch with
   `callGemini({..., timeoutMs: 60000, maxRetries: 2})`. **Keep
   `temperature: 0.2` and `maxOutputTokens: 48000` literal in the payload** —
   the model profile's temperature is 0.1 and switching would change behavior;
   note it as a candidate follow-up instead. Add salvage: if the response text
   starts with `[`, parse via `repairTruncatedJsonArray` (already imported in
   this file); otherwise keep bare `JSON.parse` (object schemas can't be
   salvaged).

`gemini-client.js` gets `loadApiKey`/model via arguments — do NOT have it read
settings itself. In `agentic-tools.js` the values are already in scope from the
`initAgenticTools` deps.

### Step 3 — tests: `tests/gemini_client_tests.mjs`

Use an injected `fetchImpl` (an async function you control per test) and small
timeout values. Test functions:

1. `testClassifyAuth` — status 400 + body containing `API_KEY_INVALID` → kind
   `auth`, `retryable: false`; also 403 → auth.
2. `testClassifyHistoryInvariant` — each of the two literal strings → kind
   `history_invariant`, not retryable.
3. `testClassifyQuotaWithRetryAfter` — status 429, `Retry-After: 2` →
   `retryAfterMs: 2000`.
4. `testServerErrorRetriesThenSucceeds` — fetchImpl fails twice with 500 then
   succeeds; `maxRetries: 3` → resolves; attempt counter === 3.
5. `testAuthNotRetried` — 403 every time, `maxRetries: 3` → exactly 1 fetch
   call, throws kind `auth`.
6. `testQuotaHonorsRetryAfter` — 429 with `Retry-After: 0` then success →
   resolves after 2 calls.
7. `testTimeoutRetried` — fetchImpl hangs (never resolves, but listens to the
   AbortSignal and rejects with an AbortError when aborted), `timeoutMs: 30` →
   retried; exhausting retries throws kind `timeout`.
8. `testCancelNotRetried` — abort the `cancelSignal` mid-call → throws kind
   `cancelled`, 1 fetch call.
9. `testKeyInHeaderNotUrl` — capture fetch args: URL contains no `key=`;
   headers contain `x-goog-api-key`.
10. `testBackoffJitterBounds` — `backoffMs(2, 1000, () => 1)` === 4000;
    `backoffMs(2, 1000, () => 0)` === 0.

### Acceptance criteria

- `grep -rn "generativelanguage" src/` → only `gemini-client.js`.
- `grep -rn "function response turn" src/` → only `gemini-client.js`.
- `grep -rn "?key=" src/` (or `key=\${`) → no Gemini call sites.
- `npm test` green; `npm run build:dev` green.
- Manual smoke in Word: one chat turn with a tool call works; the Cancel button
  still cancels.

---

## WP3 — Reliability event log (observability)

**Goal:** when a defense layer fires, it leaves a record the user can copy-paste
into a bug report. "It's being fragile again" becomes "anchor correction fired
14 times and tier-2 recovery ran twice — here's the log."

**Why:** the Phase-1 defenses report only to the devtools console, which is
closed during normal use. Both Gemini-3.5-Flash pathologies were diagnosed only
because the user happened to capture console output. Meanwhile the console is
flooded unconditionally: the **entire chat history** is stringified every loop
iteration (`taskpane.js:1672`), plus full raw API responses (`taskpane.js:1753`,
`taskpane.js:1994`, `agentic-tools.js:320`, `agentic-tools.js:340`).

**Depends on:** WP2 (for the `onRetry` hook), but can be built standalone.

**Suggested commit message:** `WP3: add reliability event log + copy-diagnostics button; gate verbose logs`

### Step 1 — create `src/taskpane/modules/diagnostics/reliability-log.js`

```js
/*
 * reliability-log.js — in-memory record of defense-layer activations.
 * Pure/Node-importable. Detail fields are whitelisted per event type so
 * document text, prompts, or API keys can never enter the log.
 */

const MAX_EVENTS = 200;

export const EVENT_TYPES = Object.freeze({
  HISTORY_RECOVERY: "history_recovery",       // { tier }
  SANITIZER_REJECTION: "sanitizer_rejection", // { reason, operation, paragraphIndex }
  ANCHOR_CORRECTED: "anchor_corrected",       // { fromIndex, toIndex }
  ANCHOR_REJECTED: "anchor_rejected",         // { paragraphIndex }
  JSON_SALVAGE: "json_salvage",               // { salvagedCount }
  CORRECTIVE_RETRY: "corrective_retry",       // { outcome }  "applied"|"failed"
  LOOP_GUARD_TRIP: "loop_guard_trip",         // { loops }
  CHECKPOINT_FAILURE: "checkpoint_failure",   // { toolName }
  NATIVE_FALLBACK: "native_fallback",         // { mode }  e.g. "empty_target"|"append_at_end"|"literal_bullet_degrade"
  API_RETRY: "api_retry",                     // { kind, attempt }
  MALFORMED_CALL_RECOVERED: "malformed_call_recovered", // { toolName }
});

// Whitelist: field names allowed in `detail` per type. Extra fields are DROPPED.
const DETAIL_FIELDS = { /* mirror the comments above */ };

const events = [];
const counters = Object.create(null);

export function recordEvent(type, detail = {}) { /* push {ts: Date.now(), type, detail: filtered}; evict oldest past MAX_EVENTS; counters[type]++ */ }
export function getCounters() { /* shallow copy */ }
export function getEvents() { /* copy of events */ }
export function resetReliabilityLog() { /* for tests */ }

export function getDiagnosticsReport({ releaseMarker = "", model = "" } = {}) {
  // Plain text: header (release, model, date), one line per counter,
  // then the last 50 events as "ISO-time  type  {detail json}".
}

// Debug flag — localStorage-backed, read lazily so the module imports in Node.
const DEBUG_KEY = "debugLogging";
export function isDebugLogging(storage) { /* (storage || globalThis.localStorage)?.getItem(DEBUG_KEY) === "true"; wrap in try/catch, default false */ }
export function setDebugLogging(enabled, storage) { /* setItem / removeItem */ }
```

### Step 2 — instrument the defense layers

One `recordEvent` line at each site (import the module; all these files are
add-in-side so importing is safe):

| Event | File / anchor |
|---|---|
| `HISTORY_RECOVERY` | `taskpane.js` tier branches (~1690/1702/1709/1717) — record `{tier: currentRecoveryTier}` once where the tier escalates (~1687) |
| `SANITIZER_REJECTION` + `ANCHOR_*` | `agentic-tools.js` `executeRedline`, where sanitizer/anchor rejections are merged (search `formatRejections(`) — iterate the rejected array, one event each; record `ANCHOR_CORRECTED` in `applyRedlineChangeSet` where `correctedIndex` is applied |
| `JSON_SALVAGE` | `agentic-tools.js:347` (the `repaired` branch) |
| `CORRECTIVE_RETRY` | `agentic-tools.js` `executeRedline`, where `buildCorrectiveRetryPrompt` is invoked (search for it) — record outcome after the retry attempt resolves |
| `LOOP_GUARD_TRIP` | `taskpane.js:2339` (inside the `>= MAX_NO_PROGRESS_TOOL_LOOPS` branch) |
| `CHECKPOINT_FAILURE` | `taskpane.js` `createCheckpoint` catch (~1080) |
| `NATIVE_FALLBACK` | `word-redline-runner.js` — where `insertContentAsNativeParagraphs` is invoked for empty-target/append, and where OOXML block generation degrades to literal text (search `insertContentAsNativeParagraphs` and the literal-bullet fallback) |
| `API_RETRY` | pass `onRetry: (kind, attempt) => recordEvent(EVENT_TYPES.API_RETRY, {kind, attempt})` where `callGemini` is invoked (WP2) |
| `MALFORMED_CALL_RECOVERED` | `taskpane.js:1940` (the `if (recoveredFunctionCall)` branch) |

### Step 3 — gate the verbose dumps

Wrap each of these in `if (isDebugLogging()) { ... }`:
`taskpane.js:1672` (full chat history), `taskpane.js:1753` (raw chat result),
`taskpane.js:1994` (content.parts), `agentic-tools.js:320` (raw diff result),
`agentic-tools.js:340` (diff JSON text). Leave `console.warn`/`console.error`
lines ungated.

### Step 4 — settings UI

Existing button wiring lives in the `Office.onReady` block, e.g.
`document.getElementById("save-api-key").onclick = saveApiKey;`
(`taskpane.js:305`). Add to `src/taskpane/taskpane.html`'s settings view (find
the container holding `save-api-key`):

- `<button id="copy-diagnostics">Copy diagnostics</button>` → onclick:
  `navigator.clipboard.writeText(getDiagnosticsReport({ releaseMarker: RELEASE_MARKER, model: loadModel('fast') }))`,
  then briefly set the button text to "Copied!".
- `<input type="checkbox" id="debug-logging-toggle">` + label "Verbose
  logging" → initialize from `isDebugLogging()`, onchange calls
  `setDebugLogging(checked)`.

### Tests — `tests/reliability_log_tests.mjs`

1. `testRingBufferCap` — record 250 events → `getEvents().length === 200`,
   oldest evicted (check first event's detail).
2. `testCounters` — 3 of one type, 1 of another → counters correct.
3. `testDetailWhitelist` — `recordEvent(SANITIZER_REJECTION, {reason: "x", documentText: "SECRET"})`
   → stored detail has `reason` but no `documentText`; report string does not
   contain "SECRET".
4. `testReportContents` — report contains release marker, model, and a counter
   line.
5. `testDebugFlagDefaultOff` — with an injected fake storage (`{getItem: () => null}`)
   → false; with `"true"` → true; no storage at all → false, no throw.
6. `testReset` — `resetReliabilityLog()` clears events + counters.

### Acceptance criteria

- Forcing one sanitizer rejection in Word (instruction targeting "paragraph
  999") shows a `sanitizer_rejection` in the copied diagnostics.
- With verbose logging off, a normal chat turn logs no full-history or
  raw-response dumps (check devtools).
- `npm test` green.

---

## WP4 — Extract testable logic out of `taskpane.js`

**Goal:** the pieces of the agentic loop that keep breaking become plain modules
with unit tests; `taskpane.js` keeps only UI wiring and Word/Office glue.

**Why:** `sendChatMessage` is ~1,280 lines. The malformed-function-call
recovery parsers — regex recovery for hostile input, exactly the code that most
needs tests — have zero. The tool dispatch is ten near-identical `else if`
branches whose small differences (checkpoint or not, `showToUser` vs `success`)
have already drifted once and are invisible without a diff.

**Do the two extractions in order, each with its own commit. Do NOT attempt to
move the whole loop.**

**Suggested commit messages:**
`WP4a: extract malformed-function-call recovery into function-call-recovery.js`
`WP4b: extract tool dispatch table into tool-registry.js`

### Extraction A — `src/taskpane/modules/chat/function-call-recovery.js`

1. Move these, **verbatim**, from `taskpane.js` into the new module as named
   exports: `tryParseArgs` (:1784–1806), `parseMalformedEditListArgs`
   (:1808–1875), `parseMalformedConvertHeadersArgs` (:1877–1901).
2. Also extract the scanning loop (:1903–1938, including the legacy
   `apply_redlines` fallback regex) as:

   ```js
   /**
    * Try to recover a {name, args} functionCall from a malformed finishMessage.
    * @param {string} finishMessage
    * @param {string[]} toolNames - ordered list of known tool names
    * @returns {{name: string, args: object} | null}
    */
   export function recoverFunctionCall(finishMessage, toolNames) { ... }
   ```

3. In `taskpane.js`, replace the whole block with an import +
   `const recoveredFunctionCall = recoverFunctionCall(candidate.finishMessage, toolNames);`
   and keep the surrounding `parts = [{ functionCall: ... }]` logic (:1940–1949)
   where it is.
4. The new module must contain no DOM/Office references (the originals don't —
   verify).

**Tests — `tests/function_call_recovery_tests.mjs`:**

1. `testValidJsonArgs` — `tryParseArgs('{"instruction": "fix typos"}')` parses.
2. `testSingleQuotedPseudoJson` — `tryParseArgs("{instruction: 'fix', count: 2,}")`
   → `{instruction: "fix", count: 2}` (unquoted keys, single quotes, trailing
   comma — all three repairs at :1795–1801).
3. `testGarbageReturnsNull` / `testEmptyStringReturnsEmptyObject`.
4. `testEditListFieldRecovery` — a rawArgs string with
   `startParagraphIndex: 3, endParagraphIndex: 5, listType: bullet` and an
   unquoted `newItems: [First item. , Second item.]` → indexes parsed, listType
   normalized to `"bullet"`, two items split on the sentence boundary.
5. `testEditListMissingIndexesReturnsNull`.
6. `testConvertHeadersRecovery` — `paragraphIndices: [2, 4, 6], numberingFormat: decimal`
   → `{paragraphIndices: [2,4,6], numberingFormat: "decimal"}`.
7. `testRecoverFunctionCallFindsTool` — finishMessage
   `edit_list({startParagraphIndex: 1, endParagraphIndex: 2})` with the real
   tool-name list → `{name: "edit_list", args: {...}}`.
8. `testRecoverFunctionCallLegacyRedlineFallback` — finishMessage
   `apply_redlines {instruction: make it shorter}` → recovered with the
   instruction string.
9. `testRecoverFunctionCallNoMatch` — plain prose → null.

### Extraction B — `src/taskpane/modules/chat/tool-registry.js`

The ten dispatch branches (`taskpane.js:2049–2269`) collapse into a registry +
one generic loop body. **Normalize carefully — the branches differ in exactly
three ways, captured by the table below. Preserve each row's behavior exactly.**

| tool | mutating (checkpoint) | success flag | on success | on failure | track-entry `instruction` string |
|---|---|---|---|---|---|
| `apply_redlines` | yes | `result.showToUser` | update msg + checkpoint | `console.log` only | `args.instruction` |
| `insert_comment` | yes | `result.showToUser` | update msg + checkpoint | `console.log` only | `args.instruction` |
| `highlight_text` | yes | `result.showToUser` | update msg + checkpoint | `console.log` only | `args.instruction` |
| `perform_research` | no | always true | update msg `Found search results for: "${args.instruction}"` | — | `args.instruction` |
| `navigate_to_section` | no | always true | update msg `Navigated to: "${args.instruction}"` | — | `args.instruction` |
| `edit_list` | yes | `result.success` | update msg + checkpoint | update msg (no checkpoint) | `` `edit_list P${args.startParagraphIndex}-P${args.endParagraphIndex}` `` |
| `insert_list_item` | yes | `result.success` | update msg + checkpoint | update msg | `` `insert_list_item after P${args.afterParagraphIndex}` `` |
| `edit_table` | yes | `result.success` | update msg + checkpoint | update msg | `` `edit_table at P${args.paragraphIndex}: ${args.action}` `` |
| `edit_section` | yes | `result.success` | update msg + checkpoint | update msg | `` `edit_section at P${args.sectionHeaderIndex}` `` |
| `convert_headers_to_list` | yes | `result.success` | update msg + checkpoint | update msg | `` `convert_headers_to_list: ${args.paragraphIndices?.join(', ')}` `` |

Registry design (executors injected so the module stays Node-importable):

```js
/**
 * createToolRegistry(executors) -> { [toolName]: entry }
 * executors: { executeRedline, executeComment, executeHighlight,
 *              executeResearch, executeNavigate, executeEditList,
 *              executeInsertListItem, executeEditTable, executeEditSection,
 *              executeConvertHeadersToList }
 *
 * entry = {
 *   mutating: boolean,
 *   statusText: (args) => string,        // pre-run loading message; the current
 *                                        //   strings are at taskpane.js:2034-2041
 *                                        //   and per-branch updateSystemMessage calls
 *   describe: (args) => string,          // the track-entry instruction column above
 *   run: async (args, ctx) => ({         // ctx = { docText }
 *     text: string,                      // functionResponse text (old toolResult)
 *     succeeded: boolean,
 *     onSuccess: "update_with_checkpoint" | "update",
 *     onFailure: "update" | "console",
 *     completionText?: string,           // research/navigate post-run message
 *   }),
 * }
 */
export function createToolRegistry(executors) { ... }
```

In `taskpane.js`:

1. Build the registry once after imports:
   `const TOOL_REGISTRY = createToolRegistry({ executeRedline, executeComment, ... });`
2. Replace the `else if` chain with one generic body: look up
   `TOOL_REGISTRY[functionCall.name]`; if missing, set
   `toolResult = "TOOL_FAILURE unknown_tool: '" + functionCall.name + "' is not a known tool."`
   and `toolSucceeded = false` (today an unknown name silently produces an
   empty functionResponse). If mutating →
   `const checkpointIndex = await createCheckpoint(true, functionCall.name);`
   then run, then apply the on-success/on-failure display per the entry, push
   the track entry, etc.
3. Replace the `mutatingToolNames` set (:2013–2022) with a lookup on the
   registry, and the hardcoded `toolNames` array in the recovery path (:1771)
   with `Object.keys(TOOL_REGISTRY)`. After this, **exactly one list of tool
   names exists**.

**Tests — `tests/tool_registry_tests.mjs`** (use stub executors that record
their arguments and return canned results):

1. `testRegistryCoversAllTools` — keys are exactly the 10 names above.
2. `testMutatingFlags` — the 8/2 split matches the table.
3. `testShowToUserToolsMapSuccess` — stub `executeRedline` returning
   `{message: "m", showToUser: false}` → `succeeded === false`,
   `onFailure === "console"`.
4. `testSuccessShapeTools` — stub `executeEditList` returning
   `{message: "m", success: false}` → `succeeded === false`,
   `onFailure === "update"`.
5. `testDescribeStrings` — each entry's `describe(args)` output matches the
   table column for representative args.
6. `testExecutorArgWiring` — `edit_table` entry passes
   `(args.paragraphIndex, args.action, args.content, args.targetRow, args.targetColumn)`
   positionally (compare with `taskpane.js:2195-2201`); same spirit for
   `edit_list`, `insert_list_item`, `edit_section`, `convert_headers_to_list`,
   and `highlight_text`'s `args.color || "yellow"` default (:2092).

### Acceptance criteria

- `taskpane.js` shrinks by ≥300 lines; `npm run build:dev` green.
- All new tests pass; `npm test` green.
- Manual smoke in Word: one redline instruction and one `edit_list` instruction
  behave as before (status messages, checkpoint/undo button on success).

---

## WP5 — Record/replay evals + incident regression cases

**Goal:** model-behavior regressions are caught by replayable fixtures in
`npm test`, not by the user mid-edit. Every live incident becomes a permanent
eval case.

**Why:** the eval harness requires a paid `GEMINI_API_KEY` and manual
invocation, so it has never gated anything. Both gemini-3.5-flash pathologies
(thought leakage 2026-07-04, repetition loop 2026-07-05) were found in
production. The harness already exposes the needed seams: `buildAnchoredText`,
`scoreChangeSet`, and cases in `tests/evals/cases/*.json`.

**Also fixes a fidelity bug:** the harness's `callGemini`
(`tests/evals/run-evals.mjs:84-108`) uses bare `JSON.parse`, but production
(`callGeminiForDiffs`) parses via `repairTruncatedJsonArray`. Replay must use
the production parse so truncation fixtures work.

**Suggested commit message:** `WP5: add --record/--replay to eval harness + synthetic incident fixtures; wire replay into npm test`

### Step 1 — recording format and directory

`tests/evals/recordings/<model>/<case-name>.json`:

```json
{
  "model": "gemini-2.5-flash",
  "case": "rename-title",
  "recordedAt": "2026-07-05T00:00:00Z",
  "rawText": "<the raw JSON text part exactly as the model returned it>"
}
```

Store `rawText` (the pre-parse string), NOT the parsed array — so replay
exercises parsing/salvage too.

### Step 2 — modify `tests/evals/run-evals.mjs`

1. `parseArgs`: add `--record` and `--replay` booleans.
2. In `callGemini`, return the raw `text` (move the `JSON.parse` out to the
   caller). In the main loop:
   - live mode (default): `rawText = await callGemini(...)`; if `--record`,
     write the recording file (mkdir -p the model dir).
   - `--replay`: skip the network entirely; read the recording; if missing,
     mark the row FAIL with reason `no recording` — unless no recordings exist
     for that model at all, in which case skip the model with a notice.
   - Parse in both modes with the production function:
     `import { repairTruncatedJsonArray } from '../../src/taskpane/modules/commands/change-validation.js';`
     `const { changes } = repairTruncatedJsonArray(rawText);` and treat `null`
     changes as a scoring failure.
3. `--replay` with no `--model` args: replay every subdirectory of
   `recordings/`. Exit 0 if `recordings/` doesn't exist yet (print a notice) —
   this keeps `npm test` green before any recordings are committed.
4. Keep the no-key guard ONLY for live mode; `--replay` must not require
   `GEMINI_API_KEY`.

### Step 3 — synthetic incident fixtures

Create `tests/evals/recordings/_synthetic/` with two recordings tied to **new
cases**. First print the paragraph anchors of the existing fixture so the
`anchorText`/`paragraphIndex` values you write are real:

```bash
node -e "import('./tests/evals/run-evals.mjs').then(async m => console.log(m.buildAnchoredText('tests/evals/cases/<the docx used by case 01>')))"
```

(Read `tests/evals/cases/01-rename-title.json` for the `documentFixture` path.)

1. **Case `tests/evals/cases/06-thought-leakage.json`** — instruction like
   "improve the definitions section"; `expect`:
   `{"minChangesApplied": 0, "maxChangesApplied": 0, "mustNotContainText": ["I should", "the user wants"]}`.
   Recording `_synthetic/06-thought-leakage.json` `rawText`: a JSON array with
   one object shaped like the 2026-07-04 live payload — a valid operation and
   real `paragraphIndex`/`anchorText`, chain-of-thought prose in
   `replacementText`, and NO `content` field. Expected pipeline behavior: the
   sanitizer's field-stripping + `empty_content` guard rejects it → 0 applied.
2. **Case `tests/evals/cases/07-repetition-loop.json`** — any instruction;
   `expect`: `{"minChangesApplied": 1, "maxChangesApplied": 1}`.
   Recording `rawText`: one **valid** `replace_paragraph` object (real anchor,
   real index, short `content`) repeated ~40 times in a JSON array, with the
   final repetition cut off mid-string and no closing `]` — mimicking the
   2026-07-05 truncation. Expected: `repairTruncatedJsonArray` salvages the
   complete objects, `duplicate_target` dedupe collapses them → exactly 1
   applied.

### Step 4 — wire replay into the runner

In `tests/run-all.mjs`, append a synthetic entry after the discovered files:
spawn `node tests/evals/run-evals.mjs --replay` and report it as one more
PASS/FAIL row.

### Step 5 — document the cadence

Add to `README.md` (Development → eval harness section) and to this WP's
status note:

- **Before adding a model to the Settings dropdown:** run live evals for it
  (`--model <name> --record`), commit its recordings, and set/adjust its
  `model-profiles.js` entry based on the results. This is the gate that was
  missing when gemini-3.5-flash shipped.
- **After editing `redline-prompt.js` or `REDLINE_DIFF_SCHEMA`:** re-run live
  evals for the default fast + slow models with `--record`, commit updated
  recordings.

### Verification

1. `node tests/evals/run-evals.mjs --replay` → the two synthetic cases PASS,
   exit 0, no API key needed.
2. Prove the fixtures bite: temporarily make `repairTruncatedJsonArray` return
   `{changes: null}` → case 07 fails. Temporarily disable the sanitizer's
   field-stripping step → case 06 fails (it would apply the leaked reasoning).
   Revert both.
3. `npm test` includes the replay row and stays green.

### Acceptance criteria

- Replay runs offline in `npm test`; synthetic incident fixtures pass and are
  proven to fail when their defenses are disabled.
- README documents the record/replay workflow and the model-addition gate.

---

## WP6 — Word-bridge verification kit

**Goal:** the layer where real failures happen (Word.run / insertOoxml / tracked
changes) gets (a) a reusable mock so its logic paths are unit-tested, and (b) a
written, repeatable manual smoke procedure for what only real Word can verify.

**Why:** the bridge (`src/taskpane/modules/docx-redline-js-integration/`,
~2,500 lines; 122 `Word.run`/`context.sync` sites repo-wide) is where
valid-looking OOXML meets Word's actual acceptance rules — the
`InvalidArgument` class of failures Node tests structurally cannot catch. Every
past plan ends with "manual in-Word smoke test still recommended" and there is
no record one ever ran. A good mock already exists ad hoc inside
`tests/empty_paragraph_insertion_tests.mjs:22-85` — it just isn't shared.

**Suggested commit message:** `WP6: shared mock-word test helper + bridge fallback tests + SMOKE-TEST.md`

### Step 1 — create `tests/helpers/mock-word.mjs`

Move (verbatim, then generalize) from `tests/empty_paragraph_insertion_tests.mjs`:

- `makeMockParagraph(log, id)` (:22–34)
- `makeFormattingMockParagraph(log, id)` (:36–52)
- `makeOoxmlMockParagraph(log, id, opts)` (:56–76) — already supports
  `opts.failOoxml` (throws on `insertOoxml`); keep it.
- the two context objects (:78–85) → export factories instead:

```js
export function makeMockContext({ failSync = false } = {}) {
  return {
    sync: failSync
      ? async () => { throw Object.assign(new Error("GeneralException"), { name: "RichApi.Error", code: "GeneralException" }); }
      : async () => {},
    document: { load() {}, changeTrackingMode: "Off" },
  };
}
export function insertedTexts(log) { /* from :87-90 */ }
export function ensureWordGlobal() {
  if (!globalThis.Word) globalThis.Word = { ChangeTrackingMode: { off: "Off" } };
}
```

Add `opts.failOoxmlOnce` (throw on first call only) — needed for degrade-path
tests. Update `empty_paragraph_insertion_tests.mjs` to import from the helper
and delete its local copies; its tests must still pass unchanged.

### Step 2 — `tests/word_bridge_fallback_tests.mjs`

New suite using the helper, targeting the highest-risk currently-untested
behaviors in `word-redline-runner.js` (read the exports first:
`applyRedlineChangesToWordContext`, `insertContentAsNativeParagraphs`,
`prepareNativeMarkdownParagraphs`, `segmentNativeInsertionBlocks`):

1. `testAppendAtEndUsesNativeInsertion` — a change set with
   `startIndex === paragraphCount` → the mock log shows `insertParagraph`
   calls after the last paragraph, no `insertOoxml` on the body.
2. `testBulletBlockDegradesToLiteralText` — with `failOoxml: true`, a content
   block containing `- item` lines → log shows `insertText`/`insertParagraph`
   with `• item` text; assert NO raw `- ` or `**` markers in any inserted text.
3. `testSkippedChangesPropagate` — a change the engine skips (e.g. empty
   target where insertion is not possible) → the returned object's `skipped[]`
   contains `{paragraphIndex, operation, reason}`.
4. `testTrackingRestoredOnFailure` — using a mock whose `insertOoxml` throws:
   the document's `changeTrackingMode` value is restored to its pre-call value
   after the operation completes (search `word-redline-runner.js` /
   `word-route-change.js` for where tracking is toggled — if restoration on the
   throw path is NOT implemented, that is a real bug: fix it minimally in the
   same WP and note it).

Adjust the four cases to the actual function signatures you find — the intent
(fallbacks + cleanup under failure) is what matters, not the literal wording.

### Step 3 — `docs/SMOKE-TEST.md`

Create with this content (verbatim, then keep it updated):

```markdown
# Release Smoke Test (real Word)

Run before every release build (npm run build:prod). ~15 minutes.
Append a row to the Runs log at the bottom each time.

Setup: sideload the add-in (npm run start), open a copy of tests/Sample NDA.docx,
set your API key, ensure Redlines toggle is ON.

| # | Step | Expected |
|---|------|----------|
| 1 | Open task pane, open Settings | Release marker shown matches the build version |
| 2 | Chat: "In the first paragraph, change 'Agreement' to 'Contract'" | Tracked insert+delete appear; chat shows success + undo button |
| 3 | Review > Reject All Changes | Document byte-identical to original text |
| 4 | Chat: "Add a bulleted list of three obligations after the definitions section" | Real Word bullets (not literal • text), tracked |
| 5 | Chat: "Turn the parties' details into a two-column table" | Real Word table appears, tracked |
| 6 | Place cursor on an EMPTY paragraph; chat: "insert a short confidentiality clause here" | Native tracked insertion, no error |
| 7 | Chat: "Append a governing-law section at the end of the document" | New paragraphs after the last one, tracked; reject-all reverts fully |
| 8 | Click the undo button on an earlier edit message | Document restored to that checkpoint |
| 9 | Chat: "Edit paragraph 99 to say hello" (doc has fewer) | Informed failure message; document unchanged; no infinite retry |
| 10 | Disconnect network, send any chat | Clean error message; UI unlocks; reconnect works |
| 11 | Repeat steps 2-4 in Word Online | Same results (w:rPrChange quirks are known — note anomalies) |

## Runs

| Date | Build | Word version | Result | Notes |
|------|-------|--------------|--------|-------|
```

### Step 4 — run it once

Execute the checklist in real Word Desktop (and Online if available) and record
the first Runs row. This closes the standing "not verified in real Word" items
from Phase 1 (empty-paragraph insertion, append-at-end).

### Acceptance criteria

- `mock-word.mjs` is imported by ≥2 suites; `npm test` green.
- The four fallback tests pass, including under `failOoxml`/`failSync`.
- `docs/SMOKE-TEST.md` exists with one completed Runs row.

---

## WP7 — Package-boundary contract tests

**Goal:** upgrading `@ansonlai/docx-redline-js` (or editing the bridge) fails
fast in `npm test`, not at runtime inside Word.

**Why:** the engine lives outside this repo, so its API is a fault line: a
subpath-export change already broke a test for weeks
(`ERR_PACKAGE_PATH_NOT_EXPORTED`), caught only by accident. Existing suites
check fragments of the surface, not the full set of imports the consumers use.

**Suggested commit message:** `WP7: contract test for the docx-redline-js package surface`

### Step 1 — enumerate the real surface

Run:

```bash
grep -rn "from '@ansonlai/docx-redline-js'" src browser-demo mcp tests --include="*.js" --include="*.mjs"
```

Known import sites as of 2026-07-05 (re-run the grep; do not trust this list
blindly): `src/taskpane/taskpane.js` (`setPlatform`),
`src/taskpane/modules/utils/markdown-utils.js`,
`src/taskpane/modules/commands/redline-operation-converter.js`, the six
`docx-redline-js-integration/*.js` files, `browser-demo/demo.js`,
`mcp/docx-server/src/services/docx-redline-js-service.mjs`, and several
`tests/**/*.mjs` files. Collect every named import into one deduplicated list.

### Step 2 — create `tests/package_contract_tests.mjs`

```js
import './setup-xml-provider.mjs';
import assert from 'assert';
import * as pkg from '@ansonlai/docx-redline-js';

// Every symbol any consumer (add-in, browser demo, MCP server, tests) imports.
// Regenerate with the grep in docs/PLAN.md WP7 when adding imports.
const REQUIRED_EXPORTS = [
  'applyRedlineToOxml',
  'preprocessMarkdown',
  'setPlatform',
  // ... the full deduplicated list from step 1 ...
];

function testAllExportsPresent() {
  const missing = REQUIRED_EXPORTS.filter((name) => !(name in pkg));
  assert.deepStrictEqual(missing, [], `Missing package exports: ${missing.join(', ')}`);
}

async function testEngineSmoke() {
  // One real end-to-end call so a behavioral break in a patch release is
  // caught, not just a missing export. Reuse the pattern from
  // tests/addin/integration_tests.mjs (applyRedlineToOxml on a minimal
  // document XML) and assert result.oxml contains a w:ins element.
}
```

### Step 3 — document the upgrade procedure

Top-of-file comment in the test + a short README note:
bump the exact-pinned version in `package.json` (keep it exact — no `^`) →
`npm install` → `npm test` (contract + replay evals) → run
`docs/SMOKE-TEST.md` steps 2–7 → commit `package.json` + `package-lock.json`
together with the note of the engine version.

### Verification

- Temporarily add `'thisExportDoesNotExist'` to `REQUIRED_EXPORTS` → suite
  fails listing it. Remove.
- `npm test` green.

### Acceptance criteria

- Every symbol any consumer imports appears in `REQUIRED_EXPORTS` (re-run the
  grep and diff against the list).
- The suite runs in `npm test` and fails on a missing export or a broken
  `applyRedlineToOxml`.

---

## WP8 — Drift control (version, docs, lint ratchet)

**Goal:** the three quiet entropy sources — version markers, documentation, and
lint debt — get mechanical gates so they stop drifting.

**Why:**
- `package.json` says `0.0.1` while `taskpane.js:64` hardcodes
  `RELEASE_MARKER = "v2.1.0.3"` — two hand-maintained versions, already
  disagreeing.
- Docs drift is chronic: README once documented four phantom suites; `SPEC.md`
  still says the core engine is "currently in
  `src/taskpane/modules/reconciliation`" — a directory removed in February —
  and shows an import example from `./modules/reconciliation/standalone.js`.
  Phase 1 fixed the docs *once*, manually; they drifted again within weeks.
- Lint debt on the three legacy files means `npm run lint` can never gate CI,
  so new violations are indistinguishable from old ones.

**Suggested commit messages:**
`WP8a: single-source the release version via webpack DefinePlugin`
`WP8b: add docs freshness check`
`WP8c: format-only reformat of legacy files (no logic changes)`
`WP8d: enable lint in CI`

### Step 1 — version single-sourcing

1. `package.json`: set `"version": "2.1.4"` (or the next real release number).
2. `webpack.config.js`: add at the top `const webpack = require("webpack");`
   and in `plugins`:

   ```js
   new webpack.DefinePlugin({
     __RELEASE_MARKER__: JSON.stringify("v" + require("./package.json").version),
   }),
   ```

3. `taskpane.js:64`: replace the literal with

   ```js
   /* global __RELEASE_MARKER__ */
   const RELEASE_MARKER = typeof __RELEASE_MARKER__ !== "undefined" ? __RELEASE_MARKER__ : "dev";
   ```

   (append `__RELEASE_MARKER__` to the existing `/* global ... */` header at
   `taskpane.js:8` instead of adding a second comment).
4. Verify: `npm run build:dev`, then grep the built `dist/taskpane*.js` for
   `v2.1.4`. Releases become: `npm version patch` → build → smoke test.

### Step 2 — docs freshness check: `tests/docs_check_tests.mjs`

```js
// Scans the listed docs for backtick-quoted repo paths and documented test
// commands, asserting each resolves on disk.
const DOCS = ['README.md', 'ARCHITECTURE.md', 'SPEC.md', 'STATE.md',
              'ROADMAP.md', 'docs/PLAN.md', 'docs/SMOKE-TEST.md'];
const PATH_RE = /`((?:src|tests|docs|plans|browser-demo|mcp)\/[^`\s*]+?)`/g;
// Skip rules: a line containing "removed", "deleted", or "phantom" is exempt
// (docs legitimately reference dead paths when describing history).
// Also collect /node tests\/[^\s`]+\.mjs/g matches and assert those files exist.
// Trim trailing "/" before fs.existsSync checks; ignore paths containing "<".
```

Then **fix what it finds**. Known offenders to fix now:
- `SPEC.md`: the "currently in `src/taskpane/modules/reconciliation`" claims
  (two places) and the `./modules/reconciliation/standalone.js` import example —
  update to `@ansonlai/docx-redline-js`.
- Re-check `README.md` test commands against the real `tests/` tree.

Iterate until the suite passes; it then runs in `npm test` forever.

### Step 3 — format-only commit (the lint ratchet, part 1)

1. Confirm `npm test` is green first.
2. Run `npm run prettier` (repo script → `office-addin-lint prettier`, formats
   per `office-addin-prettier-config`).
3. Verify no logic changed: `npm test` green, `npm run build:dev` green, and
   `git diff --stat` touches only source files (no test fixtures corrupted).
4. Commit **only formatting** with message
   `WP8c: format-only reformat of legacy files (no logic changes)`.
5. Create `.git-blame-ignore-revs` containing that commit's full hash with a
   comment line, and run
   `git config blame.ignoreRevsFile .git-blame-ignore-revs` (also note it in
   README for other clones).

### Step 4 — lint gate (part 2)

1. `npm run lint`; fix the residual non-cosmetic errors (typically `no-undef`
   for `console`/browser globals — extend the `/* global ... */` headers; do
   not disable rules file-wide without a comment saying why).
2. Add `- run: npm run lint` to `.github/workflows/ci.yml` after the test step.

### Acceptance criteria

- Settings view shows a release marker equal to `package.json` version in a
  production build, with zero manual sync.
- `docs_check` runs in `npm test`; deleting a referenced file turns it red;
  `SPEC.md` no longer claims the engine lives in-repo.
- `npm run lint` passes repo-wide and runs in CI.

---

## Out of scope for this plan

- **TypeScript migration.** Considered and deliberately deferred: converting a
  4,200-line Office.js-coupled surface is rewrite-scale risk, and WP2/WP4/WP7
  capture most of the type-safety value (typed error taxonomy, tested module
  boundaries, contract tests). Revisit after WP4 shrinks `taskpane.js`.
- Prompt rewrites, tool-schema changes, or new tools.
- `@ansonlai/docx-redline-js` internals (file issues upstream; this repo only
  gains contract tests).
- UI redesign, streaming responses.
- Fixing Word Online's `w:rPrChange` behavior (tracked in `STATE.md` blockers;
  WP6's smoke checklist makes its status visible per release, which is all this
  plan owes it).

## Overall acceptance

1. `npm test` exists, runs every suite (unit + replay evals + contract + docs
   check), and CI runs it plus the webpack build and lint on every push.
2. All Gemini traffic flows through one client with a tested retry/backoff
   policy; 429s and transient 5xx no longer surface as tool failures.
3. A reliability event log exists, is user-copyable, and every Phase-1 defense
   layer records into it; verbose console dumps are opt-in.
4. `docs/SMOKE-TEST.md` has at least one recorded full pass, closing all
   standing "not verified in real Word" items.
5. Both 2026-07 live incidents exist as replayable eval fixtures that fail when
   their defenses are disabled.
6. Version, docs, and lint each have a mechanical gate; none can silently drift.
