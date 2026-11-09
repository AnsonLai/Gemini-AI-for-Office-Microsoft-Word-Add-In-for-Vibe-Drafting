# Review fixes and follow-ups

**Created:** 2026-11-01, after a code review of the add-in on top of `d410b60` (release marker v2.3.0.1).
**Status:** Part A is committed as `110070a` (2026-11-07, at the owner's request, before the Part B checks in Word). It was re-reviewed on 2026-11-02, which found and fixed four problems, now folded into the sections below: the Revert check would have warned almost every time because of Word's background proofing; "Revert anyway" never expired and could discard changes made after the warning; a Glance card said "Not checked yet" after a check had run; and a key with characters a header can't carry failed as a network error. Part B is still to do. Part C is the backlog of remaining review findings, none started.

The review found the core editing pipeline careful and well tested. The problems sat around it: the older UI code in `taskpane.js`, the Revert button, the automatic Glance checks, the API key handling, and the project tooling. This plan records what has been fixed, how to verify it, and what is left.

## Part A: done (committed as `110070a`)

The working tree changes 18 tracked files (+599/−129 lines) and adds 5 new files. `npm test`: 51 suites pass, 0 fail. `npm run build` succeeds.

### A1. Revert can no longer silently wipe later work

**Problem:** every edit message had a "Revert changes" button that never expired. Clicking an old one replaced the whole body with that snapshot, without asking and without checking whether the document had changed since. Later AI edits and the user's own typing were lost.

**Fix:**

- **A token that ignores Word's own background changes.** After a write, Word's background proofing adds `w:proofErr` marks and splits runs at them (the golden checks already strip these for the same reason). Layout adds `w:lastRenderedPageBreak`, saving can add a `_GoBack` bookmark, and settings get a `w:proofState`. The freshness token counts all of these, so an untouched document would look changed. The new `computeWordRevertGuardToken` in [`word-source-identity.js`](../../src/taskpane/modules/docx-redline-js-integration/word-source-identity.js) starts from the freshness token's normalization, drops those marks, paragraph ids and empty runs, and merges adjacent runs with identical properties. Text, formatting, revisions, comments and real bookmarks still count. It is used only by the Revert check; the edit freshness check is unchanged.
- **Token after each write.** After a verified write, [`document-tools.js`](../../src/taskpane/modules/chat/document-tools.js) computes that token (with the host paragraph count) and reports it through a new `onVerifiedWrite({ checkpointIndex, revisionToken })` option. `taskpane.js` saves it on that edit's checkpoint as `afterToken` with the new `updateCheckpoint(id, fields)` in [`checkpoint-store.js`](../../src/taskpane/modules/storage/checkpoint-store.js). A failure in this step never changes the edit's outcome. In a turn with several edits, the message's Revert button belongs to the last edit, which is also the checkpoint that gets the token.
- **Guarded restore.** [`checkpoint-restoration.js`](../../src/taskpane/modules/chat/checkpoint-restoration.js) `restoreBodyCheckpoint` takes a `guardToken`. When one is passed, it reads the current body and paragraph count inside the same `Word.run` and compares tokens before writing. If they differ, or there is no stored token (`null`, for example checkpoints from before this change), it writes nothing and returns `changedSinceEdit: true`, the current state's `currentToken`, and `CHANGED_SINCE_EDIT_MESSAGE`. Leaving `guardToken` undefined skips the check (only the older sync-count tests do this).
- **Confirm in chat.** `window.confirm` is blocked in Office add-ins, so `restoreCheckpoint` in `taskpane.js` shows the warning in the chat message with two buttons from the new `addConfirmButtons` in [`chat-ui.js`](../../src/taskpane/modules/chat/chat-ui.js): "Revert anyway" and "Keep document" ("Revert cancelled. The document was not changed."). "Revert anyway" passes the `currentToken` from the warning as the new guard, so it discards exactly the state the user was shown. If the document changed again before the click (the buttons don't expire), it asks again instead of discarding the newer changes.
- **No revert during a request.** Revert is refused with "Wait for the current request to finish, or cancel it, before reverting." while a chat request runs. It used to abort the request, which could collide with a write in progress.
- **Missing checkpoint:** shows "This checkpoint is no longer available."

### A2. Checkpoints belong to the pane that made them

**Problem:** one pool of 10 checkpoints shared by every document, with no owner recorded. Working in a second document evicted the first one's checkpoints, and full copies of confidential documents stayed on disk indefinitely.

**Fix ([`checkpoint-store.js`](../../src/taskpane/modules/storage/checkpoint-store.js)):**

- `makeCheckpointRecord` and `saveCheckpoint(label, ooxml, { sessionId })` tag each record with the taskpane session (`documentInstanceId`).
- New pure helper `idsToPrune` keeps this session's newest 10 and removes other or untagged sessions' checkpoints older than `OTHER_SESSION_MAX_AGE_MS` (24 hours) or with no timestamp.
- New `pruneCheckpoints(sessionId)` runs at startup from `ensureCheckpointMigration`, and after every save.
- Without a `sessionId`, `saveCheckpoint` keeps the old global cap of 10.

### A3. Glance is opt-in, cancellable, size-limited and parsed robustly

**Problem:** Glance sent the whole document to Gemini every time the pane opened, without being asked, with Google Search on and safety filters off. It had no size limit, couldn't be cancelled (it borrowed a usually empty controller), and read only the first part of the reply.

**Fix:**

- **Opt-in.** New setting "Run Glance checks when the pane opens" (`glance-autorun-toggle` in [`taskpane.html`](../../src/taskpane/taskpane.html), stored as `glanceAutoRun`). Off by default, with the hint "Each check sends the full document text to Gemini…". When off, each card says "Press refresh to run this check." Saving the API key re-runs Glance only when auto-run is on.
- **Own controller.** `runGlanceChecks` has its own `AbortController`. A new run aborts the previous one, and every UI update checks it is still the current run, so an old answer can't overwrite a new one.
- **Size limit.** Documents over `DOCUMENT_LIMITS.MAX_WORDS` (30,000 words) are refused: "This document is too long for Glance checks."
- **Parsing** moved to the new [`glance-results.js`](../../src/taskpane/modules/glance/glance-results.js): `glanceResponseText` joins all non-thought text parts; `parseGlanceResults` accepts only green, yellow or red (otherwise gray), caps summaries at 200 characters, and leaves missing ids out ("No answer for this check.").
- **Behaviour change:** users who relied on automatic checks must turn the setting on.

### A4. Settings no longer break on quotes or `</textarea>`

**Problem:** the Glance settings and cards were built by pasting raw text into HTML. A title containing `"` or `</textarea>` broke the form. The Content Security Policy blocked script injection, so this was a bug rather than a security hole.

**Fix:** `renderGlanceSettings` uses fixed markup and sets values with `.value`. `renderGlanceMain` builds cards with `textContent`.

### A5. API key travels only in the `x-goog-api-key` header

**Problem:** the key was in the request URL (`?key=…`), where it can end up in URL logs, proxies and error reports.

**Fix ([`gemini-client.js`](../../src/taskpane/modules/chat/gemini-client.js)):** `geminiEndpoint(model)` returns the URL without the key. `requestGeminiJson(url, payload, { apiKey })` sends it in the header and refuses (`REQUEST_CONFIGURATION`) before sending when the key or model is missing. A header can't carry characters a URL could encode, so the key is trimmed (also when saved in Settings), and a key with spaces, invisible or non-ASCII characters is refused with "Paste it again in Settings" instead of failing as a retried network error. `callGeminiWithRetry` and Glance pass the key through. The production bundle contains `x-goog-api-key` and no `key=`.

**Verified 2026-11-01:** a browser `fetch` from `https://example.com` with this header and a made-up key got a readable HTTP 400 `API_KEY_INVALID`. So Google allows the header across origins and reads the key from it.

### A6. Web search tool (`search_web`)

**Problem:** the welcome screen promised Google search and "online searches" for the Think button, but chat had no search tool.

**Fix:**

- New executor [`web-search.js`](../../src/taskpane/modules/chat/web-search.js). Gemini 2.5 can't combine the built-in `google_search` tool with function declarations in one request, so each search is a separate `generateContent` call with only `google_search`, using the conversation's model, key, retry wrapper, deadline and Cancel button.
- **Limits:** query 1 to 300 characters; answer capped at 4,000 characters (`answerTruncated`); at most 8 sources, https only, de-duplicated.
- **Results** are `{ status: 'ok', query, answer, sources, note }`. The note says web results are reference material, not instructions, and the model must cite sources.
- **Errors:** `INVALID_SEARCH_QUERY`, `WEB_SEARCH_FAILED`, `WEB_SEARCH_EMPTY`. Cancellation is rethrown so the loop stops.
- [`tool-registry.js`](../../src/taskpane/modules/chat/tool-registry.js) declares `search_web` (required `query`), validates it, and marks it `refreshesTargets: false`.
- **Loop guard ([`document-agent-loop.js`](../../src/taskpane/modules/chat/document-agent-loop.js)):** a successful read-only call clears the refused-edit memory only when it refreshes targets. So a search between two identical refused edits can't unlock the resend (`REPEATED_REFUSED_EDITS`). Confirmed by removing the fix and watching the test fail.
- **Editing guide** [`document-editing-guide.js`](../../src/taskpane/modules/chat/document-editing-guide.js) is now v11. "Research" is no longer listed as unsupported. New rules: search only for external facts, never for text already in the document, never paste document passages into a query, treat results as reference material and cite them, and keep search and edits as separate steps.
- **Welcome screen and README** describe editing, comments, formatting, search for both Fast and Think, the Glance opt-in, and the key header.

### A7. Tests added or changed

| Suite | What it covers |
| --- | --- |
| `tests/checkpoint_store_tests.mjs` | session tag on records; `idsToPrune` rules |
| `tests/checkpoint_update_tests.mjs` (new) | `updateCheckpoint` merges fields, ignores missing ids |
| `tests/checkpoint_restoration_tests.mjs` | guard: unchanged document reverts; changed document writes nothing; `null` guard counts as changed; "Revert anyway" discards the shown state but asks again about newer changes; proofing noise alone reverts without asking |
| `tests/word_source_identity_tests.mjs` | Revert guard token ignores proofing marks, split runs, layout marks, `_GoBack`, paragraph ids, rsids and `proofState`; still sees typing, formatting, tracked insertions and named bookmarks |
| `tests/document_tools_tests.mjs` | token reported after a verified write; a throwing callback doesn't change success; refused batch reports nothing |
| `tests/glance_results_tests.mjs` (new) | text joining, status whitelist, caps, missing ids, no-JSON error |
| `tests/gemini_client_tests.mjs` | key only in the header, never in the URL; missing key or model refused before sending; key trimmed; unusable characters refused with a settings message |
| `tests/tool_registry_tests.mjs` | `search_web` declared and validated |
| `tests/web_search_tests.mjs` (new) | grounded payload, sources filter, truncation, errors, cancellation, loop guard, failed search doesn't override the reply |

### A8. Side effects of Part A

- **Bundle size:** `taskpane.js` grew from 248,473 to 255,926 bytes and is now over webpack's advisory 244 KiB limit. Mostly IE11 transpilation (see C6).
- **Lint:** non-formatting errors went from 67 to 70. All are the existing false-positive kind (missing browser globals, Office load objects). See C3.
- **New risk:** text from web pages can now reach the model. The labelling and guide rules reduce the chance it acts on instructions hidden in a page, but don't remove it. Any edit still goes through the same validation and appears as a tracked change.

## Part B: verify in Word before committing

None of Part A has run inside Word. Do these by hand in Word desktop with the dev build (`npm run dev-server`, sideloaded manifest).

1. **Revert, unchanged document:** ask for an edit, wait a minute so Word's spelling and grammar check runs over the new text (use a deliberate misspelling to be sure), save the document, then click its Revert. Expect the document restored with no prompt. **This is the most important check:** if Word adds background changes the Revert token doesn't ignore, the warning appears here, and the token needs another normalization.
2. **Revert, changed document:** ask for an edit, type a word yourself, then click Revert. Expect the warning with "Revert anyway" and "Keep document". "Keep document" leaves your word. "Revert anyway" restores the snapshot. Repeat, but type another word between the warning and "Revert anyway": expect a second warning.
3. **Revert, two edits:** make two edits, then revert the first. Expect the warning (the second edit changed the document).
4. **Revert during a request:** start a long request and click an older Revert. Expect the "Wait for the current request…" message and no write.
5. **Old checkpoints:** a checkpoint saved before this change (no `afterToken`) must ask before reverting.
6. **Glance off by default:** open the pane. Expect no request (check the network tab) and the "Press refresh" text. Press refresh: the checks run.
7. **Glance on:** turn on the setting, reopen the pane. Checks run on open. Press refresh twice quickly: only the second run's results show.
8. **Glance settings:** a title containing `"quoted" </textarea>` saves and reloads intact.
9. **Web search:** ask "What is the current UK VAT rate? Add it to the first paragraph." Expect a "Searching the web..." step, a cited answer, then a separate tracked edit.
10. **Cancel during search:** press Cancel while "Searching the web..." shows. The request stops cleanly.
11. **Network check:** in the dev tools network tab, no request URL contains `key=`.

Also run checks M1 to M9 from [the state machine fixes plan](2026-11-07-state-machine-fixes.md#manual-checks-in-word), which cover the Tier 1 and Tier 2 fixes made on 2026-11-07.

**Exit condition:** all eleven behave as described. Part A is already committed, so any failure is fixed in a follow-up commit.

## Part C: remaining review findings (not started)

Ordered by impact. Each item says where, why, what to do, and when it's done.

### C1. Glance still sends the full document with Google Search on and safety filters off

- **Where:** `runGlanceChecks` in [`taskpane.js:888`](../../src/taskpane/taskpane.js) sends `tools: [{ google_search: {} }]` and `SAFETY_SETTINGS_BLOCK_NONE`.
- **Why:** Glance questions are about the document itself, so search adds little. With search on, the model can turn document passages into public search queries, which is the leak `search_web` was written to prevent. Safety filters off is not needed for review questions.
- **Do:** remove `google_search` from the Glance payload, or make it a per-check option that is off by default. Decide whether `BLOCK_NONE` is still needed (legal text with violence or medical terms can trip the default filters, so test before changing).
- **Done when:** the Glance payload has no `tools` by default, and a test pins it.

### C2. Chat's 30,000-word limit is backwards

- **Where:** [`taskpane.js:1110`](../../src/taskpane/taskpane.js), `DOCUMENT_LIMITS.MAX_WORDS` at line 148.
- **Why:** chat no longer sends the document text. It sends a paragraph count and headings, and the agent reads only what it targets. Long contracts are refused for no reason. Glance does send everything, and now has its own check (A3).
- **Do:** remove the chat check, or replace it with a limit based on what is actually sent (for example, inspection time or the number of paragraphs). Keep the Glance check.
- **Done when:** a 50,000-word document can be edited in chat, and Glance still refuses it.

### C3. Lint is broken and has never been enforced

- **Where:** `.eslintrc.json`, `package.json` scripts `lint` and `lint:fix` (`office-addin-lint`).
- **Why:** plain `eslint` won't run (ESLint 9 can't read `.eslintrc.json`). `npm run lint` reports about 1,620 errors: about 1,550 are Prettier formatting, and 70 are false positives such as `console` or browser globals flagged as undefined. Real problems are buried.
- **Do:**
  1. Move to a flat `eslint.config.js` with browser and ES2022 globals and `Office`/`Word` as readonly globals.
  2. Decide on Prettier: either run `lint:fix` once in a formatting-only commit, or drop the Prettier rule.
  3. Add `npm run lint` to the test runner or CI once it is clean.
- **Done when:** `npx eslint src tests` exits 0 and runs in CI.

### C4. Whole-body rewrites aren't checked for non-text content

- **Where:** every edit replaces the entire body ([`word-operation-runner.js:131`](../../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js)). The after-write check in [`document-tools.js:85`](../../src/taskpane/modules/chat/document-tools.js) compares text, structure, review marks and comments.
- **Why:** images, fields, bookmarks and content controls aren't compared, so losing one would go unnoticed. A whole-body replace also affects Word's undo stack, cursor position and co-authoring.
- **Do:**
  1. Add a Word host-lane case that edits one paragraph in a document with an image, a field, a bookmark and a content control, and checks each survives Accept All and Reject All.
  2. Extend the after-write check to count these parts (drawings, `w:fldSimple`/`w:fldChar`, `w:bookmarkStart`, `w:sdt`) before and after.
  3. Longer term, insert only the changed paragraphs' range instead of the whole body.
- **Library rule:** if the library drops these parts, report it upstream with a reproducer; don't patch it in the add-in (see KNOWN_BUGS.md).
- **Done when:** the host lane passes, and a lost image or field fails verification.

### C5. Settings save inconsistently (done 2026-11-09)

- **Where:** `showSettingsView` and `saveApiKey` in [`taskpane.js:387`](../../src/taskpane/taskpane.js) and [`:473`](../../src/taskpane/taskpane.js).
- **Why:** "Back" throws away unsaved model and system-prompt changes, while Glance and redline settings save on every change. No setting can be saved without an API key.
- **Do:** pick one model, either save-on-change everywhere or one Save button for all fields with an unsaved-changes notice on Back. Let model and prompt settings save without a key.
- **Done when:** every field behaves the same way, and a test or manual check covers Back with unsaved changes.
- **Done 2026-11-09:** every field saves on change, an empty API key removes the stored key, and the Save button is gone. Covered by `tests/settings_save_tests.mjs` and manual checks M13 and M14 in [the state machine fixes plan](2026-11-07-state-machine-fixes.md#manual-checks-in-word), which have not been run.

### C6. Drop IE11 support, production source maps and the double CSS load

- **Where:** `package.json` `browserslist` (`"ie 11"`), [`webpack.config.js`](../../webpack.config.js) (`devtool: "source-map"`, the `polyfill` entry and chunks), and `taskpane.css` loaded by both `taskpane.html:20` and `import "./taskpane.css"` in `taskpane.js:15`.
- **Why:** the code already needs Word's modern browser (`crypto.randomUUID`), so IE11 can't run it anyway. The build ships a 224 KB polyfill, and transpiling for IE11 is most of why the bundle is over 244 KiB (A8). Production source maps publish the full source. The CSS loads twice.
- **Do:**
  1. Set `browserslist` to Edge/Chromium WebView2 and Safari WKWebView (Mac Word), and remove the polyfill entry.
  2. Use `devtool: false` (or `hidden-source-map`) in production.
  3. Load the CSS one way only.
- **Done when:** the production bundle is under 244 KiB, has no `.map` files, and runs in Word on Windows and Mac.

### C7. Pin the default model

- **Where:** [`taskpane.js:518`](../../src/taskpane/taskpane.js) uses `gemini-flash-latest` for Fast (Think uses `gemini-2.5-pro`).
- **Why:** the editing guide and golden scenarios were tuned against one model, and Google can change what the alias points to without notice.
- **Do:** default to a dated model id, record it next to the golden-scenario results, and change it deliberately with a golden run.
- **Done when:** the default is a fixed id, and the README says how to change it.

### C8. Remove dead code and unused packages (done 2026-11-09, committed as `1b03198`)

- **Where and what:**
  - `removeAllFunctionPairs` and `createFreshStartWithContext` in [`chat-history.js:332`](../../src/taskpane/modules/chat/chat-history.js) are exported but never called.
  - `getLastCheckpoint`, `popLastCheckpoint`, `listCheckpoints` and `clearCheckpoints` in [`checkpoint-store.js`](../../src/taskpane/modules/storage/checkpoint-store.js) are never called outside the module.
  - `diff-match-patch` and `caniuse-lite` are runtime `dependencies` but unused in `src/`.
  - The leftover comment "forum discussion you referenced" at [`chat-history.js:105`](../../src/taskpane/modules/chat/chat-history.js).
- **Do:** delete them (check tests first, since some may import them), move `caniuse-lite` to dev dependencies or drop it, and rewrite the comment.
- **Done when:** `npm test` and the build pass, and grep finds no callers.
- **Done 2026-11-09 (`1b03198`):** the six functions (`removeAllFunctionPairs`, `createFreshStartWithContext`, `getLastCheckpoint`, `popLastCheckpoint`, `listCheckpoints`, `clearCheckpoints`) were deleted. `diff-match-patch` and `caniuse-lite` were removed from `package.json` dependencies; both remain installed transitively. The comment was rewritten. `npm test` passes all 52 suites and the build succeeds.

### C9. Make version numbers agree

- **Where:** `package.json` is still the Microsoft template (`office-addin-taskpane-js` 0.0.1, with Microsoft's repository URL); `manifest.xml` says `1.2.0.0`; the release marker in `taskpane.js:61` says `v2.3.0.1`.
- **Why:** Office uses the manifest version to decide whether to update the add-in, so it should move with releases.
- **Do:** pick one version (2.3.0.1), set it in all three, fix the package name and repository URL, and add a test that fails when they disagree.
- **Done when:** the test passes and the next release bumps all three together.

### C10. Docs and tests clean-up

- **Root docs (done 2026-11-09):** the root now holds `README.md`, `ROADMAP.md` and `STATE.md`; the historical docs are in `docs/archive/`, and `ARCHITECTURE.md` is at `docs/ARCHITECTURE.md`. Original note: seven markdown files sat at the root (`ARCHITECTURE.md`, `PLAN.md`, `README.md`, `ROADMAP.md`, `SPEC.md`, `STATE.md`, `SUMMARY.md`). `PLAN.md` and `SUMMARY.md` are historical and unmaintained. Move them to `docs/` or an archive.
- **Tests named after history:** suites such as `phase2_dead_code_static_tests` and `docx_redline_v040_compat_tests` guard past phases or old library versions rather than behaviour. Rename the ones that still test behaviour, delete the rest.
- **Proportion:** about 17,600 lines of tests and scripts for about 4,500 lines of product code. Look for overlapping suites as part of the renaming.
- **2026-11-09:** a suite-by-suite audit is in [the library test hand-off](2026-11-09-library-test-handoff.md): about 2,100 lines test only the library and are to move upstream or be deleted once covered there, and it lists the renames. The broken `tests/word-desktop` list lane and `tests/verify_thought_parsing.js` were deleted.
- **Editing guide:** it has grown one model quirk at a time into a dense list of rules. Regroup it by task (reading, editing, formatting, search) and re-run the golden scenarios after.
- **Done when:** the root holds README plus at most two current docs, no suite is named after a phase or version, and the golden lanes pass on the regrouped guide.

## Suggested order

1. Part B (verify in Word). Part A is committed (`110070a`).
2. C1 and C2: small, privacy and usability.
3. C3 (lint) in its own formatting-only commit, so later diffs are readable.
4. C6, C8, C9: build and housekeeping, each small.
5. C5 and C7.
6. C4: needs Word host-lane work.
7. C10: ongoing.
