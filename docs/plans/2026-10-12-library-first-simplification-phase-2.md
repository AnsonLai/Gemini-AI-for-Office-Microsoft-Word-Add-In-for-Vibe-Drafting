# Library-first simplification, phase 2

**Date:** 2026-10-12 (America/Vancouver)
**Status:** WP0 complete (2026-10-12); WP5 complete for the exposed list actions (2026-10-14); WP4 complete except the host paragraph count (2026-10-16). WP3 complete and merged 2026-10-17 on `docx-redline-js` 0.10.0. WP1 is re-scoped: the Word host layer stays in the plugin (owner decision, 2026-10-17); 4 deep imports wait on a library decision. WP2 complete as plugin-only cleanup (2026-10-17). The browser demo and MCP server were removed as obsolete spikes on 2026-10-17, and the list follow-up plan is complete and archived.
**Baseline:**
- Exact `@ansonlai/docx-redline-js@0.9.0` in the add-in and MCP.
- HEAD `a98f339` plus the uncommitted 0.9.0 upgrade work.
- **4,437 production lines across 20 files** (`node scripts/count-document-agent-code.mjs`). For comparison: 8,632 lines at the WP0 baseline of the [phase-1 plan](completed/2026-10-02-library-first-lightweight-agent-reliability.md), and 4,402 at its WP5 close.

**Related:**
- [Known bugs and the Word JS API inventory](KNOWN_BUGS.md)
- [Canonical list migration follow-up](completed/2026-09-30-canonical-list-migration-follow-up.md)
- [0.9.0 validation](../validation-reports/2026-10-12-docx-redline-v090.md)
- [Harness (optional)](2026-10-12-harness-optional.md): a parked Pi-style harness idea, split out of this plan on 2026-10-12

## Goal

Before `docx-redline-js` existed, this plugin carried its own OOXML engine, native Word editing routes and many interpreters, and became unwieldy. Phase 1 cut the production surface roughly in half and left two tools: inspect and apply. Phase 2 continues in the same direction.

**North star:** the plugin is three thin things:

1. **Taskpane UI and settings.**
2. **A small agent loop.** This is the current Gemini loop. Keep it small and give it a harness-like shape when it is touched for other reasons: named lifecycle events, schema-defined tools, and the safety policy in one module. Adopting a third-party harness such as Pi is a separate, parked option: see [Harness (optional)](2026-10-12-harness-optional.md), which recommends against it for now.
3. **A Word host adapter.** It reads the body package, writes one package, manages tracking state, and confirms the result.

Everything that *understands or transforms a document* lives in `docx-redline-js`. What remains in the plugin is the UI, the loop and the Word boundary. A new document capability should arrive as a library release plus a thin intent mapping, never as plugin logic.

## Rules for every package

- **Upstream first.** Propose the API in the library repository, release it, update both exact pins, and delete the consumer code in the same consumer commit. No consumer-side OOXML workarounds and no patched `node_modules` (see the [ownership rule](../../.agent/rules/docx-redline-js-ownership.md)).
- **Net reduction or no merge.** Each package reports `count-document-agent-code.mjs` before and after. A smaller plugin paired with a larger replacement module in the plugin does not count.
- **Public entry points only.** The consumer imports the package root, or documented subpaths the library commits to keeping stable. Today `src/` uses six internal subpaths (`services/package-parts.js`, `services/package-builder.js`, `services/comment-thread-parts.js`, `document/zip-archive.js`, `adapters/xml-adapter.js`, `standalone-runner`). Each one is coupling to library internals that a release can break.
- **Same gates, unchanged expectations.** For every package that touches editing or transport, run:
  - offline suites;
  - production and development builds;
  - the 14-case actual Office.js manifest;
  - the independent Word oracle (`Revisions.AcceptAll/RejectAll`);
  - the 16-prompt NDA replay.

  A refusal never stands in for a capability pass.

## Where document logic still lives in the plugin

| Module | Lines | What it does that is document logic | Library home (proposed, to negotiate) |
| --- | ---: | --- | --- |
| `chat/document-inspection-session.js` | 551 | Handle store, focused inspection, search, pagination, staleness and expiry checks | Stays in the plugin (WP2 re-scoped 2026-10-17); its target matching moves onto public library resolution. Originally proposed: a published, browser-safe **agent session** API. The library's `examples/agent-session-wrapper.mjs` shows the pattern, but it is development-only and uses private imports |
| `docx-redline-js-integration/consumer-core.js` | 349 | Flat OPC part and relationship handling, comment-part merge, insertion package assembly, internal subpath imports | A **Word Flat OPC body** surface: inspect, apply a batch, and return the insertion package with comment and numbering parts merged |
| `docx-redline-js-integration/word-source-identity.js` | 186 | Word rsid and paraId export-noise normalisation, export-tail detection | The same Flat OPC surface, as Word-export normalisation |
| `chat/document-tools.js` | 184 | Foreign-revision policy choice (`slice-cross-author`), `TRACKING_POLICY_MISMATCH` guard (KB-5), review-segment merge for verification | A library **auto revision policy** and a **verify-committed-against-prepared** helper |
| `chat/document-edit-intents.js` | 306 | Intent-to-operation mapping and scope refusals | Stays in the plugin, but should shrink once the library rejects unsupported shapes with its own codes |
| `word-operation-runner.js`, `word-ooxml.js`, `checkpoint-restoration.js`, `word-document-source.js` | 302 | Host transport and tracking state | Stays (necessary host boundary; see the KNOWN_BUGS inventory), minus expedient uses |
| `taskpane.js` | 1,249 | UI, settings, glance checks, checkpoint wiring | Stays. Only the glance `body.text` read is document logic |

If WP1–WP3 land as proposed, roughly 1,000–1,200 consumer lines become library calls. The size target is an owner decision below.

## Work packages

### WP0: Land 0.9.0 and restore a deterministic suite (prerequisite)

1. **Commit the 0.9.0 work.** Make the pending upgrade (pins, tests, docs, KNOWN_BUGS, oracle change) one commit, separate from the issue fixes already on `main`.
2. **Make the XML provider explicit.** `d5809b9` added `jsdom` as a dev dependency for the chat security tests. [`tests/setup-xml-provider.mjs`](../../tests/setup-xml-provider.mjs) silently prefers `jsdom` whenever it is installed, so every library-facing suite switched DOM implementations.
   - Evidence (2026-10-12): with `jsdom` installed, 43 suites pass and 5 fail. With it moved aside, all 46 library suites pass, and only the two new `jsdom`-based UI tests fail.
   - Fix: choose the provider explicitly, keeping xmldom as the default lane and adding a named `jsdom` lane. Don't let an unrelated dev dependency choose it.
3. **Triage the `jsdom` failures.** They are `browser_document_session_tests` (generated `ns1` prefix not declared, `PACKAGE_OPERATION_FAILED`), `docx_cross_host_parity_tests` (`BATCH_OPERATION_FAILED`), `ooxml_formatting_visual_tests` and `golden-guardrail`.
   - Word runs the add-in in a browser DOM, which is closer to `jsdom` than to xmldom, so these may be real browser-path library defects.
   - Reproduce each with public APIs. If one is a library defect, document it with a portable reproducer (see the rule in [KNOWN_BUGS.md](KNOWN_BUGS.md)).
   - Decide whether the `jsdom` lane becomes a required gate.

**Exit:** the 0.9.0 work is committed; the default suite is green on an explicit provider; each `jsdom` failure is either filed upstream or explained.

**Done (2026-10-12):**
- The 0.9.0 work is committed as `d0386a6`.
- `tests/setup-xml-provider.mjs` now selects xmldom by default, or jsdom when `DOCX_TEST_XML_PROVIDER=jsdom` is set; it fails loudly on an unknown value. The default suite passes 48/48, and the jsdom lane is an opt-in diagnostic.
- None of the five jsdom-lane failures is a user defect. Each was compared against real Chromium (Edge 155 headless running the library bundle):
  - Four are jsdom-specific (class B). jsdom gives undeclared-namespace attributes generated `nsN` prefixes, and the library's regex removal of `xmlns:` declarations from serialized fragments only works when the original prefixes survive.
  - `golden-guardrail` is serializer byte differences only (class C). Chromium matches xmldom byte-for-byte on `table_reconcile`.
- The add-in never calls `configureXmlProvider`, so in Word the library uses WebView2's native DOM. Details are in [KNOWN_BUGS KB-8](KNOWN_BUGS.md).
- Optional library hardening (do not depend on serializer prefix preservation) is noted in KB-8. Nothing was filed, because no user-facing defect exists.

### WP1: Word Flat OPC body layer (re-scoped 2026-10-17)

**Decision (owner, 2026-10-17):** the Word host layer stays in the plugin, and the library stays host-independent. Library ARCHITECTURE.md says "No Word add-in entrypoints or host-specific integration layers are part of this package." A `/word` subpath was built and reviewed for 0.10.0 and then dropped for that reason. `consumer-core.js` and `word-source-identity.js` stay plugin code.

What remains of WP1 is to remove the plugin's imports of unexported library internals, without adding Word-specific API to the library. The plugin imports from outside the library root today (sent to the library side on 2026-10-17 for a decision):

| Import | Use | Plugin can do without it? |
| --- | --- | --- |
| `applyOperationsToDocumentXml` (`/standalone-runner`) | Batch execution; no-write format re-check | No, but it is a documented public subpath |
| ~~`createParser`/`createSerializer`, `parseXml`/`serializeXml` (`adapters/xml-adapter.js`)~~ | Flat OPC part parse and serialise | **Done 2026-10-17:** replaced with the root `parseOoxmlSafe`/`serializeOoxml`, the same functions those adapter helpers wrap |
| ~~`getPartSpec` (`services/package-parts.js`)~~ | Content and relationship types for comment/numbering parts | **Done 2026-10-17:** a local five-row part table, as the library recommended (+5 lines, one fewer unversioned import) |
| ~~`buildDocumentFragmentPackage` (`services/package-builder.js`)~~ | Wrap a bare fragment; test fixtures | **Done 2026-10-17:** the Word layer no longer builds fragment packages. It inspects a Flat OPC package or a bare `w:document`, and refuses to prepare a write without a package (`WORD_FLAT_OPC_REQUIRED`). Tests still use the builder as a fixture helper |
| `reconcileCommentSiblingParts`, `MemoryZip` | Comment-part merge with the root's zip-shaped `ensureComments*ArtifactsInZip` | Only if the library offers a general-purpose parts-level comment merge or exports these. The library lists a parts-level merge as a candidate for a later release |

**Progress (2026-10-17):** step 1 removed the bare-fragment path (`buildDocumentFragmentPackage`, `extractReplacementNodesFromOoxml` and two local helpers). Production never used it, because every `getOoxml()` source and checkpoint is a full package. New assertions in `docx_consumer_core_tests` pin the contract. 4,363 → 4,330 lines.
Step 2 replaced `getPartSpec` with a local part table (4,330 → 4,335). Two deep imports remain, `reconcileCommentSiblingParts` and `MemoryZip`, both for the comment-part merge.
Step 3 (2026-10-17): sent the library owner a request for a general-purpose, parts-level comment merge (string in, string out, with sibling reconciliation and the existing-thread-metadata refusal), or else documented exports of those two. WP1 waits on that decision. Until then the two imports stay, with the library pinned exact.
Host lanes on steps 1–2 (2026-10-17): Office.js passed (20 cases, 26 checks, 3 restores), the Word oracle gives 88 PASS / 0 FAIL, and the NDA replay passes 16/16 (51/51). Artifacts: `.cache/reliability/wp1-officejs`, `wp1-oracle`, `.cache/golden/wp1-direct`.
Step 4 (2026-10-17): the owner approved a library `mergeCommentParts` for 0.11.0. Branch `wp1-comment-merge` (`e232326`) switches the comment merge to it, which drops the last two deep imports (`MemoryZip`, `reconcileCommentSiblingParts`) and the zip-shaped helpers, and adds a test for the existing-thread-metadata refusal. Production lines 4,335 → 4,297. Against the 0.11.0-pre.1 tarball: offline 44/44, Office.js passed (20 cases, 26 checks, 3 restores), Word oracle 88/0, NDA replay 16/16. Results were sent to the library. **Waiting on:** 0.11.0 on npm. Then pin it, re-run the gates and merge the branch, which completes WP1.

**Exit:** `src/` imports only the library root and documented subpaths; the Word layer shrinks where a general-purpose library primitive replaces local code. Steps that need a library change wait for the library owner's decision.

### WP2: Tighten the inspection session (plugin-only; re-scoped and done 2026-10-17)

**Decision (owner, 2026-10-17):** no library agent-session API. Opaque handles and cursors, prompt text budgets and handle lifecycle (invalidate after a write, expiry) are LLM-harness policy, and the export-tail hook is Word glue, so `document-inspection-session.js` stays in the plugin. The library-side review agreed: `examples/agent-session-wrapper.mjs` stays a dev-only example, and production harnesses own their transport. Search, ranges, context and paging are already delegated to `inspectDocumentParts`.

Work:
- Replace the hand-rolled descriptor checks (`hasStrongDescriptor`, `descriptorsMatch`) with the library's public strict target resolution: `preflightOperations(..., { strictTargets: true })` from `/standalone-runner` (read-only, returns the resolved paragraph or a structured stale/ambiguous error), or the root `resolveTargetParagraph` / `findStrictTargetCandidates` for a pure lookup. Keep binding handles to {document id, revision view, revision token}.
- Shrink whatever else becomes redundant once resolution is the library's, without changing the tool contract.
- If resolution doesn't expose something the session needs, send the case to the library owner as a possible general-purpose primitive. No Word- or agent-specific API goes upstream.

**Exit:** the session has no local target-matching logic, production lines fall, and stale, ambiguous, cross-document and expired-handle cases still refuse with no write (the existing offline contracts pass unchanged, plus the host lanes).

**Progress (2026-10-17):**
- Step 1, local target matching removed. `descriptorsMatch` and `hasStrongDescriptor` are gone from `document-inspection-session.js`. `resolveHandles` already refuses unless the current source has the same document id, revision view and Word revision token as the handle's snapshot. An equal token means the canonical source is unchanged, so a field-by-field descriptor comparison could never fail. The apply path then resolves every target strictly in the library (`strictTargets: true`, `expectedRevision: true`), so `preflightOperations` would only repeat that check. Library inspection always returns strong descriptors, so the strong-descriptor filter and its "omitted" warning were dead. Production lines 4,436 → 4,415. Offline 44/44.
- Step 2, redundant copies trimmed. Handles store only the paragraph index, since the descriptor returned to the compiler is always re-read from the fresh library inspection, so the deep clone and freeze helpers are gone. Session and cursor IDs use `crypto.randomUUID` (WebView2 over HTTPS and Node), without the old fallbacks. The intent compiler's `WEAK_DOCUMENT_EDIT_TARGET` check is removed: resolved descriptors always come from the library and are always strong. The handle-kind check stays. Production lines 4,415 → 4,382. Offline 44/44.
- Step 3: one helper now builds the `inspectDocumentParts` options for fresh queries and cursor continuations, replacing three copies (one was unreachable). Production lines 4,382 → 4,363. Offline 44/44.
- **Done (2026-10-17).** Host lanes on the final state: Office.js passed (20 cases, 26 checks including the stale-handle refusal control, 3 checkpoint restores), the Word oracle gives 88 PASS / 0 FAIL, and the NDA replay passes 16/16 (51/51 checks, 40 chat requests, 0 auxiliary). Artifacts: `.cache/reliability/wp2-officejs`, `wp2-oracle`, `.cache/golden/wp2-direct`. WP2 total: 4,436 → 4,363 production lines, with no library change needed.

### WP3: Upstream revision policy and verification; retire workarounds

- **Auto revision policy.** The library chooses `slice-cross-author` or same-author merge from the target's revisions. Delete the add-in's choice in `document-tools.js`.
- **Verification helper.** The library verifies the committed body against the prepared output, using its public revision segments. Delete the local review-segment merge.
- **Requested-text verification (KB-7).**
  - 0.9.0's `ACCEPT_INTENT` skips Markdown-looking text. That is why `rewrite_paragraph` can return `ok` with literal underscore runs collapsed.
  - Ask upstream to apply the intent check to `structuredContent: false` operations, and to fix the Markdown pre-processor (`pipeline/markdown-processor.js`) for literal text.
  - Delete the integration test's known-defect gate when its reproducer exits 1.
- **Retire `TRACKING_POLICY_MISMATCH` (KB-5).** Rely on the library's `UNTRACKED_CREATED_REVISION`.

**Exit:** `document-tools.js` holds only tool wiring. KB-5 and KB-7 are closed.

**Done (2026-10-17, pending the 0.10.0 release):**
- Library 0.10.0 adds `existingRevisions: 'auto'`, `compareDocumentInspections` and the KB-7 fix (CommonMark underscore flanking, and literal inline Markdown for a caller's `structuredContent: false`).
- Plugin branch `phase2-wp1-wp3`: plain-text replace/rewrite operations send `existingRevisions: 'auto'` instead of choosing `slice-cross-author` locally; post-commit verification uses `compareDocumentInspections` instead of the local `acceptedShape`/`reviewSegments`; the `TRACKING_POLICY_MISMATCH` guard (KB-5) and the "Add Date Row" known-defect gate (KB-7) are removed.
- Production lines 4,478 → 4,436. Root and MCP pins are exact 0.10.0 (published 2026-10-17, tag `v0.10.0` at `04faf9b`); the MCP server was removed later the same day. Offline 49/49, the MCP workflow test, both builds, and the KB-7 reproducer exits 1.
- 0.10.0 also counts unattributed revisions as another author's (`EXISTING_REVISIONS` instead of `UNSAFE_REVISION_NESTING` under merge-same-author). The plugin maps neither code, so nothing changes here.
- Host lanes (2026-10-17): Office.js passed (20 cases, 26 checks, 3 restores), the Word oracle gives 88 PASS / 0 FAIL, and the NDA replay passes 16/16 (51/51). See the [0.10.0 validation](../validation-reports/2026-10-17-docx-redline-v0100.md).

### WP4: Remove expedient Word JS API uses and dead code

These items have no library dependency and can start right after WP0.

- **Glance checks.** Use library inspection text instead of `body.text` (`taskpane.js`, around line 804).
- **Dead code.** Delete:
  - `src/taskpane/debug_structure.js` (unimported);
  - `src/commands/commands.js` (an Outlook template stub), after checking it isn't wired to a function file;
  - the unreachable checkpoint-capture branch in `taskpane.js`;
  - the unused `baseTrackingMode` parameter in `word-ooxml.js`.
- **Host paragraph count.** Remove it and `detectExportTailIndex` once WP1 provides export-tail handling.
- **Test XML setup (done 2026-10-13).** The default lane relies on the library's own xmldom fallback.
  - `tests/setup-xml-provider.mjs` went from 97 lines to 16. It no longer installs global `DOMParser`, `XMLSerializer`, `document`, `window` or `Node`, and no longer patches NodeList iteration. It only configures jsdom when that lane is explicitly selected.
  - The tests that used the globals now call the library's `parseXml` / `serializeXml`, the same entry point production uses: `integration_tests`, `dom_fallback_smoke` (now a library-fallback check) and the manual `word-desktop/list-regression` lane.
  - Default lane: 48/48. The jsdom lane is unchanged at the same 5 KB-8 diagnostics.

**Exit:** the KNOWN_BUGS "Expedient uses to retire" list is empty except items blocked on WP1.

**Done (2026-10-16), except the host paragraph count:**
- Glance checks build their prompt text from `inspection.paragraphs[].exactText` through `createWordDocumentSourceReader`. They no longer read Word's `body.text`.
- `createCheckpoint` became `createAutoCheckpoint(toolName, sourceOoxml)`. The only caller always passes the prepared source, so the `getOoxml` fallback and the manual/non-silent paths were dead.
- `baseTrackingMode` is gone from `withNativeTrackingDisabled` and the batch runner. `word_pure_ooxml_batch_tests` now counts the tracking-mode load sync.
- Deleted `src/taskpane/debug_structure.js` (unimported) and `src/commands/commands.js`. The manifest still declares `FunctionFile` → `commands.html` but has no `ExecuteFunction` action. So the page stays as an empty function file, and webpack no longer builds a `commands` entry for it.
- Production lines went from 4,549 to 4,478, and files from 20 to 19 (`count-document-agent-code.mjs`; `debug_structure.js` was outside the counted surface).
- Remaining items: the KNOWN_BUGS list keeps only the WP1-blocked export-tail item, KB-5 (WP3) and the upstream package logic.
- Gates:
  - Offline: 49/49.
  - Production and development builds pass.
  - Actual Office.js: passed, 20 cases with 26 host checks. The run needs a `WORD_HOST_VALIDATION` build passed with `--dist-dir`.
  - Word oracle: 88 PASS, 0 FAIL. Run it against `officejs-fixtures.json`, not the export `manifest.json`.
  - NDA replay: 16/16, 51 checks, 0 failed.
  - No host lane runs glance checks, so their new text source is covered only by the offline build.

### WP5: New capabilities only as library releases plus a thin mapping

- Adopt 0.9.0 `list-format` (`convert-plain`, `change-format`, `join`) through the [list migration follow-up](completed/2026-09-30-canonical-list-migration-follow-up.md) WP1–WP3. Each new intent should be a schema entry plus an operation mapping, with no plugin-side numbering or marker logic.
- Measure the lines added per capability.
- Re-enabling structural table edits or table append stays out of scope until KB-1 has Office.js evidence.

**Exit:** governed by the list plan's own acceptance.

**Done (2026-10-14):** three list actions are exposed as one-to-one `list-format` mappings, adding 112 production lines (4,437 → 4,549). Offline is 49/49; Office.js passes with 20 cases; the oracle gives 88/0; the real-Word NDA replay is 16/16 with step 12 applied. See the [2026-10-14 list-format validation](../validation-reports/2026-10-14-list-format-intents.md). The list plan's facade-parity item remains; its library-issues cleanup was done on 2026-10-17 ([closeout](../validation-reports/2026-10-17-library-issues-closeout.md)).

## Separate track: release readiness (not simplification)

These don't reduce the plugin, but they gate users actually getting today's work:

- **Deploy.** The hosted add-in has none of the 0.9.0 upgrade, the #2 sanitizer and CSP, or the #3 custom-model fixes until a deployment.
- **Live models and Word Online.** Run the live fast/slow provider replay and the Word Online checks. Both are still unrun shipping gates.
- **Issue #3.** Decide on automatic model fallback after transient retries run out.
- **Dev certificate.** The local HTTPS dev certificate lasts 30 days. Renew it before host lanes; it last expired on 2026-10-03.

## Measures

- Production lines and files from `count-document-agent-code.mjs`, reported per package against this plan's 4,437-line baseline.
- Internal library subpath imports in `src/`: six today, target zero.
- Consumer modules that parse or build OOXML directly, as opposed to through library calls: target zero outside host transport.
- KNOWN_BUGS expedient Word JS uses: target zero.
- Gates unchanged and passing after every package. No capability is converted to a refusal to make a package pass.

## Open owner decisions

1. **Size target.** Decided 2026-10-17: no fixed target. Keep shrinking.
2. **Order.** The recommended order is WP0 → WP4 → WP1 → WP3 → WP2.
   - WP4 is quick and independent.
   - WP1 removes the most library coupling.
   - WP2 was the largest API design; since 2026-10-17 it is plugin-only cleanup.
   - Alternatively, start WP2 first if inspection reliability is the bigger pain.
3. **Library API shape.** Decided 2026-10-17: the library takes only host-independent, general-purpose API. Word-specific layers (Flat OPC round-trip, Word export noise, and probably the WP2 agent session's Word binding) stay in the plugin.
4. **Browser lane.** Recommendation from WP0: don't make jsdom a required gate. Keep xmldom as the gate and add a headless Chromium/Edge lane that runs the library bundle, advisory at first and required once it is stable. `scripts/browser-document-validation.html` is a manual starting point. Agree? *(2026-10-17: with the browser demo and MCP server gone, the add-in's only runtime is Word's WebView2, which the Office.js, Word oracle and NDA replay lanes already exercise. A separate headless library-bundle lane would mostly check the library, so it belongs with the library if anywhere.)*
5. **Release track.** Deploy the current fixes before phase 2 starts, or batch them with WP0?
