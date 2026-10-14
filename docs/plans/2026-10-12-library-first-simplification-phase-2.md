# Library-first simplification, phase 2

**Date:** 2026-10-12 (America/Vancouver)
**Status:** WP0 complete (2026-10-12); WP5 complete for the exposed list actions (2026-10-14); WP4 partly done (test XML setup). WP1–WP3 not started; the open owner decisions are at the end.
**Baseline:**
- Exact `@ansonlai/docx-redline-js@0.9.0` in the add-in and MCP.
- HEAD `a98f339` plus the uncommitted 0.9.0 upgrade work.
- **4,437 production lines across 20 files** (`node scripts/count-document-agent-code.mjs`). For comparison: 8,632 lines at the WP0 baseline of the [phase-1 plan](completed/2026-10-02-library-first-lightweight-agent-reliability.md), and 4,402 at its WP5 close.

**Related:**
- [Known bugs and the Word JS API inventory](KNOWN_BUGS.md)
- [Canonical list migration follow-up](2026-09-30-canonical-list-migration-follow-up.md)
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
| `chat/document-inspection-session.js` | 551 | Handle store, focused inspection, search, pagination, staleness and expiry checks | A published, browser-safe **agent session** API. The library's `examples/agent-session-wrapper.mjs` shows the pattern, but it is development-only and uses private imports |
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
   - Reproduce each with public APIs. If one is a library defect, file it in `docs/library-issues/` with a reproducer.
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

### WP1: Upstream the Word Flat OPC body surface

- Move package-part, relationship, comment-merge and insertion-package assembly from `consumer-core.js` behind one public library function. The same applies to the export-noise normalisation and tail detection in `word-source-identity.js`.
- The function takes the Word body Flat OPC, a batch and options, and returns the result, the insertion package and receipts.
- Remove the six internal subpath imports.
- The Word transport itself (`getOoxml`, one `insertOoxml`) stays in the plugin.
- **Exit:** `consumer-core.js` and most of `word-source-identity.js` are deleted or reduced to host glue, with package-root imports only. All gates pass unchanged, including the sequential comment workflow and checkpoint restores.

### WP2: Upstream the agent session (inspection, handles, search)

- Replace the handle store, focused inspection, search and pagination in `document-inspection-session.js` with a published library session API.
- That API binds handles to a source revision and view, expires them after a write, and returns bounded pages with locations.
- The plugin keeps only the tool schema and the mapping from the live Word source to the library session.
- **Exit:** most of `document-inspection-session.js` is gone. Stale, ambiguous, cross-document and expired-handle cases still refuse with no write (the existing offline contracts pass unchanged).

### WP3: Upstream revision policy and verification; retire workarounds

- **Auto revision policy.** The library chooses `slice-cross-author` or same-author merge from the target's revisions. Delete the add-in's choice in `document-tools.js`.
- **Verification helper.** The library verifies the committed body against the prepared output, using its public revision segments. Delete the local review-segment merge.
- **Requested-text verification (KB-7).**
  - 0.9.0's `ACCEPT_INTENT` skips Markdown-looking text. That is why `rewrite_paragraph` can return `ok` with literal underscore runs collapsed.
  - Ask upstream to apply the intent check to `structuredContent: false` operations, and to fix the Markdown pre-processor (`pipeline/markdown-processor.js`) for literal text.
  - Delete the integration test's known-defect gate when its reproducer exits 1.
- **Retire `TRACKING_POLICY_MISMATCH` (KB-5).** Rely on the library's `UNTRACKED_CREATED_REVISION`.

**Exit:** `document-tools.js` holds only tool wiring. KB-5 and KB-7 are closed.

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

### WP5: New capabilities only as library releases plus a thin mapping

- Adopt 0.9.0 `list-format` (`convert-plain`, `change-format`, `join`) through the [list migration follow-up](2026-09-30-canonical-list-migration-follow-up.md) WP1–WP3. Each new intent should be a schema entry plus an operation mapping, with no plugin-side numbering or marker logic.
- Measure the lines added per capability.
- Re-enabling structural table edits or table append stays out of scope until KB-1 has Office.js evidence.

**Exit:** governed by the list plan's own acceptance.

**Done (2026-10-14):** three list actions are exposed as one-to-one `list-format` mappings, adding 112 production lines (4,437 → 4,549). Offline is 49/49; Office.js passes with 20 cases; the oracle gives 88/0; the real-Word NDA replay is 16/16 with step 12 applied. See the [2026-10-14 list-format validation](../validation-reports/2026-10-14-list-format-intents.md). The list plan's facade-parity and library-issues cleanup items remain.

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

1. **Size target.** Proposed: at most 3,300 production lines after WP1–WP4, about a 25% cut from 4,437. Is there a harder target?
2. **Order.** The recommended order is WP0 → WP4 → WP1 → WP3 → WP2.
   - WP4 is quick and independent.
   - WP1 removes the most library coupling.
   - WP2 is the largest API design.
   - Alternatively, start WP2 first if inspection reliability is the bigger pain.
3. **Library API shape.** Should the Flat OPC body surface and the agent session be public library modules (browser-safe, documented, semver-covered), or a Word-specific companion package?
4. **Browser lane.** Recommendation from WP0: don't make jsdom a required gate. Keep xmldom as the gate and add a headless Chromium/Edge lane that runs the library bundle, advisory at first and required once it is stable. `scripts/browser-document-validation.html` is a manual starting point. Agree?
5. **Release track.** Deploy the current fixes before phase 2 starts, or batch them with WP0?
