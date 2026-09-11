# `@ansonlai/docx-redline-js` v0.5.4 Upgrade Plan

**Date:** 2026-09-09  
**Status:** Active — WP0 through WP3 and automated WP4 validation are complete.
Interactive browser and Word Desktop smoke tests remain manual because those
application controllers were unavailable in the validation environment.  
**Target:** Upgrade both consumers from exact version `0.4.0` to exact version
`0.5.4`.

This plan supersedes and merges:

- `2026-09-02-docx-redline-js-v0.4.0-upgrade.md`;
- `2026-09-07-docx-redline-js-v0.5.0-upgrade.md`.

The immediate goal is a small, low-risk dependency migration. Broader work to
make this repository a thinner host for the package is recorded under
**Future learnings**, but is deliberately outside this upgrade.

## Current baseline

- Both `package.json` files and lockfiles resolve exact `0.4.0`.
- The v0.4 compatibility layer already checks structured failures before
  `hasChanges`, explicitly sanitizes model-originated content, and preserves
  error codes at the Word, browser, and MCP boundaries.
- `integration.js` and `word-route-change.js` were removed on 2026-09-04 after
  their reachability was characterized. The add-in now uses the injected
  `loadRedlineAuthor` provider. Resurrection guards cover the deleted modules.
- The browser demo still applies operations one at a time. The MCP server has a
  paragraph edit service, not a public batch tool. This upgrade must not add a
  batch feature merely to consume the new package version.
- Existing v0.4 workarounds include caller-side input preparation because the
  v0.4 standalone runner did not forward `sanitizeInput`, and clean-range
  filtering for transient multi-paragraph revisions. They may be removed only
  after focused v0.5.4 parity tests pass.
- The repository has no root `npm test` script. Use the targeted Node suites and
  builds listed below, plus the wider runnable test inventory where practical.

## Release changes that affect this repository

The release notes for
[v0.5.0](https://github.com/AnsonLai/docx-redline-js/releases/tag/v0.5.0),
[v0.5.1](https://github.com/AnsonLai/docx-redline-js/releases/tag/v0.5.1),
[v0.5.2](https://github.com/AnsonLai/docx-redline-js/releases/tag/v0.5.2),
[v0.5.3](https://github.com/AnsonLai/docx-redline-js/releases/tag/v0.5.3), and
[v0.5.4](https://github.com/AnsonLai/docx-redline-js/releases/tag/v0.5.4) are the
source of truth for this migration.

| Release | Relevant change | Minimal consumer response |
|---|---|---|
| 0.5.0 | `getParagraphText` now returns accepted/current canonical text, and paragraph fingerprints changed. | Rebuild descriptors from the current document; do not reuse fingerprints produced under v0.4.0. Add canonical-text coverage for deletions, moves, breaks, and special hyphens. |
| 0.5.0 | Whole-paragraph deletion with comments and invalid/ambiguous comment anchors fail closed with structured errors. | Preserve package codes and original document state. Do not convert errors to no-ops or retry commented paragraph deletion automatically. |
| 0.5.0 | Batch execution defaults to progressive (`atomic: false`). | Pass `atomic: true` on any existing all-or-nothing batch call. Do not introduce a new batch workflow in this migration. |
| 0.5.0 | Same-author revisions merge by default; third-party revisions still fail closed. Replacement pairing, structured content, mutation receipts, stricter package validation, and paragraph-mark revisions are built in. | Keep the package defaults unless current behavior requires an explicit option. Preserve receipts and validation diagnostics through adapters. Verify rather than reimplement these behaviors. |
| 0.5.1 | `slice-cross-author` allows editing inside another author's revisions. It is opt-in and backward-compatible. | Do not enable it during the version migration. Keep third-party revisions fail closed under `merge-same-author`. |
| 0.5.2 | Cross-author slicing preserves whitespace and hyperlinks more reliably and fails closed with `PATCH_ROUNDTRIP_MISMATCH` when accepted-view verification fails. | Propagate this and all future structured codes generically; assert that the original OOXML remains unchanged on failure. |
| 0.5.3 | Adds explicit paragraph `restore` operations and refuses ordinary edits to foreign whole-paragraph deletions with `FOREIGN_PARAGRAPH_MARK_DELETION`. | Do not translate ordinary edits into `restore` automatically. Surface the refusal and preserve the source paragraph. Restoration is a future explicit product feature. |
| 0.5.4 | Adds operation-level, baseline-delta, mutation-envelope, and whole-document duplicate revision-ID validation. New errors include `GENERATED_OOXML_INVALID`, `REJECTED_INSERTION_STATE_REQUIRED`, and `UNSAFE_REVISION_BOUNDARY`. | Retain generic error codes, receipts, and validation summaries; verify no mutation is committed on failure. |
| 0.5.4 | Adds explicit rejected-view insertion and improves exact occurrence/offset targeting, whitespace alignment, hyperlink preservation, right-to-left surgical edits, and revision-ID allocation. | Benefit from the reliability fixes without enabling rejected-view editing. Refresh live target descriptors before applying edits and supply `occurrence` for repeated exact anchors. |

The v0.5.4 JavaScript API is documented as backward-compatible with existing
redline, comment, formatting, list, and table operations. Its CLI output contract
changed, but this repository calls the JavaScript API and therefore needs no CLI
parser migration.

## Migration principles

1. Pin `0.5.4` exactly in both npm projects and regenerate both lockfiles with
   npm; do not hand-edit lockfile resolution or integrity fields.
2. Keep current user-visible behavior. New cross-author slicing, restoration,
   rejected-view insertion, browser batching, MCP batching, and normalization UI
   are not part of this upgrade.
3. Treat `status: 'error'` and `result.error` as authoritative before checking
   `hasChanges` or output fields.
4. Preserve any package `error.code` generically. Tests should cover known codes,
   but production code must not depend on an exhaustive allowlist.
5. Keep model-boundary sanitization explicit and keep user/MCP-authored literal
   text unsanitized by default.
6. Let the package own diffing, revision lifecycle rules, validation, accepted-
   view verification, and revision-ID allocation. Do not add consumer-side
   versions of these algorithms.
7. A failed operation must leave document XML, comments, numbering,
   relationships, package state, and MCP dirty state unchanged.
8. Remove a v0.4 workaround only when a focused v0.5.4 regression proves the
   public package path has equivalent behavior.

## Required work

### WP0 — Freeze the v0.4 baseline

**Status:** Complete on 2026-09-09.

Before changing dependencies:

1. Record `git status --short` and preserve all unrelated working-tree changes.
2. Run the existing compatibility and adapter suites:

   ```powershell
   node tests/redline_result_contract_tests.mjs
   node tests/docx_redline_v040_compat_tests.mjs
   node tests/addin/word_operation_runner_adapter_tests.mjs
   node tests/addin/shared_operation_bridge_tests.mjs
   node tests/mcp_docx_redline_service_tests.mjs
   node tests/no_legacy_shared_operation_bridge_tests.mjs
   npm run build:dev
   ```

3. Save failures as baseline observations rather than changing behavior during
   the dependency commit.

**Acceptance:** current failures, if any, are understood; the version change is
not used to absorb unrelated refactors.

#### WP0 execution record — 2026-09-09

- Captured the dirty working-tree baseline before running checks. It contains
  the existing v0.4 compatibility implementation, tests, fixture updates, and
  the completed deletion of `integration.js` and `word-route-change.js`; WP0 did
  not modify those files.
- Confirmed both dependency trees resolve exactly
  `@ansonlai/docx-redline-js@0.4.0`:
  - `npm ls @ansonlai/docx-redline-js` — passed;
  - `npm --prefix mcp/docx-server ls @ansonlai/docx-redline-js` — passed.
- Targeted results:
  - `node tests/redline_result_contract_tests.mjs` — passed;
  - `node tests/docx_redline_v040_compat_tests.mjs` — passed;
  - `node tests/addin/word_operation_runner_adapter_tests.mjs` — passed;
  - `node tests/addin/shared_operation_bridge_tests.mjs` — passed;
  - `node tests/mcp_docx_redline_service_tests.mjs` — passed;
  - `node tests/no_legacy_shared_operation_bridge_tests.mjs` — passed;
  - `npm run build:dev` — passed; webpack 5.102.1 compiled successfully.
- No baseline failures were found. Node emitted the existing
  `MODULE_TYPELESS_PACKAGE_JSON` warning while importing add-in `.js` modules as
  ES modules in three suites. This is non-blocking and is not being addressed by
  the dependency migration because changing the root package module type could
  affect webpack and Office add-in behavior.

### WP1 — Add focused v0.5.4 contract coverage

**Status:** Complete on 2026-09-09.

Add `tests/docx_redline_v054_compat_tests.mjs`, using public APIs wherever
possible. Cover only contracts that can affect existing hosts:

1. Canonical paragraph text excludes deleted and moved-from content while
   preserving breaks, soft hyphens, and non-breaking hyphens.
2. Paragraph descriptors/fingerprints are freshly generated from current
   canonical text and no v0.4 fingerprint is restored from a cache or session.
3. `INVALID_OPERATION`, `COMMENTED_CONTENT_DELETE`, `ANCHOR_NOT_FOUND`, and
   `AMBIGUOUS_ANCHOR` are errors, not no-ops.
4. `PATCH_ROUNDTRIP_MISMATCH`, `FOREIGN_PARAGRAPH_MARK_DELETION`,
   `GENERATED_OOXML_INVALID`, `REJECTED_INSERTION_STATE_REQUIRED`, and
   `UNSAFE_REVISION_BOUNDARY` pass through adapters unchanged. Use representative
   synthetic results when constructing every engine state would make the
   consumer test brittle.
5. A package verification failure returns the original input unchanged and does
   not expose committable artifacts.
6. An existing batch path, if found, explicitly uses `atomic: true` and does not
   commit partial results. If no batch path exists, add a source assertion that
   prevents accidentally relying on the package default when one is introduced.
7. Default `merge-same-author` behavior works for same-author follow-up edits and
   continues to reject foreign revisions. Do not exercise `slice-cross-author`
   as a product path.
8. Whole-paragraph Accept All / Reject All remains symmetric, revision IDs are
   unique at document scope, and tabs/breaks/hyperlinks survive representative
   edits.
9. The consumer sanitization bridge produces the current model-boundary behavior
   when invoking v0.5.4. Literal dollar values and template tokens remain
   unchanged when sanitization is disabled.

**Acceptance:** the new suite detects the meaningful v0.4-to-v0.5.4 contract
changes without duplicating package algorithms in the consumer tests.

#### WP1 execution record — 2026-09-09

- Added `tests/docx_redline_v054_compat_tests.mjs` using root exports and the
  supported `standalone-runner` subpath. The suite always runs consumer error-
  propagation and explicit-atomicity guards, and version-gates package behavior
  while the project remains pinned to v0.4.0.
- Verified the full suite against an unpacked official npm v0.5.4 artifact via
  `DOCX_REDLINE_PACKAGE_ROOT`; all assertions passed without changing either
  project dependency tree.
- Covered canonical accepted-view text and fresh fingerprints, structured
  operation/comment errors, commented paragraph deletion, atomic rollback and
  receipt dispositions, same-author revision merging, foreign-author refusal,
  document-scoped revision-ID uniqueness, Accept All / Reject All symmetry,
  hyperlinks and structural text, and literal-safe sanitization policy.
- Audited `src`, `browser-demo`, and `mcp`: no production fingerprint persistence
  or reuse exists today.
- Finding: v0.5.4's full-document `applyOperationToDocumentXml` path still does
  not forward `sanitizeInput` to the redline engine. A direct parity assertion
  retained the model preface. The suite therefore verifies the safe current path
  through `prepareOperationInput`, and WP3 must keep that bridge until an upstream
  release closes the gap.
- Baseline verification under installed v0.4.0 passed the always-on guards and
  skipped only the v0.5.4 runtime section as designed. The v0.4 compatibility,
  result-contract, Word adapter, shared bridge, and MCP service suites all
  remained green.
- The existing non-blocking `MODULE_TYPELESS_PACKAGE_JSON` warning remains; no
  module-mode configuration was changed.

### WP2 — Upgrade both dependency trees

**Status:** Complete on 2026-09-09.

Run:

```powershell
npm install --save-exact @ansonlai/docx-redline-js@0.5.4
npm --prefix mcp/docx-server install --save-exact @ansonlai/docx-redline-js@0.5.4
npm ls @ansonlai/docx-redline-js
npm --prefix mcp/docx-server ls @ansonlai/docx-redline-js
```

Verify both manifests and lockfiles resolve exact `0.5.4`, with no second copy or
unexpected prerelease. Keep the dependency and lockfile changes in one
independently revertible commit.

**Acceptance:** a clean install in either project resolves `0.5.4` exactly and
both development and production bundles resolve current public exports.

#### WP2 execution record — 2026-09-09

- Installed exact `@ansonlai/docx-redline-js@0.5.4` in the root add-in and
  `mcp/docx-server` projects using npm, updating both manifests and generated
  lockfiles.
- `npm ls @ansonlai/docx-redline-js` passed in both projects and reported one
  exact `0.5.4` copy with no duplicate or prerelease resolution.
- Both development and production webpack builds resolved the installed public
  exports successfully.

### WP3 — Make the smallest compatibility changes

**Status:** Complete on 2026-09-09.

Audit these boundaries:

- `src/taskpane/modules/docx-redline-js-integration/redline-result.js`;
- `src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js`;
- `src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js`;
- `browser-demo/demo.js`;
- `mcp/docx-server/src/services/docx-redline-js-service.mjs`;
- `mcp/docx-server/src/server.mjs`.

Apply only changes demonstrated by WP1:

1. Normalize a package result once at each host boundary. Preserve `status`,
   `error.code`, message, warnings, operation index, receipt(s), rollback state,
   and validation summary when present.
2. Check failure before `hasChanges`, replacement OOXML, or native-fallback
   decisions. A package error must never select another local reconciliation
   algorithm.
3. Prefer generic code propagation over adding a growing switch statement. Add
   product-specific handling only where behavior differs:
   - do not retry `COMMENTED_CONTENT_DELETE`;
   - do not reinterpret `FOREIGN_PARAGRAPH_MARK_DELETION` as a restore;
   - refresh context for ordinary stale-target errors at most once;
   - never retry validation, round-trip, or unsafe-boundary failures against the
     same document state.
4. Pass `sanitizeInput: true` directly for raw model content and leave it false
   for literal caller content. Delete `prepareOperationInput` only if the v0.5.4
   parity test proves native package sanitization covers every current call path;
   otherwise keep it temporarily with a documented deletion condition.
5. Delete the clean multi-paragraph range workaround only if v0.5.4 produces the
   expected clean paragraph structure in both package and Word adapter tests.
6. Re-extract paragraph text/descriptors immediately before operations that may
   otherwise use stale state. Do not persist fingerprints across package-version
   or document-session boundaries.
7. Pass `atomic: true` on existing transactional batch calls. Do not add a new
   batch API or convert a sequential UI flow as part of this work package.
8. Retain the injected author provider. Do not rely on the package's fallback
   author to replace product configuration.

**Acceptance:** the production diff is limited to dependency pins, proven
contract adaptation, and removal of obsolete workarounds; no new feature or
architecture layer is introduced.

#### WP3 execution record — 2026-09-09

- Preserved package diagnostics at the add-in, browser, and MCP boundaries,
  including warnings, package error metadata, operation index, receipt(s),
  rollback state, resolved targets, and validation summaries when supplied.
- Kept generic structured-code propagation and failure-before-no-op handling;
  no code allowlist, fallback redline algorithm, automatic restore, cross-author
  slicing, rejected-view insertion, or new batch API was introduced.
- Retained `prepareOperationInput` because WP1 proved that the v0.5.4 full-
  document runner still does not forward `sanitizeInput`. Retained the clean-
  range workaround because removing it was not proven safe by the focused
  contracts.
- Added focused adapter assertions for `GENERATED_OOXML_INVALID`, package error
  stage, receipt, and rollback propagation. Targeted result, package, Word,
  bridge, MCP, and legacy-module guard suites passed.

### WP4 — Verify safety and document fidelity

**Status:** Automated validation complete on 2026-09-09; interactive browser
and Word Desktop smoke tests remain manual.

Run the targeted suites from WP0 and WP1, then the wider runnable Node test
inventory. Exclude paid/live evals and fixture-rewriting performance/golden
harnesses unless intentionally running them in isolation. Upstream's
`DOCX_TEST_CONCURRENCY` setting does not automatically govern this repository's
ad hoc Node commands; serialize any local suites that share mutable fixtures.

Run:

```powershell
npm run build:dev
npm run build
npm run validate
npm ls @ansonlai/docx-redline-js
npm --prefix mcp/docx-server ls @ansonlai/docx-redline-js
```

Perform focused host smoke tests:

- **Browser:** unchanged single and sequential edit flows, redlines on/off,
  comments, lists/tables, repeated anchors, and an intentional failure that
  proves prior UI/package state is not accidentally overwritten.
- **MCP:** open, inspect, edit, save, reopen; verify a failed edit does not change
  XML or dirty state and returns its stable code.
- **Word Desktop:** documents containing existing revisions from the same and a
  different author, comments, hyperlinks, bookmarks, tabs/breaks, lists, tables,
  section properties, and repeated exact target text.

For generated DOCX files, verify that Word opens without repair, Reviewing Pane
replacement grouping remains sensible, Accept All yields the requested current
view, Reject All yields the expected baseline, and save/reopen succeeds. Record
manual results under `docs/validation-reports/`.

**Acceptance:** automated tests and builds pass; failures are non-mutating and
coded; Word round trips representative files without repair or revision-ID
collisions.

#### WP4 execution record — 2026-09-09

- A serialized wider Node inventory passed 28/28 runnable suites. Paid/live
  evals, fixture-rewriting golden/performance harnesses, the XML-provider setup
  module, and the Word Desktop regression script were intentionally excluded.
- Replaced the obsolete numbering prototype mock with public
  `buildDocumentFragmentPackage` coverage for `includeNumbering` true, false,
  and omitted behavior.
- Added an MCP workflow smoke covering create/open, stale-target refusal,
  unchanged XML and dirty state after failure, tracked edit, save, reopen,
  package validation, and session close.
- `npm run build:dev`, `npm run build`, and `npm run validate` passed. The
  production build retained only the existing webpack asset-size warnings.
- Both dependency-tree checks passed at exact `0.5.4`. Relevant add-in, browser,
  and MCP source files also passed Node syntax checks.
- The in-app browser reported no available browser instance, and the Word
  Desktop controller's native pipe was unavailable. Interactive host checks are
  therefore recorded as manual follow-ups rather than inferred successes. See
  `docs/validation-reports/2026-09-09-docx-redline-js-v0.5.4-automated.md`.

## Explicit non-goals for this migration

- Enabling `slice-cross-author`.
- Automatically converting failed redlines into `restore` operations.
- Exposing rejected-view insertion in agent tools or UI.
- Adding browser atomic batching or `docx_apply_batch` to MCP.
- Adopting `@ansonlai/docx-redline-js/node` and `openDocx`.
- Creating a new façade layer, moving large modules, enforcing a new import
  architecture, or pursuing line-count targets.
- Adding a runtime package-version assertion when exact manifests, lockfiles,
  dependency-tree checks, and CI verification already establish the version.
- Changing global revision normalization or comment-deletion policy.
- Parsing the package CLI's compact contract-v3 output.

## Suggested commit sequence

1. `test: add docx-redline v0.5.4 compatibility contracts`
2. `chore: pin docx-redline-js 0.5.4 in add-in and MCP`
3. `fix: adapt host boundaries to docx-redline 0.5.4`
4. `test: record docx-redline 0.5.4 validation results`

Each commit should pass its targeted checks and remain independently reviewable.

## Rollback

1. Restore both manifests to exact `0.4.0` and regenerate both lockfiles using
   npm in their respective projects. Do not hand-edit lockfile metadata.
2. Restore a deleted v0.4 compatibility workaround only if rollback testing
   proves it is required.
3. Retain generic structured-error propagation and explicit sanitization policy;
   both are safe compatibility practices for v0.4.0.
4. Re-run the v0.4 contract suites, builds, and dependency-tree checks.

## Definition of done

- Both consumers and lockfiles resolve exact `0.5.4`.
- Every package failure is checked before no-op or output handling, and arbitrary
  structured codes survive host boundaries.
- Canonical text is used for fresh descriptors; old fingerprints are not reused.
- Existing transactional batch calls explicitly request `atomic: true`.
- Model-originated content is sanitized explicitly while literal caller content
  is preserved.
- Validation, round-trip, unsafe-boundary, comment, and foreign-revision failures
  leave all document/package/session state unchanged.
- Same-author merging and replacement pairing work without opting into new
  cross-author product behavior.
- Representative Word files open, accept/reject, save, and reopen without repair.
- Any remaining v0.4 workaround has a regression and an explicit deletion
  condition.
- No broad refactor or new user-facing feature is required to complete the
  package upgrade.

## Future learnings and architectural follow-ups

These are useful directions for a later rewrite, not acceptance criteria for
v0.5.4.

1. **Use one narrow package façade per host.** The add-in, browser, and MCP have
   different commit mechanics, but can share a canonical operation/result
   vocabulary. This would stop package result normalization and artifact handling
   from spreading through UI and agent modules.
2. **Thin `word-redline-runner.js` to Office.js coordination.** Keep proxy
   loading, `context.sync`, scope acquisition, `insertOoxml`, tracking-mode
   restoration, and checkpoints locally. Move target repair, list/table
   synthesis, diff/revision rules, and revision lifecycle behavior to public
   package operations.
3. **Route agent tools through canonical operations.** `agentic-tools.js` should
   validate product intent and resolve live Word scope, then submit operations.
   It should not instantiate reconciliation pipelines or maintain a fallback
   document engine.
4. **Use full-document batching where the host owns the package.** A future
   browser/MCP design can use explicit `atomic: true`, single-DOM execution,
   receipts, and complete artifact rollback. This should be a separately shipped
   behavior change with byte-level rollback tests.
5. **Evaluate the Node façade for MCP.** Adopt `openDocx` only if it preserves the
   server's session, dirty-state, save/close, and filesystem ownership boundaries
   while returning the artifacts and per-operation diagnostics the server needs.
6. **Replace local workarounds with package guarantees.** Sanitization, clean
   multi-paragraph edits, replacement-node extraction, structured list/table
   rendering, accepted-view verification, and revision-ID allocation should each
   have one owner. Characterize, route, then delete; do not copy implementations
   into a differently named consumer module.
7. **Add import-boundary enforcement after façades exist.** Restrict low-level
   package imports to host adapters and explicitly allow stable pure presentation
   helpers. Avoid enforcing a speculative boundary before the replacement API is
   in place.
8. **Design cross-author features as explicit product actions.** Slicing,
   paragraph restoration, and rejected-view insertion affect reviewer intent and
   need clear tool schemas, authorization/confirmation policy, and lifecycle tests.
   A fail-closed error should not silently activate them.
9. **Use receipts for observability.** Preserve operation IDs, durable revision
   IDs, dropped anchors, validation summaries, and rollback dispositions so UI,
   MCP clients, and corrective agents can explain outcomes without inspecting
   internal OOXML algorithms.
10. **Keep the completed deletion guard.** The 571-line removal of
    `integration.js` and `word-route-change.js` showed that checking existing
    dependency injection and reachability first can remove layers without moving
    their behavior elsewhere.
