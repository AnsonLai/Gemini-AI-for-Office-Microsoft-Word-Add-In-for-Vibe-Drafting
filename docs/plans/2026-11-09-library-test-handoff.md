# Library-only tests: hand-off to docx-redline-js

**Created:** 2026-11-09. **Status:** ported and deleted. The library carries the suites as of its commit `6d6bc5f`, and the eleven whole-file plugin copies and `tests/sample_doc/` were deleted here on 2026-11-09 (`npm test`: 42 suites pass). Still in the plugin, owner to decide: the eight library cases in the `test:word` lane and the sections of `tests/ooxml_formatting_visual_tests.mjs` that record them (with `header-footer.docx` and `multi-paragraph-thread.docx`), (`tests/phase4/*` was dropped without a port, and the 0.5.4 compat suite was replaced by `tests/atomic_batch_policy_tests.mjs`, both on 2026-11-09). The tables below are the original hand-off list.

**Goal:** the plugin leans on `@ansonlai/docx-redline-js` (pinned 0.11.1) for every document transformation, so tests that only exercise the library belong in the library repo (`..\..\Docx Redline JS`). The plugin keeps tests of its own code: the chat loop and tools, the intents compiler, the Flat OPC bridge, storage and UI.

**How the list was made:** a read of every test entrypoint and the library's `tests/` and `scripts/`. "Covered upstream" means the distinctive identifiers, error codes or fixtures appear in the named upstream file. It was not compared assertion by assertion, so diff before deleting.

## What the plugin uses from the library

Used in `src/`: `applyOperationsToDocumentXml`, `inspectDocumentParts`, `mergeCommentParts`, `mergeNumberingXmlBySchemaOrder`, `normalizeBodySectionOrderStandalone`, `computeDocumentPartsRevisionToken`, `areRevisionTokensEqual`, `compareDocumentInspections`, `parseOoxmlSafe`, `serializeOoxml`, `getDefaultAuthor`, `setPlatform`, `NS_W`.

Not used in `src/`: `openDocx`, `resolveRevisions`, `validateDocxPackage`, `applyRedlineToOxml`, `configureXmlProvider`. Tests of these guard nothing the plugin ships.

Plugin tests deep-import `@ansonlai/docx-redline-js/services/package-builder.js` (`buildDocumentFragmentPackage`) to build fixtures, so the library must keep exporting it.

## Move upstream (needs porting)

| Plugin file | Lines | What it asserts | Needs | Upstream today |
| --- | --- | --- | --- | --- |
| `tests/addin/integration_tests.mjs` | 627 | Ten NDA edits on `tests/sample_doc`, each checked paragraph by paragraph after Accept All and Reject All; four `generateRedlines:false` cases | `tests/sample_doc/**` (18 files, byte-identical upstream) | Operations are covered separately (`list_tests`, `table_tests`, `formatting_tests`, `standalone_smoke`); no whole-document view. Table with redlines off not found |
| `tests/ooxml_formatting_visual_tests.mjs` L74-120 | ~47 | `formatting-layout-localized`: bold/italic inheritance, both `sectPr` layouts, tab stops and hyperlinks kept, styles and rels byte-identical | `tests/fixtures/wp6/formatting-layout.xml` | No clear match |
| `tests/phase4/golden-guardrail.mjs` + `phase4-golden-baseline.json` | 313 + 63 | Hash snapshots of nine engine outputs | Deep imports of library internals, fixed `Date` | None. Port only if byte snapshots are wanted |
| `tests/phase4/perf-harness.mjs` + `phase4-perf-baseline.json` | 243 + 32 | Timing thresholds for engine internals (not in `npm test`) | | None by these workloads. Port or drop |
| `tests/reconciliation_author_platform_defaults_tests.mjs` | 84 | `setDefaultAuthor` and `setPlatform` reach revision metadata, comments and the pipeline | | No upstream test names these |
| `tests/targeting_helpers_extraction_tests.mjs` | 75 | `resolveParagraphRangeByRefs` and table-targeting helpers | | `resolveParagraphRangeByRefs` untested upstream |
| `tests/include_numbering_behavior.mjs` | 29 | `buildDocumentFragmentPackage` adds the numbering relationship only with `includeNumbering` | | No direct test |

## Delete from the plugin once upstream coverage is confirmed

| Plugin file | Lines | Upstream file(s) that appear to cover it |
| --- | --- | --- |
| `tests/docx_redline_v081_compat_tests.mjs` | 203 | `docx_package_facade_tests`, `comment_resolve_and_delete_tests`, `comment_reply_tests`, `header_footer_tests`, `package_content_type_contract_tests` |
| `tests/docx_redline_v083_compat_tests.mjs` | 140 | `list_reject_fidelity_tests`, `facade_list_numbering_tests` |
| `tests/docx_redline_v040_compat_tests.mjs` | 74 | `error_contract_tests`, `input_mutation_policy_tests` (the literal `$1,000` case was not found) |
| `tests/numbering_helpers_extraction_tests.mjs` | 65 | `numbering_allocation_tests`, `phase3_numbering_routing_tests` |
| `tests/edit_list_numbering_tests.mjs` | 39 | `numbering_allocation_tests` (not verified) |
| `tests/dom_fallback_smoke.mjs` | 27 | `list_parser_diagnostics_tests` (partial) |
| `tests/agentic_list_generation_tests.mjs` | 20 | `phase3_list_structural_fallback_tests`, `list_tests` (not verified) |
| `tests/ooxml_formatting_visual_tests.mjs` L122-156 | ~35 | Word thread, sibling reply, resolve/reopen/delete: `comment_thread_word_fixture_tests`, `comment_resolve_and_delete_tests`, `comment_parts_merge_tests` |
| same file, L237-270 | ~34 | First-page footer edit: `header_footer_tests` L185-197 |
| same file, L301-341 | ~41 | Hyperlink-boundary and structural-lifecycle regressions: `reconstruction_boundary_fidelity_tests` |
| `tests/docx_redline_v054_compat_tests.mjs`, all but L56-81 | ~281 | Gated on version 0.5.4, so it only prints SKIP today. Behaviours live in `agent_operation_contract_tests`, `docx_package_transaction_edge_tests`, `error_recovery_contract_tests`, `existing_revision_mutation_policy_tests`, `paragraph_identity_tests` |

Fixtures that leave with them: `tests/sample_doc/**`, `tests/fixtures/wp6/header-footer.docx`, `tests/fixtures/wp6/multi-paragraph-thread.docx` (all byte-identical upstream), and the generated `tests/fixtures/phase4-*-latest.json`.

## The Word oracle

`npm run test:word` (default fixtures) runs ten cases recorded by `tests/ooxml_formatting_visual_tests.mjs`. Eight are library fixtures and facade calls: `formatting-layout-localized`, `word-multi-paragraph-reply`, `word-optional-sibling-reply`, `word-authored-footer-localized`, the four `regression-*` cases, plus the `control-bad-commentsExtended-type` control. Two go through the plugin bridge: `word-addin-sibling-reply` and `word-addin-plain-replacement`.

The library's own oracle (`scripts/word-com-differential.ps1`, driven by `word-com-suite.ps1`) compares Accept/Reject text and has a `headers` scope, but has no checks for hyperlinks, per-section header or footer text by type, page setup sections or tab stops. The branches of `scripts/verify-wp6-word.ps1` driven by `expectedHyperlinks`, `expectedSections`, `expectedTabStop` and `expectedHeaderFooterText` are used only by the eight library cases, so they need porting with those cases or dropping.

One thing learned here on 2026-11-09 that the port should carry: `Document.Revisions.AcceptAll()` and `RejectAll()` cover the body story only. Header and footer revisions must be resolved through each part's `Range.Revisions` (commit `d8fa53e`).

## What stays in the plugin afterwards

- `tests/ooxml_formatting_visual_tests.mjs` shrinks to its bridge sections (L158-235, L272-299 and the export tail), about 200 lines, and the default `test:word` lane runs two cases.
- `tests/docx_redline_v054_compat_tests.mjs` shrinks to the scan that every `applyOperationsToDocumentXml` call passes `atomic: true`.
- The KB-9, KB-14, KB-15 and KB-19 sections of `tests/document_edit_intents_tests.mjs` stay: they pin library fixes through the plugin's compiler.
- Optional: one small test that the library exports listed above exist at the pinned version, replacing the version-named compat suites.

## Plugin-side follow-ups (not part of the hand-off)

Renames suggested by the audit, for suites named after history:

| Now | Suggested |
| --- | --- |
| `document_tools_phase2_regression_tests.mjs` | `document_tools_existing_revisions_and_verification_tests.mjs` |
| `document_inspection_session_regression_tests.mjs` | `document_inspection_handle_freshness_tests.mjs` |
| `phase2_dead_code_static_tests.mjs` | `taskpane_no_body_text_static_tests.mjs` |
| `no_legacy_shared_operation_bridge_tests.mjs` + `addin/migrated_tool_cutover_tests.mjs` | merge into `typed_path_architecture_tests.mjs` |
| `agentic_navigation_dispatch_tests.mjs` | `request_scoped_dispatch_and_checkpoint_ui_tests.mjs` |
| `docx_cross_host_parity_tests.mjs` | `word_bridge_matches_library_facade_tests.mjs` |
| `direct_document_fidelity_manifest_tests.mjs` | `scripts/export-officejs-lane-manifest.mjs` (it is an exporter and rewrites `.cache/` on every `npm test`) |

## Size

About 2,100 lines leave under the two tables above, out of roughly 17,850 lines of tests and scripts.
