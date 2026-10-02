# Lightweight document agent WP0 baseline (2026-10-02)

This is the implementation baseline for the approved [library-first reliability plan](../plans/2026-10-02-library-first-lightweight-agent-reliability.md). It records current tracked production code, a fresh offline run, a fresh deterministic desktop Word replay, and the boundaries that constrain the retained operation set. The plan owns the product decisions; this report maps them to current code and evidence. The baseline source was Git HEAD 23cbb33c3ed5bdeb6e8f28bc64317cc2647bbf03, before the cutover edits.

## Version and source-size baseline

The add-in pins **@ansonlai/docx-redline-js 0.8.4** in package.json and package-lock.json; the MCP consumer uses the same exact release. The deterministic replay used the recorded model label **gemini-flash-latest**, with requests intercepted and answered locally by the replay harness. It made no provider calls.

Production size is **8,632 physical lines across 24 tracked JavaScript files**:

| Included area | Files | Lines |
| --- | ---: | ---: |
| src/taskpane/taskpane.js | 1 | 2,658 |
| src/taskpane/modules/commands/ | 9 | 3,452 |
| src/taskpane/modules/docx-redline-js-integration/ | 7 | 1,683 |
| src/taskpane/modules/chat/ | 6 | 804 |
| src/commands/commands.js | 1 | 35 |
| **Total** | **24** | **8,632** |

Method: count every tracked .js file in the listed roots, recursively within the three module directories. A physical line is each CRLF, LF, or CR terminator plus a final unterminated line when present. Exclude tests, generated bundles, assets, and files not tracked at the WP0 source snapshot. The untracked, unwired function-call-recovery.js and tool-registry.js drafts were not production code at this snapshot and are excluded. Use these same roots and counting rule for the eventual net change; include replacement modules once they are wired into production.

## Current tool plumbing

The active tool schema is one function_declarations array in [taskpane.js](../../src/taskpane/taskpane.js) with 11 entries, in order:

apply_redlines, insert_comment, highlight_text, perform_research, navigate_to_section, edit_list, insert_list_item, edit_table, edit_section, convert_headers_to_list, get_selection_stats.

The declarations, executor dispatch, malformed-call recovery whitelist, and mutability classifier are separate lists/branches in sendChatMessage:

| Surface | Current count and behavior |
| --- | --- |
| Declarations | 11 tool schemas. |
| Dispatch | 11 explicit serial if/else-if executor branches. Multiple calls in one model response are processed in order; they are not rejected as a group. |
| Malformed-call recovery | A separate 11-name allowlist. Special argument parsing exists for edit_list and convert_headers_to_list; a regex fallback reconstructs raw apply_redlines instructions. |
| Mutation classification | 8 names: apply_redlines, insert_comment, highlight_text, edit_list, insert_list_item, edit_table, edit_section, and convert_headers_to_list. Research, navigation, and statistics are not classified as mutations. |
| Progress labels | Six explicit labels (apply_redlines, insert_comment, highlight_text, perform_research, navigate_to_section, get_selection_stats); other tools show the generic “Working…” label. |

The provider loop allows six model turns, stops after two no-progress mutation cycles, and permits one stale-context refresh. The existing shared Gemini client handles bounded transport retries. These defenses are already implemented and should remain part of the cutover.

## Nested model calls and editing routes

The chat model selects a tool, and some executors ask Gemini to interpret the same instruction again:

| Executor | Secondary model work |
| --- | --- |
| executeRedline / apply_redlines | callGeminiForDiffs drafts concrete changes; one corrective generation is allowed after a no-progress result. |
| executeComment / insert_comment | callGeminiForJSON generates comment targets and text. |
| executeHighlight / highlight_text | callGeminiForJSON generates highlight targets. |
| executeNavigate / navigate_to_section | callGeminiForJSON chooses a paragraph before Word selects it. |
| executeResearch / perform_research | A separate Gemini request enables google_search. |

The fresh replay recorded **43 local model requests**: 32 chat requests and 11 auxiliary JSON/diff requests, all under the gemini-flash-latest label. The 32 chat requests comprised 16 function calls and 16 post-tool chat responses. The 11 auxiliary requests were nine redline diff generations, one highlight-target generation, and one comment-target generation. No corrective generation, research, or navigation call occurred; there were no replay errors or desynchronizations.

The 16 top-level function calls were:

| Tool | Calls |
| --- | ---: |
| apply_redlines | 9 |
| insert_list_item | 2 |
| edit_list | 1 |
| edit_table | 1 |
| convert_headers_to_list | 1 |
| highlight_text | 1 |
| insert_comment | 1 |

The taskpane and tool module already use the canonical library for much of the content transformation. apply_redlines prepares a source-based library batch and commits OOXML through the Word adapter. edit_list uses applyRedlineToOxml and one range insertion. Comment and highlight suggestions use the library bridge, but are currently transformed and committed one suggestion at a time. The adapter's getOoxml/insertOoxml calls are Office.js host transport and remain in scope.

Direct Word editing algorithms remain in [agentic-tools.js](../../src/taskpane/modules/commands/agentic-tools.js):

- insert_list_item has a canonical batch branch for a verified narrow set of existing bullet/decimal list shapes (level 0 or 1, known numbering id, exact source text). Other shapes fall through to paragraph insertion, hand-built list OOXML, or native listItem.level changes.
- convert_headers_to_list uses insertText, startNewList, list-level numbering, and attachToList.
- edit_table directly replaces cell text and exposes Table.addRows and TableRow.delete. Markdown-formatted cell updates use the shared library bridge; plain cell updates and structural table actions use direct Word APIs.
- edit_section scans Word list levels to infer a section boundary, then directly inserts text, deletes paragraphs, and inserts paragraphs.
- navigate_to_section calls Word's select() after the secondary model call. This is host navigation, not a content edit.

The taskpane also uses Word APIs for source/context and metadata reads, change-tracking state, selection, confirmation, and checkpoint restoration. Those host responsibilities are distinct from the direct editing algorithms above. The separate [commands entrypoint](../../src/commands/commands.js) posts an Office notification; it is not an additional document-edit route. [consumer-core.js](../../src/taskpane/modules/docx-redline-js-integration/consumer-core.js) remains 633 lines of consumer-side package/relationship handling and canonical preparation around the upstream engine.

## Approved operation matrix and evidence

The treatment column follows the owner-approved plan. Current successful behavior is evidence about the existing route; where that route is native Word or lacks required independent fidelity evidence, it does not establish a retained library capability.

| Operation | Approved treatment | Current evidence and boundary |
| --- | --- | --- |
| Text correction, localized replacement, clause rewrite, paragraph insertion/deletion | Retain as typed, source-bound edits in one atomic library batch. | Current apply_redlines has library preparation and source guards. The fresh replay passed all applicable content checks for steps 1, 8, 11, and 16. Existing source-baseline and batch tests passed in the fresh offline suite. |
| New comments | Retain as batch intents. | The library bridge and batch tests cover comment parts; fresh step 14 passed. Current tool makes a secondary AI call and writes one suggestion at a time. Text-plus-comment on the same provision needs verified post-edit anchor mapping before it is represented as one batch. |
| Basic inline emphasis and removal (bold, italic, underline, strikethrough) | Retain only this narrow formatting allowlist. | Existing Markdown formatting path and 0.8.4 occurrence-targeting coverage; fresh steps 2, 5, 10, and 15 passed. This does not authorize a general formatting surface. |
| Everyday list text changes and insertion | Retain only the list shapes, levels, and styles with canonical source-bound evidence; refuse the rest. | edit_list uses the library route. insert_list_item uses a canonical branch only for its checked shape subset and otherwise falls back to native editing today. Fresh step 3 (edit_list) and step 7 (simple insertion) passed; step 9's level-3 insertion also passed content but used the native fallback, so it is not canonical evidence. The offline list suites passed with the numbering-defect cases reported as skips. |
| Conversion of plain headers to numbered lists; numbering-style or deep-list changes | Unsupported in the new path until upstream capability and numbering fidelity gates pass. | The [active list follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md) and [issue index](../library-issues/README.md) document unsupported plain-to-list conversion and a 0.8.4 regression where separate non-adjacent list operations restart numbering. Fresh step 12 passed content using today's native route and failed author attribution for 27 new revisions; it does not validate a canonical replacement. |
| Existing table-cell text changes | Retain only where the library route preserves the inspected table reliably. | The approved starting scope is cell text, not structural table editing. Today plain update_cell writes directly through Word; Markdown formatting uses the shared bridge. The fresh replay did not cover a cell-text-only action, so the future library mapping still needs direct evidence. |
| Table row add/delete, whole-table replacement, or paragraph-to-table conversion | Gate until independent fidelity evidence supports the specific structural action. | Fresh step 6's native row addition passed its two content checks but added five revisions by Reviewer. The [0.8.4 report](2026-10-02-docx-redline-v084.md) records an extra empty paragraph after Word Reject All for a final tracked table append. Fresh step 4's accepted table result is useful replay coverage, but it is not an independent Word Accept/Reject oracle. |
| Inferred edit_section rewrites | Retire this inferred Word list-scanning route. | The approved replacement is an explicitly inspected range with typed text edits. Existing edit_section directly scans list levels and edits paragraphs; the golden replay did not exercise it. |
| Highlighting | Remove the standalone tool; keep a narrow intent only if existing use warrants it. | Fresh step 13's highlight content check passed through the current tool, which used one auxiliary model call. This alone does not justify keeping a separate tool or broad color palette. |
| Navigation | Optional third, handle-based host tool; no secondary model call. | Existing tool uses an auxiliary Gemini call followed by Word selection. The golden replay did not exercise it. |
| Selection statistics | Compute through inspection or a simple UI command, not a separate agent executor. | Current executor is local text-statistics code; it was not used in the replay. |
| Research | Optional explicit-use workflow, isolated from ordinary edits and its mutation budget. | Current separate Google Search model route was not used in the replay. Usage evidence is unavailable in this harness. |

The documented weak-target soft-break deletion remains a required guardrail: the new path must keep strong source-bound targets and refuse when the snapshot cannot be aligned. Do not use native fallback to conceal an unsupported library operation.

## Fresh baseline runs

### Offline consumer suite

Command: npm test

Result: **57 suites passed, 0 failed; 4 entrypoints excluded.** Suite-internal skips were the two non-contiguous header-numbering cases linked to the documented 0.8.4 defect and version-specific v0.5.4 behavior (the installed package is 0.8.4). Excluded entrypoints were live Gemini evals, observational performance, XML-provider setup, and the desktop Word fixture entrypoint. Desktop Word was exercised by the separate golden run below; that run does not replace the excluded fixture lane. No provider credentials or paid provider calls were used.

### Deterministic desktop Word golden replay

The baseline bundle was built before the implementation cutover and served from a dedicated output directory:

    npx webpack --mode development --env GOLDEN_SCENARIO=1 --output-path .cache/reliability/wp0-baseline/dist

Webpack 5.102.1 compiled successfully (exit 0). The build emitted a blocked optional Application Insights connection message; it did not affect compilation. The replay then ran with:

    node scripts/run-golden-scenario.mjs --launch --skip-build --dist-dir .cache/reliability/wp0-baseline/dist --artifacts-dir .cache/reliability/wp0-baseline/golden

The first sandboxed launch was denied the runner's scoped HKCU add-in-registration write before Word started. The same requested runner succeeded with elevated execution. It registered the runner's separate golden identity, served on port 3001, and launched Word /x as PID 61020. The runner targets only its own disposable Word instance and closes/unregisters it on completion; no existing Word session was targeted. The Word host metadata in the fresh report says Word / PC and does not include an Office build number. The executable launched by the runner reports FileVersion 16.0.20430.20092; the replay JSON itself does not attach a build to its host record.

The replay completed **16/16 steps** with **29/29 content checks passing**. Three of the 16 author checks failed, leaving 42/45 total checks passing:

| Step | New revision author evidence |
| --- | --- |
| 6, signature date row | 5 revisions by Reviewer, not Golden Scenario |
| 9, archival subbullet | 3 revisions by Reviewer, not Golden Scenario |
| 12, headers to list | 27 revisions by Reviewer, not Golden Scenario |

There were zero content failures, cross-step regressions, replay desynchronizations, or known-issue classifications. The runner's exit code is 1 because of those three author checks. This reproduces the author-attribution finding in the existing 0.8.4 validation report.

Artifacts are in .cache/reliability/wp0-baseline/: the frozen dist/ bundle and golden report.md, report.json, 16 per-step DOCX/JSON exports, and model-log.jsonl.

### Measurements not available from this harness

The golden runner exports and scores a DOCX after each prompt and records model requests, console/transcript data, and revision authors. It does not instrument Office.js proxy reads or host write attempts/confirmations. Therefore runtime source-read counts, host-write counts, and per-tool structured mutation receipts are **unavailable**; they must not be inferred from model requests, elapsed time, or passing content checks. The available outcome evidence is the per-step exported-document score and the author delta above.
