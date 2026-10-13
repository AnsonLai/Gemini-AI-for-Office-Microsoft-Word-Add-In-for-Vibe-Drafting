# Lightweight document agent WP1 validation (2026-10-02)

WP1 is complete: **61 offline suites pass**, development/production builds pass, and the completed desktop Word replay passes **29/29 content checks** with the same three native-route author failures as WP0. This report compares the extracted loop and registry against the frozen [WP0 baseline](2026-10-02-lightweight-harness-wp0-baseline.md) and the approved [reliability plan](../plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md). All executor algorithm bodies/signatures, 11 declarations, the system instruction, exact library pins, and Word transport are preserved; executor dependencies now belong to each request.

## Run

The runner used the already-built WP1 bundle and preserved the existing 16 prompts and replay protocol:

    node scripts/run-golden-scenario.mjs --launch --skip-build --dist-dir .cache/reliability/wp1/completed-dist --artifacts-dir .cache/reliability/wp1/golden-completed

The runner served the dedicated golden add-in on port 3001 and launched a separate Word /x instance as PID 21364. No existing Word session was targeted. The host report identifies Word on PC; the launched WINWORD.EXE reports FileVersion 16.0.20430.20092.

This was local replay mode. All 43 model requests were intercepted by the scenario harness and answered from the script; there were no live provider calls or paid usage. The runner completed and closed its disposable Word instance.

## Result and WP0 comparison

The replay completed **16/16 steps**. All **29/29 content checks passed**. Three of the 16 author checks failed, for **42/45 total checks passing**. The runner returned exit code 1 because of those three author checks.

| Step | Result |
| --- | --- |
| 6, signature date row | 5 new revisions by Reviewer rather than Golden Scenario |
| 9, archival subbullet | 3 new revisions by Reviewer rather than Golden Scenario |
| 12, headers to list | 27 new revisions by Reviewer rather than Golden Scenario |

There were zero content failures, cross-step regressions, known-issue classifications, or replay desynchronizations. The three author failures match WP0 exactly.

| Measurement | WP0 | WP1 | Change |
| --- | ---: | ---: | --- |
| Steps completed | 16/16 | 16/16 | None |
| Content checks passed | 29/29 | 29/29 | None |
| Total checks | 45 | 45 | None |
| Failed checks | 3 author checks | 3 author checks | None |
| Local model requests | 43 | 43 | None |
| Chat requests | 32 | 32 | None |
| Auxiliary/nested requests | 11 | 11 | None |
| Replay desynchronizations | 0 | 0 | None |
| Cross-step regressions | 0 | 0 | None |

The WP1 model log again shows 16 chat function calls followed by 16 chat responses and 11 auxiliary requests. The top-level distribution is unchanged: apply_redlines ×9, insert_list_item ×2, and edit_list, edit_table, convert_headers_to_list, highlight_text, and insert_comment ×1 each. The auxiliary requests are nine redline diff generations, one highlight-target generation, and one comment-target generation. Every request used the gemini-flash-latest label in replay mode.

This run exercises the extracted production loop through the real taskpane UI while preserving the old tool protocol. It confirms unchanged scenario behavior and calls; it does not validate the later inspect_document/apply_document_edits proposal, which was not part of WP1.

The runner reports exported-document outcomes and author deltas, but does not instrument Office.js proxy reads or host write attempts/confirmations. Runtime read counts, write counts, and structured mutation receipts remain unavailable. The Word host metadata does not contain a build value; the executable file version above is from the WINWORD.EXE path launched by the runner.

Earlier WP1 replays used `dist`/PID 31460 and `final-dist`/PID 59936, before completing request-scoped executor dependencies. They produced the same scores and request counts. The completed run above validates the final production source, including that dependency boundary, with zero model-log errors.

## Artifacts

Completed bundle/replay: `.cache/reliability/wp1/completed-dist` and `.cache/reliability/wp1/golden-completed`. Earlier iteration artifacts remain under `dist`/`golden` and `final-dist`/`golden-final`. Each golden directory contains report.md, report.json, model-log.jsonl, and per-step JSON and DOCX exports.

## Offline and build verification

- `npm test`: **61 suites passed, 0 failed; 4 entrypoints excluded**. This includes four new loop/registry/recovery/request-scope suites and the updated navigation/baseline, table, Gemini, and migrated-tool boundary tests. Existing suite-internal numbering/version skips remain reported separately; exclusions are live Gemini evals, observational performance, XML setup, and the separate desktop fixture entrypoint. The golden replay exercises desktop Word separately, without replacing the excluded fixture lane.
- `npm run build:dev` and `npm run build`: both exit 0. Production reports one webpack asset-size warning for the 530 KiB deferred chunk. The completed golden bundle also builds successfully with `node node_modules/webpack/bin/webpack.js --mode development --env GOLDEN_SCENARIO=1 --output-path .cache/reliability/wp1/completed-dist`.
- `git diff --check`: passes. No dependency/lockfile changes. No live Gemini or Word Online run; broader Accept All/Reject All certification is unchanged from existing reports.

Completed logs: `.cache/reliability/wp1/offline-agent.log`, `build-dev-final.log`, `build-prod-final.log`, and `build-golden-completed.log`. Generated artifacts are ignored and are not test prerequisites on another checkout.

Deterministic coverage includes serial queued calls with one snapshot per exchange; refreshed context/baseline before the next model turn; one stale-refusal replan without resetting the failed-mutation budget; two failed mutation cycles; malformed-call recovery from registry names; unknown/invalid calls; complete function-response pairing; provider transport retries without tool replay; cancellation before/within checkpoints and while a write settles; timeout; failed refresh; and each uncertain host outcome. Taskpane boundary tests execute the actual request-identity guard blocks to ensure refresh/supersession cannot restore cleared history, write obsolete UI, or unlock the next request. The request-scope regression pauses an old redline during a host read, aborts it, completes a newer instance, and resumes the old call: it still uses its aborted signal, refuses before provider/mutation work, and cannot reuse the newer credentials.

## Net production size

Same physical-line method and roots as WP0, including all three new wired production modules and the executor dependency factory:

| Included area | WP0 lines | WP1 lines | Change |
| --- | ---: | ---: | ---: |
| `src/taskpane/taskpane.js` | 2,658 | 1,657 | −1,001 |
| `src/taskpane/modules/commands/` | 3,452 | 3,469 | +17 |
| `src/taskpane/modules/docx-redline-js-integration/` | 1,683 | 1,683 | 0 |
| `src/taskpane/modules/chat/` | 804 | 1,789 | +985 |
| `src/commands/commands.js` | 35 | 35 | 0 |
| **Total** | **8,632** | **8,633** | **+1** |

Production files increase from 24 to 27. The new registry is 604 lines (including unchanged declarations), recovery 168, and loop 213. Net physical size is essentially flat, **one line larger**; moving 1,001 lines out of the taskpane is not equivalent to deleting them. WP1 removes competing declaration/dispatch/mutability/recovery lists and creates testable request boundaries while preserving the current protocol. Most tool/executor deletion and the required net reduction belong to WP3–WP5.

## Remaining scope

All 11 tools and their current nested AI/native routes remain. The standalone highlight palette, structural list/table routes, and nonblocking checkpoints are not newly certified or retired by this extraction. The three author failures and documented upstream list/table/soft-break constraints stay open. Source-bound handles, direct atomic typed batches, stricter scope refusal, checkpoint blocking, and native-route removal remain the later approved work packages.
