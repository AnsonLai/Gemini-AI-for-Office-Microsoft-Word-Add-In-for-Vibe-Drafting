# docx-redline-js 0.9.0 consumer validation (2026-10-12)

The add-in and MCP server now pin exact `@ansonlai/docx-redline-js@0.9.0`
(from 0.8.4) in `package.json` and `mcp/docx-server/package.json`. This report
records the verification of the upgrade, the test and oracle changes it
required, what it closed, and what remains. Earlier dated reports (for example
[v0.8.4](2026-10-02-docx-redline-v084.md)) remain historical records.

Host: Word 16 desktop (2026-10-12). Remaining consumer-side bugs and the Word
JavaScript API inventory live in [`docs/plans/KNOWN_BUGS.md`](../plans/KNOWN_BUGS.md).

## Library issue verification

Each previously open report's portable reproducer was rerun against the
installed package, followed by independent checks.

| Report | Result on 0.9.0 |
| --- | --- |
| [Separate list operations restart numbering](2026-10-17-library-issues-closeout.md) | Fixed. Reproducer prints "NOT REPRODUCED: both headers share one list" (A., B., one `numId`). |
| [Weak-target soft-break deletion](2026-10-17-library-issues-closeout.md) | Fixed. Reproducer: NOT REPRODUCED. |
| [`generateRedlines: false` cross-author insertion](2026-10-17-library-issues-closeout.md) | Fixed. Reject All keeps the clean text. The reproducer's defect assertion now throws, so it exits 1. |
| [Word Reject All after final table append](2026-10-17-library-issues-closeout.md) | Fixed at package level. The public-library package passes source, tracked, Accept All and Reject All in Word using `Revisions.RejectAll()`. The live-insert lane still fails: see KB-1. |
| [Canonical list capabilities](2026-10-17-library-issues-closeout.md) | Released as `list-format` (`convert-plain`, `change-format`, `join`) and `paragraph-format.numbering`. `reproduce-plain-header-conversion.mjs` is now refused (`BATCH_OPERATION_FAILED`, strict list conversion): the reproducer is obsolete because the capability moved to `list-format` `convert-plain`. Consumer adoption is tracked in the [list plan](../plans/2026-09-30-canonical-list-migration-follow-up.md) and KB-4. |

## Test and oracle changes (no `src/` change)

- `tests/numbering_helpers_extraction_tests.mjs` imports
  `mergeNumberingXmlBySchemaOrder` from `services/numbering-allocation.js`, its
  new home in 0.9.0.
- `tests/document_tools_tests.mjs` now asserts that untracked pure insertion
  beside a foreign revision SUCCEEDS. It previously asserted the
  `TRACKING_POLICY_MISMATCH` refusal. That guard stays as a backstop (KB-5).
- `tests/addin/integration_tests.mjs` cases moved from the unsupported
  whole-document `applyRedlineToOxml` to targeted
  `applyOperationsToDocumentXml`, with exact Accept/Reject assertions (KB-2).
- Phase-4 golden guardrail hashes were regenerated for `mixed_insert_delete`,
  whose punctuation diff is now surgical, and `text_to_table_transform`, which
  now has tracked `w:trPr/w:ins` rows and tracked cell paragraph marks.
- **Oracle change:** `scripts/verify-wp6-word.ps1` now uses
  `Document.Revisions.AcceptAll()/RejectAll()`, which is what the ribbon does,
  instead of `Document.AcceptAllRevisions()/RejectAllRevisions()`. The latter
  leave a body-final table append's tracked mark behind, a Word API quirk
  (KB-3). Reports produced before 2026-10-12 used the `Document.*` methods.

## Results

- **Offline:** `npm test` passes 46 of 46 suites, with 3 entrypoints excluded.
  Production and development builds pass.
- **Actual Office.js lane:** `node scripts/run-officejs-validation.mjs --launch`
  gives status passed: all 14 default capability cases, 20 of 20 host checks and
  3 checkpoint restores. This includes the previously blocked untracked pure
  insertion on a revised target,
  `direct-literal-replacement-preserves-comment-threads`, with its unchanged
  applied-edit expectation. Artifacts: `.cache/reliability/v090-officejs`.
- **Independent Word oracle** on those exports: 68 PASS, 0 FAIL. Artifacts:
  `.cache/reliability/v090-oracle`.
- **16-prompt NDA direct replay:** 16/16 steps, 49/49 checks, 0 failed,
  0 regressions, 38 chat model requests, 0 auxiliary. Artifacts:
  `.cache/golden/v090-direct`.

## Environment note

The first host runs hung because the local dev HTTPS certificate expired on
2026-10-03. The owner renewed it, and the runs above were made after renewal.
This was an environment fault, not a product defect.

## What closed

- Lightweight-harness WP5, whose engine blocker (untracked surgical insertion)
  is fixed in 0.9.0: see the
  [plan](../plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md).
- The table-creation incident plan: see the
  [plan](../plans/completed/2026-09-30-table-creation-reliability.md).
- Library issues: every report in `docs/library-issues/` was then in
  `completed/`; the open table was empty (folder removed 2026-10-17, see the [closeout](2026-10-17-library-issues-closeout.md)).

## What remains

See [`docs/plans/KNOWN_BUGS.md`](../plans/KNOWN_BUGS.md):

- KB-1: body-final table append via live insertion leaves an extra paragraph on
  Reject All (table append stays a refused capability).
- KB-2: whole-document `applyRedlineToOxml` is unsupported usage (no caller left).
- KB-3: `Document.RejectAllRevisions()` differs from the ribbon.
- KB-4: list conversion and list-format changes refused despite library support
  (list plan WP1 to WP3).
- KB-5: `TRACKING_POLICY_MISMATCH` guard is a 0.8.4 workaround kept as a backstop.
- KB-6: revision-fidelity refusals reach the model unmapped.

The [canonical list migration follow-up](../plans/2026-09-30-canonical-list-migration-follow-up.md)
stays open.

## Gates not run

- Live fast/slow provider replay.
- Word Online operation checks.

These remain shipping gates. Desktop evidence does not establish either.
