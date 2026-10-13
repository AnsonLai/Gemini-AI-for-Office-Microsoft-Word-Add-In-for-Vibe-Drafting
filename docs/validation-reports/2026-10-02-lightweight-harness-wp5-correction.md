# WP5 correction: preserve ordinary commenting and revision-aware editing

**Date:** 2026-10-02 (America/Vancouver).
**Status:** Consumer repairs verified in actual Office.js and independent desktop Word; an untracked surgical-insertion library defect keeps WP5 open.
**Library:** Exact `@ansonlai/docx-redline-js@0.8.4`, unchanged in both consumers.

The owner rejected two blanket refusals introduced in `2aa6aca`. Adding a
comment to a document that already has comments and editing a paragraph that
contains pending revisions are ordinary professional workflows. The previous
[WP5 report](2026-10-02-lightweight-harness-wp5.md) records valid test results for
the committed behavior, but refusal tests did not establish preservation of
those capabilities. Marking WP5 complete and describing the restrictions as
authorized scope reductions was incorrect.

## Comments: repair transport rather than reduce scope

The standalone library returns new comment definitions as a merge delta. The
old whole-body adapter installed that delta as the complete comments part,
discarding existing definitions. Its subsequent refusal masked the defect.

The correction composes `ensureCommentsArtifactsInZip` over a small in-memory
parts object with the exported `reconcileCommentSiblingParts` helper. The
library merges comment definitions and reconciles IDs/extensible metadata.
The existing adapter still reads and inserts Flat OPC; no binary Word facade,
private merge implementation or native comment fallback is introduced.
Existing thread metadata, replies, resolved state and durable IDs are retained.
Fresh host verification now also compares reply relationships and resolved
state; losing either reports an unverified write requiring inspection, with no
automatic replay.
It also compares public revision segments and reviewer identities. Revision IDs
may be renumbered by Word; adjacent spans may split or coalesce while preserving
their type, author, text and accepted/rejected offsets. Flattened or reattributed
carriers fail verification even when accepted text is unchanged.
Merge/replace modes remain explicit, and invalid merges fail before checkpoint
or insertion.

Capability gates include a first new comment on a comment-bearing source, a
second addition after fresh inspection, a supported combined rewrite/comment
batch, and checkpoint restoration after additions. Expected counts and original
thread identities are independent of the prepared result.

## Revisions: use supported action and policy semantics

The broad `UNTRACKED_EDIT_ON_REVISED_TARGET` guard is removed. For wording edits
on targets with foreign or unattributed revisions, the adapter uses the
documented `slice-cross-author` policy per operation, which preserves foreign
revision carriers during supported edits. Own-only targets keep the documented
same-author baseline merge; they are not unnecessarily forced into surgical mode.
Structural/list actions keep their existing policy and transformation mode.
The tracked-edit setting still controls whether
the new change receives new markup. Localized replacements and formatting do
not accept unrelated revisions. Changes within reviewed text use the library's
surgical slice semantics, which can alter the requested span while retaining
unaffected review marks; no blanket identical rejection baseline is promised
for an intentional direct change to that span.

The earlier probe used a whole rewrite of the same author's revision. The
library documents that case as merging against the author's original baseline;
with tracking disabled the result is clean text. That established behavior was
mistakenly generalized into a prohibition on every revised paragraph. The
correction retains same-author merging, foreign review preservation and specific
library refusals for unsafe structures. No acceptance or rejection policy is
silently selected.

### Required upstream correction, not a scope retirement

The [upstream issue and portable public reproducer](../library-issues/completed/2026-10-02-generate-redlines-false-cross-author-insertion.md)
record the exact dependency and expected post-fix behavior.

The expanded actual Word gate found a separate published 0.8.4 engine defect:
adding `applicable ` within an ordinary run through surgical mode creates a new
tracked insertion despite `generateRedlines:false`. The retained source has
another reviewer's pending paragraph-mark/text insertion. The Office.js output
has two native revisions rather than the source's one, and independent Word
Reject All removes the requested clean wording. Per-operation tracking settings
and both public `replace`/`redline` mappings reproduce the defect. This is not
harmless Word revision-ID renumbering.

The required insertion capability remains in scope and its original expectation
is retained. The consumer now checks public mutation receipts against prepared
revision segments: newly created own tracked text while tracking is disabled
produces `TRACKING_POLICY_MISMATCH` before checkpoint or content insertion.
Preserved/cloned foreign review carriers are allowed. This protects the document
without accepting history, enabling tracking, adding native fallback or bypassing
the engine defect with an unrelated affinity setting. A passing refusal test is
safety evidence, not evidence that the requested insertion works. WP5 cannot be
marked complete until a published library fix passes the original host gate.

The original failed package, checkpoint and COM evidence remain under
`.cache/reliability/wp5-correction/pre-untracked-insertion-failure/`; the
independent one-case report is
`.cache/reliability/wp5-correction/literal-revision-probe/word-report.json`.

## Verification

The repaired production source is frozen. Earlier passing refusal controls
are not reused as capability evidence. Live-provider and Word Online release
gates remain pending.

- `npm test`: 46 suites pass, zero fail, three excluded entrypoints. The
  historical package-version skip remains explicit; atomic consumer checks pass.
- Production and replay development builds pass. The deferred vendor bundle
  still produces webpack size warnings; Office telemetry access warnings are
  nonfatal. Logs are under `.cache/reliability/wp5-correction/`.
- Final actual Word NDA replay: 16/16 original prompts and 49/49 checks, zero
  regressions, replay failures or model failures. There are 38 mocked chat calls
  and zero auxiliary/other calls. Eleven batches produce ten verified content
  writes and one no-op. All author checks pass. Artifacts:
  `.cache/reliability/wp5-correction/final-golden-policy-correction/`.
- Actual Office.js on Word **16.0.20430.20092 / PC**: 13 supported edit cases
  pass, with 19 checks including one stale-handle refusal, one verified no-op,
  one invalid-handle refusal and three verified IndexedDB checkpoint restores.
  The sequential comment case uses two atomic writes/two checkpoints: three
  source comments become four, then five after a combined tracked rewrite/new
  comment. Restoring the latest checkpoint recovers four comments and the
  pre-rewrite wording. Original replies, resolved state and durable IDs remain.
  Artifacts: `.cache/reliability/wp5-correction/final-direct-supported13-final/`.
- A separate actual Office.js safety control returns `TRACKING_POLICY_MISMATCH`
  with zero writes/checkpoints and unchanged canonical source, revisions and
  comments. Its fixture is the same Word-exported pending-review source; native
  tracking is restored. This is not a capability pass for the excluded insertion.
  Artifacts: `.cache/reliability/wp5-correction/final-direct-blocked-control-final-v2/`.
- Independent desktop Word **16.0.20430** passes all **64 applicable checks**
  across the 13 supported edits and three checkpoint restorations, with zero
  failures and 16 explicitly not-applicable engine-reference checks. These use
  actual Office.js exports in source, tracked, Accept All and Reject All views;
  opening requires no repair. Rendering is skipped. The separate safety refusal
  is excluded from these capability totals. Artifacts:
  `.cache/reliability/wp5-correction/final-direct-supported13-oracle-final/word-report.json`.
- Fixed production measurement: 4,402 lines across 20 files, including all
  replacement modules. WP0 was 8,632 lines across 24 files: a decrease of 4,230
  lines (49.0%). This repair adds 47 lines compared with the regressed WP5 commit.

The full default manifest retains **14 applied-edit expectations**. Only the
explicit 13-case diagnostic subset passes; the original untracked insertion
gate remains unresolved. A separate safety-manifest clone opts into the refusal
control, and a generator assertion prevents changing the default capability into
an expected refusal. The passing subset and independent oracle do not close the
remaining required insertion gate.

The first correction host run refused its initial list insertion before any
checkpoint or write. A public-only reproduction on that exported source showed
that applying `slice-cross-author` globally forced the complete-list operation
into surgical text mode and produced `PATCH_ROUNDTRIP_MISMATCH`. Per-operation
selection repairs that mistake while retaining the certified list recipe.
An offline regression verifies actual accepted two-item list structure, and the
final rerun covers all four retained list shapes. The failed run remains at
`.cache/reliability/wp5-correction/pre-per-op-policy-failure/`.

The next host run passed all five list writes (four retained list shapes and
the tracking-restoration variant), then stopped on a collector assertion-helper
error before the comment edit. The harness incorrectly called Node-style
`assert.equal` on its browser assertion function. Local equality helpers repair
that test harness error; production source is unchanged. The partial run remains
at `.cache/reliability/wp5-correction/pre-assert-helper-failure/` and is not counted
as final capability evidence.

The first sequential-comment run used numeric `w:id` as a durable lookup key,
then reported the wrong comment after Word reordered IDs. Public preparation and
Word export both retain the first addition's durable ID and expand its anchor
from `local law` to `applicable local law` during the targeted rewrite. The
corrected assertion uses durable identity and that explicit intended anchor;
all three original comments retain strict content, thread/state and anchor
checks. This is an assertion correction based on independently specified edit
semantics, not a dropped preservation requirement. The partial run remains at
`.cache/reliability/wp5-correction/final-direct-supported13/`.

The initial separate safety control stopped on a proven no-write stale fixture
preparation. Reusing the passing run's Word-exported prepared source isolates
the policy control from that setup. A second attempt incorrectly compared token
objects by reference; structural equality repaired the harness assertion. Both
failed attempts remain in the `final-direct-blocked-control` and
`final-direct-blocked-control-final` artifact folders. The final control above
passes without weakening source equality.

The first independent oracle had 63 passes and one source-format assertion
failure: the untracked-format fixture starts without underline, while both
Accept All and Reject All retain the requested untracked underline. The oracle
previously assumed source formatting equals rejected formatting. An explicit
`underlineSource: false` expectation separates these views while retaining all
format checks; direct Office.js readback also confirms None to Single. The
earlier report remains at
`.cache/reliability/wp5-correction/final-direct-supported13-oracle/word-report.json`.
