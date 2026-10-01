# Unpublished library list fixes: consumer review

Reviewed on 2026-09-30 against the uncommitted working tree at
`C:/Users/Phara/Desktop/Projects/Docx Redline JS`. Its package version remains
0.8.2. This review does not change either installed package pin or production
routing, and does not certify a published release.

## Findings

| Concern | Review result |
| --- | --- |
| Historical `pPrChange` inspected as current properties | Fixed in local source: inspection and list targeting use direct current paragraph properties. Focused edge tests pass for numbering, style and heading level. |
| Plain insertion leaves an empty paragraph on Reject All | Fixed for the reported case. Add-in planner plus local standalone/merge output passes independent Word. |
| List-range replacement merges source paragraphs on Reject All | Fixed for the reported two-item case, including source numbering/formatting through the add-in path. |
| Empty middle list item | New focused suite passes `Alpha`, empty, `Gamma` accept/reject checks; those checks are offline, not Word automation. |
| All-empty source range | Still loses source empty paragraphs on rejection; separately reproduced and reported. |
| Full canonical header/list-format migration | Still proposed, not implemented. The library plan is useful but does not satisfy this acceptance gate. |

Independent focused runs pass `list_reject_fidelity_tests.mjs`,
`document_inspection_edge_tests.mjs` and
`phase3_list_structural_fallback_tests.mjs`. The upstream agent's 128-suite,
types/isolation results were supplied by the user; this review did not rerun
those full gates. The test helper named `wordParagraphs` models paragraph-mark
semantics; it does not call Word.

The two original add-in diagnostic cases were regenerated with the local library
and the existing consumer planner/standalone numbering merge. Independent Word
passes **12 checks**, including source, tracked, Word Accept/Reject,
library-resolved accept/reject, labels, levels, logical numbering identity and
untouched bold formatting. [Word evidence](2026-09-30-unpublished-list-fixes-word.json).
This uses static exported packages, not a fresh Office.js transport run.

A separate public DOCX facade comparison passes eight checks and fails four
numbering checks. Published 0.8.2 and the local facade produce identical
numbering XML for that edit, so this is pre-existing. See the
[separate facade report](../library-issues/2026-09-30-public-facade-list-numbering.md).

## Header receipts and indexes

The unchanged decimal-header edit reproduces `RECEIPT_RECONCILIATION_FAILED`
with rolled-back receipts on a bare minimal `<w:document>`. The frozen packaged
Word fixture and a minimal body placed in a complete DOCX both return `ok` for
the corresponding unchanged marked-header edit. Keep the bare-XML defect in
library WP0; do not generalize it to every packaged header conversion.

The alleged index discrepancy is expected: library inspection emits
`index = zeroIndex + 1`, and the consumer uses that inspected 1-based index.
JavaScript fixture array positions are 0-based. There is no demonstrated
consumer off-by-one defect; subtracting one from the canonical target would be
wrong. Keep exact text and stable identities as additional guards.

One minor review observation: the new list-targeting direct-child helper matches
`localName` without a Word namespace predicate. Inspection is stricter. Align
those helpers and add a foreign-namespace control before considering inspection
hardening complete; normal Word-authored input is covered by the current tests.

## Resume gates

The agentic plan stays open. Resolve the all-empty-range case, implement and
verify the planned canonical conversion/remapping capabilities, and preserve
numbering through the public facade. After a committed upstream release,
update exact pins together and rerun supported/diagnostic/native matrices,
actual Office.js transport and independent Word checks before widening routing.
No library edits, commits, or consumer fidelity workarounds were made here.
