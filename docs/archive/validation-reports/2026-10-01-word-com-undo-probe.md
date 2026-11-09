# Word COM Undo probe (2026-10-01)

## Result

On the same generated five-paragraph DOCX fixture, Word desktop (16.0, build 16.0.20430, driven over COM with a hidden window) restored the original paragraph text and fingerprints after `Document.Undo()`. In both scenarios every post-Undo capture matched the initial baseline for all five paragraphs (text and fingerprint; no changed indexes):

1. Tracked edit of P4 via `Range.Text`, then `Undo(1)` (returned true, revision count 0), then tracking off.
2. Untracked P4 rewrite via `Range.InsertXML` (the edit visibly changed P4), then `Undo(1)` (returned true), then a second read.

The tracked and untracked captures after Undo are captures 11, 12, 14 and 15 in [the saved comparison](2026-10-01-word-com-undo-probe.json).

## Limits

- The Word COM `Content.WordOpenXML` output contained no `w14:paraId` attributes at all, even though the fixture supplied them (the paraId count was 0 in the source, edited and post-Undo captures). Parsed `paragraphId` was therefore null throughout, so this probe cannot say whether Undo changes paragraph IDs as seen by Office.js. That remains the leading untested candidate for the user's double STALE_DOCUMENT_CONTEXT refusal after Undo.
- COM was used, not Office.js `body.getOoxml()`, and Word was driven by script rather than the add-in's Ctrl+Z path, so the user's failure is not reproduced.
- Word restart (save/reopen) was not run; the earlier probe hung while saving, and this one avoids saving.
- The fixture is minimal (five plain paragraphs), not the user's document.

No fingerprint guard logic was changed.

## Reproduce

`node scripts/word-desktop/run-baseline-stability-probe.mjs --undo` (writes to `.cache/word-undo-probe`, 120 s timeout).
