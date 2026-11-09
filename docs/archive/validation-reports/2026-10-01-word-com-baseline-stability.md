# Word COM baseline stability probe (2026-10-01)

## Result

On a generated five-paragraph DOCX fixture, the consumer's parsed paragraph text and fingerprints stayed the same across two source reads, switching Word revision tracking on and off, and rejecting a tracked edit. After `RejectAllRevisions`, Word returned the original five paragraph baselines, including `Target paragraph P4 baseline.` at index 4.

This does not reproduce the reported missing P4 anchor after a Word restart. The probe did not complete save/reopen or undo: an attempt to continue past the seventh capture hung while saving the disposable fixture. Those phases are listed as unrun in the capture output. This probe used Word COM `Document.Content.WordOpenXML`; it does not establish what Office.js `body.getOoxml()` returns.

## Captures

The capture sequence was:

1. Read the source twice.
2. Turn `TrackRevisions` on and read; turn it off and read.
3. Replace P4 with tracked changes and read.
4. Call `RejectAllRevisions()` and read; turn tracking off and read again.

The five compared baseline snapshots were all equal to the initial snapshot: source read B, tracking on, tracking off, after Reject All, and after Reject All with tracking off. Each had the same text and fingerprint for all five paragraphs. All parsed `paragraphId` values were `null`; this Word XML did not provide paragraph IDs to the consumer parser. The tracked edit capture showed the edited P4 and a temporary empty paragraph before P5; Reject All restored the original five-paragraph baseline.

The [saved comparison](2026-10-01-word-com-baseline-stability.json) contains individual fingerprints and raw XML SHA-256 values. The analyzer also writes `.cache/word-baseline-probe/baseline-comparison.json`. Raw XML hashes differed between some no-edit reads even though the parsed baseline was identical, so this result concerns the consumer's parsed paragraph fields rather than byte-for-byte XML equality. The partial capture record did not retain Word version/build; these fields are unknown in this saved run. Only the two Word processes created by the probe were stopped after the hung continuation; the pre-existing user Word process was left untouched.

## Reproduce

From the repository root, run `node scripts/word-desktop/run-baseline-stability-probe.mjs` to generate the fixture and capture the seven COM snapshots. This opens only the generated fixture at `.cache/word-baseline-probe/baseline-source.docx`. To analyze snapshots already in that directory without launching Word, add `--analyze-existing`.
