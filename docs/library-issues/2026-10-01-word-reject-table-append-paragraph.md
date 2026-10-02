# Word Reject All leaves an extra paragraph after a final table append

## Status after 0.8.4 (2026-10-02): still open

0.8.4 changed tracked tables to row-level `w:trPr/w:ins` marks (no `w:ins`
around `w:tbl`), and Word's Reject All now removes the table without hanging.
A fresh desktop Word check (Word 16.0, build 16.0.20430, `OpenNoRepairDialog`,
`Revisions.RejectAll()` on both the plain-append and prior-underline tracked
packages built with 0.8.4) still ends with the original seven paragraphs **plus
one extra empty paragraph** and one revision left. Accept All in Word keeps
the 3×3 table and the paragraph underline. The engine's own
`rejectTrackedChangesInOoxml` still returns exactly seven paragraphs, so the
remaining mismatch is in how Word resolves the trailing paragraph mark of an
end-of-document table append. The underline-loss half of this incident is
fixed (see `completed/2026-10-01-table-append-tracked-formatting.md`).

## Reproduction

From the repository root, export the table-append fixtures and run the Word oracle:

```powershell
node tests/table_append_batch_tests.mjs --export-host-dir .cache/table-append-host
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 `
  -FixtureManifest .cache/table-append-host/manifest.json `
  -ArtifactsDir .cache/table-append-word
```

The captured run used Microsoft Word 16.0, build 16.0.20430. The oracle report and fixture set are saved at:

- `.cache/table-append-word/word-report.json`
- `.cache/table-append-word/word-stdout.log`
- `.cache/table-append-host/manifest.json`
- `.cache/table-append-host/table-append-tracked.docx`
- `.cache/table-append-host/table-append-coalesced-plain-edit-native-insert.docx`
- `.cache/table-append-host/table-append-inline-format-known-library-defect-native-insert.docx`
- `.cache/table-append-host/table-append-rejected.docx`

The native `InsertXML` operation passed for both cases. The lane reads Word's live `Content.WordOpenXML`, runs the production batch bridge, inserts the resulting OOXML with Word's `InsertXML`, and saves the document. Subsequent Word revision-view checks still found the Reject All and underline issues described below. This verifies the Word-native insertion path but does not establish Office.js transport provenance.

## Results

The plain edit-plus-table case passes Word's tracked and accepted checks, including the complete 3 by 3 table. Word's accepted view has one trailing empty direct paragraph after the final table, which is expected for the table-at-end layout and is recorded only in the accepted expectation.

Word's native Reject All view of the tracked package leaves the original seven paragraphs and an eighth empty direct paragraph:

```text
Opening paragraph 1.
Opening paragraph 2.
Opening paragraph 3.
Opening paragraph 4.
Opening paragraph 5.
Opening paragraph 6.
Final paragraph stays unchanged.
<empty paragraph>
```

This happens for both the tracked package and the Word `InsertXML` result. The source document has exactly seven paragraphs, and its rejected expectation remains seven. The separately engine-resolved Reject All DOCX passes with exactly those seven paragraphs. The standalone reproducer also confirms `rejectTrackedChangesInOoxml` returns the original seven paragraphs and no table. This isolates the extra paragraph to Word's revision resolution of a table appended at the document end; it is not required by the source or by the engine's own reject resolver.

The Word report records 13 passed checks and 7 failures. Those failures include the rejected-view paragraph mismatch above and the separate known underline-fidelity failure. The inline-format known-defect case expects an underline after Accept All; Word confirms it is absent in both the engine-resolved accepted package and the native Word view.

## Handling

Keep the exact seven-paragraph Reject All expectation in the Word manifest so the structural mismatch remains visible. The accepted expectations alone include the terminal empty paragraph. The current consumer planner leaves plain table appends available; if exact Reject All restoration is required before this Word behavior is resolved, the narrow safety boundary to consider is end-of-document table insertion. Do not treat the trailing accepted-view paragraph as evidence that Reject All is safe.
