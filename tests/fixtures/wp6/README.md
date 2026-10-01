# WP6 regression fixtures

These fixtures verify delegated library editing and the consumer's retained
Flat OPC transport separately. See the [library offload review](../../../docs/library-offload-review.md)
and [verification commands](../../../scripts/README.md) for that ownership split.

The three DOCX files are copied byte for byte from the local `Docx Redline JS` reference project, `tests/fixtures/word-authored/`, on 2026-09-30. They were authored and saved in Microsoft Word Desktop through that project's COM generator scripts. These are source packages, not engine-generated expected outputs.

| File | Source generator | SHA-256 |
| --- | --- | --- |
| multi-paragraph-thread.docx | scripts/generate-word-thread-state-fixtures.ps1 | ADB90C5E274CA6D8818A332DF85A5C433C7CEE0D1BF78CF3A4E47051412FDE0C |
| header-footer.docx | scripts/generate-word-header-footer-fixture.ps1 | 963F53D5B390676394221CA9F312AED74EAD43A1A110FB02B6D7D8B8F68DB4F9 |
| threaded-comments.docx | scripts/generate-word-comment-thread-fixture.ps1 | D8FA5D22E92F29C3B74FDF1FCA71C928E93AE00D2680D8878E448DC3E46032B5 |

`formatting-layout.xml` is a small synthetic fixture authored for this consumer upgrade. It contains distinct bold/plain/italic runs, a right tab stop and tab, a line break, an external hyperlink, a paragraph-level portrait section boundary and a landscape two-column final section. Tests package it with an external hyperlink relationship and a style definition.

Methodology follows `Docx Redline JS/docs/TESTING.md`, `tests/visual_failure_regression_tests.mjs` and `tests/comment_thread_word_fixture_tests.mjs`: inspect XML semantics, require exact accepted/rejected text, preserve unaffected package bytes, validate real Word package threading and sibling-part identities. The suite's historical `visual` name means semantic guards against visible formatting regressions; it does not prove rendered page appearance or opening without repair. Word Desktop opening/rendering is a separate validation lane.

Run the verified lane with `node tests/ooxml_formatting_visual_tests.mjs`. Add `--export-dir .cache/wp6/fidelity` to create source/tracked/accepted/rejected DOCX states and a manifest for the Word oracle; the bridge case also exports its actual Flat OPC insertion payload.

The two defects originally reproduced with v0.8.1 are fixed in v0.8.2. Four mandatory regressions now cover localized and full-paragraph forms of each edit: hyperlink punctuation must stay outside the link, and Reject All must exactly restore text around the tab/manual break. They execute without optional probe flags and enter the default live Word lane. The combined formatted/tab/break/hyperlink source remains atomically refused by the bridge with a no-write guard.

The Word-authored footer case edits only the first-page footer. The default footer PAGE field, all headers, section references, body and other package parts must remain byte-identical through edit/accept/reject. Its manifest header/footer expectations use COM one-based section numbering, with explicit first-page selection.

Export includes all four fixed-defect regression packages for independent Word Accept/Reject comparison. Expected body text comes from fixed source literals and the requested replacement; observed engine output is diagnostic only. `expectedHyperlinks` specifies the hyperlink display boundary in source, accepted and rejected views. Required assertions check exact text, hyperlink boundaries and relationships, and package validity before export. These cases have no `knownFailure` exclusion.

`word-addin-plain-replacement` isolates the actual Flat OPC bridge payload with three plain paragraphs, empty relationships, and no comments/styles/hyperlinks/fields/tables. Independent literal text expectations require `Alpha BETA gamma.` after Accept All and exact `Alpha beta gamma.` after Reject All, with the two remaining paragraphs unchanged. It supplies `insertionXml` for live Word insertion checks.
