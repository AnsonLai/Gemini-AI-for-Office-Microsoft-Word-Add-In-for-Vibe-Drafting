# Reject All leaves an extra paragraph after a plain-anchor list insertion

> **Resolved:** Fixed in 0.8.3; re-verified offline against 0.8.4 on 2026-10-02 (rejected text equals the source).

**Resolution:** Fixed in installed 0.8.3. The reproduction below describes the earlier defect; see [0.8.3 validation](../../validation-reports/2026-09-30-docx-redline-v083.md).

**Package:** `@ansonlai/docx-redline-js@0.8.2`  
**Host:** Microsoft Word 16.0, build 16.0.20430  
**Observed:** 2026-10-01  
**Scope:** standalone tracked insertion after a plain paragraph

## Reproduction input

Use the Word-authored source `tests/fixtures/agentic-lists/nested-lists-source.docx`. Paragraph 2 is `Plain paragraph before bullet list.`; paragraph 3 begins the following bullet list. The request was:

```json
{
  "tool": "insert_list_item",
  "afterParagraphIndex": 2,
  "text": "Planner plain insertion after paragraph",
  "indentLevel": 1
}
```

The add-in planner maps this plain anchor to the following canonical standalone operation:

```json
{
  "type": "redline",
  "target": {
    "index": 2,
    "exactText": "Plain paragraph before bullet list.",
    "paragraphId": "31AD2DC5",
    "fingerprint": "fnv1a32:173e2ac2",
    "inTable": false
  },
  "modified": "Plain paragraph before bullet list.\nPlanner plain insertion after paragraph"
}
```

The operation was applied with the public 0.8.2 `applyOperationsToDocumentXml` standalone API using the source `word/document.xml`, `word/numbering.xml` and `word/styles.xml` parts:

```js
const result = await applyOperationsToDocumentXml(
  sourceDocumentXml,
  [operationAbove],
  'Agentic List Fidelity Test',
  { numberingXml: sourceNumberingXml, stylesXml: sourceStylesXml },
  {
    atomic: true,
    strictTargets: true,
    structuredContent: true,
    generateRedlines: true,
    existingRevisions: 'merge-same-author',
    date: '2026-09-30T12:00:00Z'
  }
);
```

The generated tracked package opens in Word with one body revision and no comments. Accept All produces the expected insertion after paragraph 2. Reject All should restore the source, including the adjacency of paragraphs 2 and 3.

## Word result

On Word 16.0.20430, source, tracked, and accepted views passed. Reject All failed the exact-text comparison. The expected local sequence was:

```text
Plain paragraph before bullet list.
Bullet Root A
```

Word produced:

```text
Plain paragraph before bullet list.

Bullet Root A
```

The rejected package has an additional empty paragraph between the insertion anchor and the first list item. The rest of the document text matches the source. This result was reproduced by the independent Word host check; see [the saved Word report](../../validation-reports/2026-09-30-agentic-list-known-defects-word.json) and `tests/agentic_list_fidelity_tests.mjs` for the host evidence and package exporter.

## Minimal standalone reproduction

From the repository root, generate the source, tracked, accepted, and rejected packages for both known defects, then run the Word host verifier on the known-defect manifest:

```powershell
node tests/agentic_list_fidelity_tests.mjs --export-host-dir .cache/agentic-tools/list-known-defects-word
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 `
  -FixtureManifest .cache/agentic-tools/list-known-defects-word/known-defects-manifest.json `
  -ArtifactsDir .cache/agentic-tools/list-known-defects-word
```

The command is expected to report the rejection check as failed: that is the recorded library defect, not a successful verification. The generated tracked package is opened and resolved by the Word host; compare the `rejected` view with the source text in the manifest. The full standalone API call for this case is in `tests/agentic_list_fidelity_tests.mjs` and the generated operation above. The standalone accepted-text result alone does not reveal the defect.

The exporter writes `list-insert-after-plain-paragraph-diagnostic-tracked.docx`, `list-insert-after-plain-paragraph-diagnostic-accepted.docx`, and `list-insert-after-plain-paragraph-diagnostic-rejected.docx` into the selected output directory.

**Required library behavior:** Reject All for this insertion must restore the original source exactly, without leaving the inserted paragraph or creating an empty paragraph. No consumer-side document or numbering repair is included in this report.
