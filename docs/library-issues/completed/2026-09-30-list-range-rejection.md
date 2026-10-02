# Reject All merges the original paragraphs of a replaced list range

> **Resolved:** Fixed in 0.8.3; re-verified offline against 0.8.4 on 2026-10-02 (rejected text equals the source).

**Resolution:** Fixed in installed 0.8.3. The reproduction below describes the earlier defect; see [0.8.3 validation](../../validation-reports/2026-09-30-docx-redline-v083.md).

**Package:** `@ansonlai/docx-redline-js@0.8.2`  
**Host:** Microsoft Word 16.0, build 16.0.20430  
**Observed:** 2026-10-01  
**Scope:** standalone tracked replacement of a multi-paragraph bullet-list range

## Reproduction input

Use the Word-authored source `tests/fixtures/agentic-lists/nested-lists-source.docx`. Paragraphs 3–4 are separate entries in one nested bullet list: `Bullet Root A` and `Bullet Insertion Anchor`. The request was:

```json
{
  "tool": "edit_list",
  "startParagraphIndex": 3,
  "endParagraphIndex": 4,
  "newItems": [
    "Planner replacement parent",
    "    Planner replacement child"
  ],
  "listType": "bullet",
  "numberingStyle": "decimal"
}
```

The add-in planner maps this exact source range to the following canonical standalone operation:

```json
{
  "type": "redline",
  "target": {
    "index": 3,
    "exactText": "Bullet Root A",
    "paragraphId": "318182D1",
    "fingerprint": "fnv1a32:912ad3ce",
    "inTable": false
  },
  "targetEnd": {
    "index": 4,
    "exactText": "Bullet Insertion Anchor",
    "paragraphId": "00FEF05D",
    "fingerprint": "fnv1a32:c64da078",
    "inTable": false
  },
  "modified": "- Planner replacement parent\n    - Planner replacement child"
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

The tracked package opens in Word with two body revisions and no comments. Accept All produces the expected replacement parent and nested child. Reject All should restore the two original paragraphs as separate list entries.

## Word result

On Word 16.0.20430, source, tracked, and accepted views passed. Reject All failed the exact-text comparison. The expected source text was:

```text
Bullet Root A
Bullet Insertion Anchor
```

Word produced one paragraph with concatenated text:

```text
Bullet Root ABullet Insertion Anchor
```

The rejected package has one fewer paragraph and no longer preserves the source boundary between the two list items. This result was reproduced by the independent Word host check; see [the saved Word report](../../validation-reports/2026-09-30-agentic-list-known-defects-word.json) and `tests/agentic_list_fidelity_tests.mjs` for the host evidence and package exporter.

## Minimal standalone reproduction

From the repository root, generate the source, tracked, accepted, and rejected packages for both known defects, then run the Word host verifier on the known-defect manifest:

```powershell
node tests/agentic_list_fidelity_tests.mjs --export-host-dir .cache/agentic-tools/list-known-defects-word
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 `
  -FixtureManifest .cache/agentic-tools/list-known-defects-word/known-defects-manifest.json `
  -ArtifactsDir .cache/agentic-tools/list-known-defects-word
```

The command is expected to report the rejection check as failed: that is the recorded library defect, not a successful verification. The generated tracked package is opened and resolved by the Word host; compare the `rejected` view with the source text in the manifest. The full standalone API call for this case is in `tests/agentic_list_fidelity_tests.mjs` and the generated operation above. The standalone accepted-view output by itself does not reveal the defect.

The exporter writes `list-edit-bullet-range-diagnostic-tracked.docx`, `list-edit-bullet-range-diagnostic-accepted.docx`, and `list-edit-bullet-range-diagnostic-rejected.docx` into the selected output directory.

**Required library behavior:** Reject All must restore `Bullet Root A` and `Bullet Insertion Anchor` as two distinct source paragraphs. No consumer-side document or numbering repair is included in this report.
