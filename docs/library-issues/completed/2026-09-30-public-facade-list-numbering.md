# Public DOCX facade changes unrelated list numbering during range replacement

> **Resolved:** Fixed in 0.8.3 per the dated Word evidence in `docs/validation-reports/2026-09-30-docx-redline-v083.md`; the 0.8.3 compatibility suite still passes on 0.8.4 (2026-10-02). Word was not re-run for 0.8.4.

**Resolution:** Fixed in installed 0.8.3. The reproduction below describes the earlier defect; see [0.8.3 validation](../../validation-reports/2026-09-30-docx-redline-v083.md).

**Checked:** 2026-09-30, published 0.8.2 and the unpublished library working tree.
**Scope:** public `openDocx(...).applyOperations` facade, used by browser/MCP;
the add-in's existing standalone runner and numbering merge path passes the same
Word expectations. This is a pre-existing numbering issue, not a regression
introduced by the paragraph-mark fixes.

Using `tests/fixtures/agentic-lists/nested-lists-source.docx`, apply:

```js
const result = await doc.applyOperations([{
  type: 'redline',
  target: { index: 3, exactText: 'Bullet Root A' },
  targetEnd: { index: 4, exactText: 'Bullet Insertion Anchor' },
  modified: '- Planner replacement parent\n    - Planner replacement child'
}], {
  author: 'Consumer Review', atomic: true, strictTargets: true,
  structuredContent: true, generateRedlines: true
});
```

The operation returns `ok`. Independent Word reopening confirms restored source
text boundaries but rejects the numbering expectations:

- Accept All: untouched `Bullet Untouched Tail` has `ListValue` 1 instead of 2.
- Reject All: source `Bullet Root A` has `ListType` 2 instead of the source's 4.
- These failures occur for both Word-resolved output and library-resolved packages.

[Word report](../../validation-reports/2026-09-30-unpublished-list-facade-word.json)
records eight passing and four failed checks across the plain insertion control
and range replacement. Plain insertion passes all six checks.

Running the same facade operation with published 0.8.2 produces a
`word/numbering.xml` byte-identical to the unpublished facade output; both differ
from the source numbering part. The existing
[canonical capabilities report](../2026-09-30-canonical-list-operations.md)
already tracks numbering-definition allocation/remapping limitations. Extend
its tests to preserve unrelated numbering on same-kind edits and restore source
numbering on Reject All through the complete DOCX facade. Do not compensate by
rebuilding numbering in browser/MCP consumers.
