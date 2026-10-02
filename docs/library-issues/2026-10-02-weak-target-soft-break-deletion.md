# Weakly targeted soft-break replacement silently deletes the paragraph

Repository: `AnsonLai/docx-redline-js`. Reproduces in installed 0.8.4 (the
path only became reachable when 0.8.4 enabled localized replacements in
soft-break paragraphs). Not yet filed upstream.

## Impact

Data loss with a success status. A localized `replacements` edit inside a
paragraph that contains `<w:br/>`, targeted only by `index` + `exactText`
(no `paragraphId`/`fingerprint`), returns `status: 'ok'` but tracks the whole
paragraph as deleted and inserts nothing; Accept All removes the paragraph.

The add-in is not exposed today: its planner always sends the full strong
target (`paragraphId` and `fingerprint` from the inspected source), and with
that target the same edit is correct. Other consumers (CLI/MCP callers that
target by text and index only) are exposed.

## Expected

Either apply the replacement correctly (as with the strong target) or refuse
the operation. Never report success while deleting the paragraph.

## Reproduction

```powershell
node docs/library-issues/2026-10-02-weak-target-soft-break-deletion-reproducer.mjs
```

Prints the accepted text for both targets: empty for the weak target,
`First recital.\nParties desire project Titan.` for the strong target. Exits 0
while the defect reproduces. A paragraph without soft breaks is unaffected.
