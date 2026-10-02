# 0.8.4: separate list operations in one batch restart numbering

Repository: `AnsonLai/docx-redline-js`. **Regression in 0.8.4**; 0.8.3 behaved
correctly. Not yet filed upstream.

## Consumer impact

The add-in's canonical `convert_headers_to_list` plan (`list-operation-plan.js`)
emits one redline operation per header: `A. Supported Header First`,
`B. Supported Header Second`. On 0.8.3 both generated list paragraphs shared
one numbering instance and Word labelled them A., B. On 0.8.4 each operation
allocates its own numbering instance, so the second header restarts at "A.".

The production `convert_headers_to_list` tool still runs on Word's native list
API, so users are not affected today, but this blocks migrating that tool to
the canonical engine path (and any batch that continues one generated list
across non-adjacent paragraphs). The add-in fidelity case
`list-convert-noncontiguous-manual-headers` is skipped with a pointer here.

## Expected

Operations in one batch that generate items of the same list kind and format
(continuing markers `A.` then `B.`) continue one numbering instance, as in
0.8.3, while still using IDs that do not collide with the source numbering.

## Actual

```powershell
node docs/library-issues/2026-10-02-separate-list-operations-restart-numbering-reproducer.mjs
```

On the Word-authored `tests/fixtures/agentic-lists/nested-lists-source.docx`
the two headers resolve to `numId 7` and `numId 11`, both labelled `A.`. The
reproducer exits 0 while the defect reproduces.

Likely cause: the 0.8.4 numbering-collision fix allocates a fresh instance per
generated list without reusing an instance created earlier in the same batch
for a continuing marker.
