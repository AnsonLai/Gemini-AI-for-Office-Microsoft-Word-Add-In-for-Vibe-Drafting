# Existing-comment delta transport in the consumer

## Status

Resolved in consumer package preparation and verified in actual Office.js and
independent desktop Word. The [WP5 correction report](../validation-reports/2026-10-02-lightweight-harness-wp5-correction.md)
records the final evidence: three existing comments become four, then five in a
second rewrite/comment batch; checkpoint restoration returns to four. Original
replies, resolved state and durable identities are preserved throughout.

## Classification

This was a consumer integration bug, not an `@ansonlai/docx-redline-js` engine
defect or a missing-library-API limitation. The standalone batch reports
ordinary comment output with `commentsXmlMode: 'merge'`; its `commentsXml` is a
delta by design. The consumer must merge that delta before putting the complete
`word/comments.xml` part into the Office.js Flat OPC payload.

The public `injectCommentsIntoPackage` helper is not a merge operation. It
appends another `/word/comments.xml` package part. However, the supported
`ensureCommentsArtifactsInZip` helper can be composed with public `MemoryZip`
and `reconcileCommentSiblingParts` to merge the delta into the existing comments
part and bring optional sibling IDs in line with that complete result. The
consumer then places those full parts into its Flat OPC payload and maintains
the package relationship through its existing adapter.

## Evidence

The public-only reproducer uses the Word-authored
[`threaded-comments.docx`](../../tests/fixtures/wp6/threaded-comments.docx)
fixture:

```powershell
node docs/library-issues/2026-10-02-word-existing-comments-delta-reproducer.mjs
```

With the installed 0.8.4 package, it reports three source comments, a two-edit
batch with `commentsXmlMode: "merge"`, one comment in the returned delta, and
four comment anchors in the returned document. Applying the published package
injector to an existing comments part creates two comments parts, with three
and one comments respectively. Composing the ZIP merge helper with
`reconcileCommentSiblingParts` instead yields a package with all four comments,
the original reply and resolved-thread state, the original extended-thread
part and durable IDs, plus reconciled sibling records for the new comment. The
same batch retains its requested paragraph edit. The reproducer also keeps a
text-only edit through the complete public DOCX API as a comparison.

Before the consumer fix, package preparation replaced the original comment
part with the one-comment delta. A fresh Office.js run on this fixture
attempted one `body.insertOoxml` and returned `GeneralException`; the receipt
marked the mutation indeterminate and required inspection, so the operation
was not retried. This remains historical pre-fix evidence; it does not show
that the composed merge path fails. The captured report is
`.cache/reliability/wp5/exploratory-comment-add/officejs/officejs-report.json`.
The captured source and fixture are under
`.cache/reliability/wp5/exploratory-comment-add/officejs/`.

## Consumer boundary

`prepareCanonicalBatch` now uses the public ZIP helper on a `MemoryZip`
containing the transported comment parts, then runs
`reconcileCommentSiblingParts` on the complete merged `comments.xml`. It copies
only changed comment parts back into the Flat OPC package and uses the existing
relationship adapter. A full `replace` result still replaces the comments
part, and full replacement `commentsExtended.xml` output remains honored. A
partial extended-thread delta against an existing extended part is rejected
rather than installed as a replacement.

The public-only reproducer and consumer-core regression suite now verify the
composition on the Word-authored threaded fixture. Actual Office.js/Word
validation of this consumer fix is pending; root will add the linked report
after the host gate completes.
