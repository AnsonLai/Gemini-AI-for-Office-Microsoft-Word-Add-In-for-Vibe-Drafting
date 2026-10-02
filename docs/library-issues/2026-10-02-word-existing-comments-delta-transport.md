# Existing comments need a safe merge-delta transport

## Classification

This is a consumer transport/API gap, not an `@ansonlai/docx-redline-js` engine
defect. The standalone batch reports ordinary comment output with
`commentsXmlMode: 'merge'`; its `commentsXml` is a delta by design. The consumer
must not install that delta as the complete `word/comments.xml` part.

The public `injectCommentsIntoPackage` helper is not a merge operation. It
appends another `/word/comments.xml` package part. The library also exports
`ensureCommentsArtifactsInZip`, which merges the comments part in a JSZip-style
`.docx` archive. That ZIP helper does not build the Flat OPC payload accepted by
Word's `body.insertOoxml`, or reconcile all existing comment-thread parts. There
is no supported public helper for producing a complete merged threaded-comment
package at the current Office.js Flat OPC boundary.

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
and one comments respectively. The public ZIP helper merges the XML comment
delta into a `.docx` archive and yields four comments; that is a different
transport boundary and only merges `comments.xml`. The complete public DOCX API
also applies a text-only edit and retains all three existing comments.

The app's package preparation used to replace the original comment part with
the one-comment delta. A fresh Office.js run on this fixture attempted one
`body.insertOoxml` and returned `GeneralException`; the receipt marked the
mutation indeterminate and required inspection, so the operation was not
retried. The captured report is
`.cache/reliability/wp5/exploratory-comment-add/officejs/officejs-report.json`.
The captured source and fixture are under
`.cache/reliability/wp5/exploratory-comment-add/officejs/`.

## Consumer boundary

`prepareCanonicalBatch` now refuses with `UNSUPPORTED_EXISTING_COMMENTS` before
building an insertion payload when the library returns a non-empty merge-mode
comment delta and canonical inspection found comments in the source. This is a
fail-closed fallback; it does not alter the library's delta contract. A
complete `replace` result remains supported. Text-only edits continue to
preserve existing comment parts. The taskpane also rejects new comments when
the source already contains comments, before invoking the library.

Until a supported complete-package merge route is available at the Office.js
Flat OPC boundary, keep comment additions on documents with existing comments
refused. Do not call `injectCommentsIntoPackage` as a workaround, and do not
change the engine's merge semantics to return a full replacement implicitly.
