# `generateRedlines: false` still tracks surgical insertions

> **Resolved in 0.9.0 (2026-10-12):** Fixed in `@ansonlai/docx-redline-js@0.9.0`. Reject All now keeps the clean text, and the reproducer's defect assertion throws, so it exits 1. The unchanged default 14-case Office.js manifest, including untracked pure insertion on a revised target, passes. The body below is the historical 0.8.4 report. See the [0.9.0 validation](../../validation-reports/2026-10-12-docx-redline-v090.md).

## Status

Unresolved upstream issue reproduced against the installed
`@ansonlai/docx-redline-js@0.8.4` on 2026-10-02.

## Expected and actual behavior

The public `RedlineOptions` and standalone `DocumentOperationBase` expose
`generateRedlines`. With `generateRedlines: false`, a text insertion should be
committed as ordinary text, without a `w:ins` revision. The installed library
instead wraps an interior insertion in a new tracked insertion when
`existingRevisions: 'slice-cross-author'` selects surgical reconciliation.

The portable public-only reproducer is:

```powershell
node docs/library-issues/completed/2026-10-02-generate-redlines-false-cross-author-insertion-reproducer.mjs
```

It calls both the public `applyRedlineToOxml` facade and
`applyOperationsToDocumentXml`. The batch variants cover both `type: 'redline'`
and `type: 'replace'`, using both `modified` and `replacements`, while setting
`generateRedlines: false` on the operation and in the batch options. Each
returns success and creates a new `w:ins` by `Probe`. Public Accept All yields
the requested text, while Reject All restores the original text. Thus accepted
text validation alone does not expose the disabled-tracking violation.

As a control, the same direct call with `existingRevisions: 'merge-same-author'`
produces plain text: both accepted and rejected views are
the modified text, with no new insertion revision. This isolates the failure
to the cross-author surgical route.

The behavior is independent of how the replacement is expressed and of the
per-operation option. The application forwards the documented option; this is
inside the published library's surgical insertion path, not a consumer option
mapping omission. The probe exits 0 while the defect reproduces and 1 when the
observed behavior changes.

## Add-in impact

The original trigger came from an actual Word checkpoint at
`.cache/reliability/wp5-correction/pre-untracked-insertion-failure/direct-literal-replacement-preserves-comment-threads-checkpoint.xml`.
Its targeted clause was `This agreement is governed by local law.`; a literal
replacement inserted `applicable ` with `generateRedlines: false` and
`existingRevisions: 'slice-cross-author'`. The generated insertion appeared as
a separate tracked revision despite the option being disabled. This changes
review behavior: Reject All can discard text that the caller requested as an
untracked edit.
Independent Word 16.0.20430 confirms one native revision in the source, two
after the edit, and the old wording after Reject All. Its report is
`.cache/reliability/wp5-correction/literal-revision-probe/word-report.json`.

The consumer's `TRACKING_POLICY_MISMATCH` check now blocks new own tracked text
identified by public preparation segments and receipt-created IDs before
checkpoint/write. It preserves supported untracked replacement/formatting and
existing foreign carriers. This is safety evidence; the failed insertion remains
a required capability gate in the
[WP5 correction report](../../validation-reports/2026-10-02-lightweight-harness-wp5-correction.md).

The synthetic reproduction removes document-specific factors and invokes only
published library APIs. No installed package files were changed. Keep the exact
library pin until an upstream release fixes this case; do not compensate with
local OOXML rewriting or a different insertion affinity.
