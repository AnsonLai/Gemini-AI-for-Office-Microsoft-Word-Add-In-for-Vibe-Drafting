# Table creation and chat reliability validation

**Date:** 2026-10-01. **Package:** exact `@ansonlai/docx-redline-js@0.8.3`.

## Consumer checks

- `npm test`: 54 suites passed, zero failed; four excluded entrypoints.
  The older 0.5.4 compatibility suite reports its existing version-specific skip.
- `npm run build`: production build passed with the existing chunk-size warning.
  Blocked optional Application Insights telemetry did not prevent compilation.
- Regression coverage includes context refresh after confirmed structural writes,
  source freshness checks, retained user requests and complete tool exchanges,
  failed-mutation budgets across read-only calls, terminal loop status, safe
  warning diagnostics, append anchors, narrow batch coalescing and receipt mapping.
- Formatting combinations known to lose fidelity are refused before a Word write.
  No library patch or dependency change is included.

## Independent Word checks

Microsoft Word 16.0, build 16.0.20430, opens the generated DOCX fixtures without
repair. The oracle checks direct body paragraphs separately from table cells,
including all nine noun values and the three-by-three table dimensions. It also
checks underline and resolves tracked changes through Word's own Accept All and
Reject All commands.

The initial run found an additional direct paragraph after accepting the table
and after rejecting its tracked insertion. The accepted document needs Word's
required paragraph after a final table. The rejected document is still required
to restore exactly the seven source paragraphs; its extra paragraph is a failure.
The library's separately resolved rejected package restores seven paragraphs,
which does not establish Word Reject All fidelity.

The final run completed **20 checks: 13 passed, seven failed**. Both native
`InsertXML` calls succeeded, using operations prepared against genuine Word
`Content.WordOpenXML`, followed by save, reopen and independent revision resolution.

| Fixture | Word result |
| --- | --- |
| Plain final-paragraph edit plus table | Accept All preserves the edited paragraph, all nine nouns and the exact 3×3 table, through both direct DOCX and native insertion. |
| Plain tracked table, Reject All | Both paths restore the seven original paragraphs **plus an unwanted eighth empty paragraph**. The source has seven. |
| Same-author underline then table | Accept All loses the underline in the tracked DOCX and the engine's accepted export. |
| New underline and table in one replacement | Native insertion succeeds, but Word Accept All loses the underline. |
| Formatting fixture, Reject All | Direct DOCX and native insertion also leave the extra empty paragraph. |
| Engine-resolved rejected exports | Both contain the original seven paragraphs and pass Word inspection; they differ from Word independently rejecting the tracked output. |

The [raw Word report](2026-10-01-table-creation-word-checks.json) records each
check. These are fidelity failures, not failures to open or insert the packages.
The incident plan remains open for the unresolved library requirements.

## Reproduction

```powershell
node tests/table_append_batch_tests.mjs --export-host-dir .cache/table-append-host
node docs/library-issues/completed/2026-10-01-table-append-tracked-formatting-reproducer.mjs
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 -FixtureManifest .cache/table-append-host/manifest.json -ArtifactsDir .cache/table-append-word -TimeoutSeconds 180
```

The standalone formatting diagnostic succeeds when the known library defects
reproduce; that success is not a fidelity pass. The Word oracle exits unsuccessfully
while the independent fidelity expectations fail. Generated documents and JSON
reports are under the specified `.cache` directories. No live Gemini request,
actual Office.js collector run or PDF export is part of this incident validation.

See the [incident plan](../plans/2026-09-30-table-creation-reliability.md) and the
[separate formatting report](../library-issues/completed/2026-10-01-table-append-tracked-formatting.md)
and [Word Reject All paragraph report](../library-issues/2026-10-01-word-reject-table-append-paragraph.md).
