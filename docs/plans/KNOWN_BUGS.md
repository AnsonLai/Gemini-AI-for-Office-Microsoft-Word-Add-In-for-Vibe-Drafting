# Known bugs and host-boundary exceptions

**Updated:** 2026-10-17, phase-2 WP3 against a local `@ansonlai/docx-redline-js` 0.10.0 build (unpublished; the pin is still 0.9.0).
Previous: 2026-10-12, after the 0.9.0 upgrade ([validation](../validation-reports/2026-10-12-docx-redline-v090.md)).

This file tracks two things:

- Defects and gaps we know about but have not fixed.
- Every place production code uses the Word JavaScript API (Office.js) instead of the library.

The goal is to use the library wherever possible. Each Word API use must say why it exists and what would let us remove it.

Library defects still go in [`../library-issues/`](../library-issues/README.md), with a portable reproducer. This file links to them and does not duplicate them.

## Open bugs and gaps

### KB-1: Body-final table append via live Word insertion leaves an extra paragraph on Reject All

- **Where:** the production Word insertion of a library-prepared table append at the end of the body, in [`word-operation-runner.js`](../../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js). The library's own output is correct.
- **Evidence (2026-10-12, Word 16 desktop):**
  - Exporter: [`2026-10-01-word-reject-table-append-paragraph-export.mjs`](../library-issues/completed/2026-10-01-word-reject-table-append-paragraph-export.mjs).
  - Oracle command: `verify-wp6-word.ps1 -CaseName public-library-table-append-reject-all`.
  - The source, tracked, Accept All and Reject All checks pass on the 0.9.0 package: 7 paragraphs and no revisions after `Revisions.RejectAll()`.
  - The oracle's native-insert lane fails Reject All with 8 paragraphs against 7, plus a trailing empty paragraph. That lane inserts the production adapter's payload into a live document with COM `Range.InsertXML`.
  - Word appears to add an untracked empty paragraph after an inserted body-final table.
- **User impact today:** none. The agent refuses structural table edits and table appends. This is a latent blocker for bringing table append back.
- **Exit condition:** an Office.js run of a body-final table append through the production adapter. Its saved export must give 7 paragraphs and 0 revisions under `Revisions.RejectAll()`, checked by the independent oracle. Any fix must stay in the transport (for example, how the replace range ends). No OOXML rewriting of library output.
- **Not yet established:** whether Office.js `body.insertOoxml` behaves like COM `InsertXML` here. COM evidence does not certify Office.js.

### KB-2: Whole-document `applyRedlineToOxml` corrupts unrelated paragraphs (unsupported usage)

- `applyRedlineToOxml` is for paragraph, range or table scope. A complete `document.xml` goes through `applyOperationsToDocumentXml`, as production does.
- Given a whole body, the library parses the modified text as Markdown. Unchanged literal text such as `7. REMEDIES`, `A. …`, `____` and `**x**` is reinterpreted, so unrelated paragraphs change in Accept All. This happens in 0.8.4 and 0.9.0.
- In this path, 0.9.0 also drops an inserted `2`, producing "two () copies".
- Production never uses this shape.
- `tests/addin/integration_tests.mjs` used it and passed only through OOXML substring assertions. Those cases now use targeted `applyOperationsToDocumentXml` and assert exact per-paragraph Accept All and Reject All text.
- **Follow-up (upstream, optional):** have `applyRedlineToOxml` refuse a multi-paragraph body input rather than return `ok`.
- **Exit condition:** no test or production caller uses the whole-body shape, which is met as of 2026-10-12. An upstream refusal would make this impossible to repeat.

### KB-3: `Document.RejectAllRevisions()` differs from the ribbon

- After a body-final table append, Word COM `Document.RejectAllRevisions()` and `AcceptAllRevisions()` leave the tracked last paragraph mark behind. `Document.Revisions.RejectAll()` and `AcceptAll()`, which the ribbon uses, resolve it fully.
- This is a Word API quirk, also noted in the library's 0.9.0 changelog.
- [`verify-wp6-word.ps1`](../../scripts/verify-wp6-word.ps1) now uses `Revisions.AcceptAll/RejectAll`.
- Oracle reports produced before 2026-10-12 used the `Document.*` methods.

### KB-4: List conversion and list-format changes (resolved for the exposed scope, 2026-10-14)

- 0.9.0 ships `list-format` (`convert-plain`, `change-format`, `join`) and `paragraph-format.numbering`. It also supports a composed `redline` plus `list-format` for changed-text headers, and returns explicit refusal codes.
- **2026-10-14:** `apply_document_edits` now exposes `convert_to_list`, `change_list_format` and `continue_list` as one-to-one maps onto `list-format` (contract in [`agentic-tool-contracts.md`](../agentic-tool-contracts.md)). The add-in parses no markers and writes no numbering. Offline evidence: [`document_list_intents_tests.mjs`](../../tests/document_list_intents_tests.mjs) (Accept All and Reject All on synthetic and Word-authored fixtures, library refusals with no write) and the direct golden step 12 replay.
- **Host evidence (2026-10-14):** five new Office.js capability cases plus one refusal control all pass. The Word oracle gives 88 PASS / 0 FAIL, with labels read from `ListFormat.ListString`. The real-Word NDA replay passes 16/16 (51/51 checks) with step 12 applied. See the [2026-10-14 list-format validation](../validation-reports/2026-10-14-list-format-intents.md).
- **Still open:** the live-provider lane, Word Online, and browser/MCP facade parity. Not exposed: range-scoped changes, `paragraph-format.numbering`, depth/indent changes, table lists, `detach`.
- Tracked in the [canonical list migration follow-up](2026-09-30-canonical-list-migration-follow-up.md) (WP1 to WP3).
- **Exit condition:** that plan's remaining acceptance items are met (facade parity; library-issues cleanup).

### KB-5: `TRACKING_POLICY_MISMATCH` guard (retired by phase-2 WP3, 2026-10-17)

- The add-in refused an untracked batch whose prepared output carried new own-author revisions, a workaround for the 0.8.4 surgical-insertion bug.
- 0.9.0 fixed that bug and enforces the same rule itself (`REVISION_FIDELITY_VIOLATION` / `UNTRACKED_CREATED_REVISION`). The guard never fired on the certified cases during the 0.9.x cycle.
- **Retired:** the guard is removed from `document-tools.js`; the library's refusal is relied on. The untracked capability cases in `document_tools_tests.mjs` still pass.

### KB-6: Revision-fidelity refusals reach the model unmapped

- 0.9.0 refuses edits that would break Reject All (`REVISION_FIDELITY_VIOLATION`, default `enforce`). We never pass `verifyRevisionFidelity`, which is correct.
- [`document-tools.js`](../../src/taskpane/modules/chat/document-tools.js) passes the library's `error.recovery` through verbatim. No retained case hits this in the current gates, but the model gets the library's recovery text, not a product-specific instruction.
- **Exit condition:** a deterministic test showing the agent's response to a `REVISION_FIDELITY_VIOLATION`. Map it only if that test shows the model mishandles the library guidance.

### KB-7: Library Markdown pre-processor collapses literal underscore runs

- **Where:** `@ansonlai/docx-redline-js` 0.8.4 and 0.9.0, `pipeline/markdown-processor.js` (`preprocessMarkdown`, the `__x__` and `_x_` patterns). Report: [`2026-10-12-table-reconciliation-underscore-collapse.md`](../library-issues/2026-10-12-table-reconciliation-underscore-collapse.md).
- **Evidence (2026-10-12):**
  - `tests/addin/integration_tests.mjs`, case "Add Date Row", prints `KNOWN_LIBRARY_DEFECT:` under its PASS line. A `table-reconciliation` on the sample NDA signature table shrinks `_________________________` to `_` in Accept All, including cells the edit did not change.
  - The case keeps its strict assertions. The marker is printed only when the sole mismatch is this underscore collapse. Any other mismatch still fails.
  - Reproducer: [`2026-10-12-table-reconciliation-underscore-collapse-reproducer.mjs`](../library-issues/2026-10-12-table-reconciliation-underscore-collapse-reproducer.mjs) (exit 0 while it reproduces; 0.8.4 behaves the same).
- **User impact today:** the table-reconciliation route is not reachable. `rewrite_paragraph` and `replace_text` forbid line breaks, and the route needs a multi-line pipe table. Insertions refuse targets and text that `preprocessMarkdown` would change (`plainParagraph`, `checkedInsertionTarget`). `replace_text` is safe because the library refuses a localized replacement it cannot round-trip (`PATCH_ROUNDTRIP_MISMATCH`). **However, `rewrite_paragraph` is exposed through the same root cause.** It sends a `redline` with `structuredContent: false`, and the library still runs `preprocessMarkdown` on it. Rewriting a paragraph whose text holds an underscore run (signature or date line, `__x__`) returns `status: 'ok'` with the run shrunk. Example: `_________________________` with ` (sign)` appended accepts as `_____ (sign)`. `document-tools.js` verifies Word against the prepared text, not the requested text, so nothing catches it. Signature-block rewrites are the likely trigger.
- **Exit condition:** a published library fix, with the reproducer exiting 1. Then remove the known-defect gate from "Add Date Row" (the `squash`/`knownDefect` logic and the marker). Consider a guard that compares the requested `rewrite_paragraph` text with the prepared text before the interim period ends.
- **Fixed in 0.10.0 (local build, 2026-10-17; not yet published):** underscore emphasis follows CommonMark flanking rules, and a caller's `structuredContent: false` keeps inline Markdown literal. The reproducer exits 1 against 0.10.0, and the "Add Date Row" known-defect gate is removed on branch `phase2-wp1-wp3`. Close this entry and move the library issue to `completed/` once 0.10.0 is published and pinned.
- **Related observation (not filed):** a `list-change` whose kept item carries `-` markers on a decimal list converts the list to a new bullet list (new `numId`, bullet labels). The only diagnostic is "Paragraph expanded to list fragment". The 0.9.0 README ("List operations") and `AGENT_KNOWLEDGE_BASE.md` say retyped markers keep the old numbering and warn `LIST_FORMAT_NOT_APPLIED`. Checked on the sample's decimal list with a single-paragraph target. This looks like a gap against the documented contract, probably on the single-paragraph or fragment path. It is not filed because Reject All and the multi-paragraph target form were not checked. The add-in avoids it by deriving markers from the target list's kind (`listInsertionMarkers`).

### KB-8: Test XML provider was chosen silently; the `jsdom` lane fails 5 suites (all provider-specific, none a user defect)

- **Where:** [`tests/setup-xml-provider.mjs`](../../tests/setup-xml-provider.mjs). `d5809b9` (the issue #2 fix) added `jsdom` as a dev dependency, and the setup then preferred it whenever it was importable.
- **Resolved (2026-10-12):** the provider is now explicit. The default is `@xmldom/xmldom` (48/48 suites pass). `DOCX_TEST_XML_PROVIDER=jsdom` selects `jsdom` and fails loudly if it is missing. Unknown values throw. [`TESTING.md`](../TESTING.md) documents it. Since 2026-10-13 the default lane configures nothing and uses the library's own xmldom fallback, with no test-installed DOM globals.
- **Production environment:** the add-in never calls `configureXmlProvider`, so the library uses the host's native `DOMParser`/`XMLSerializer`. In Word that is the WebView2 (Chromium) DOM. Every failure below was re-run against real Chromium (Microsoft Edge 155, headless) through the library's browser bundle (`dist/docx-redline.bundle.js`, 0.9.0), alongside xmldom and `jsdom` in Node. Same scenarios, same inputs.
- **Root cause of 4 of 5 failures (class B, `jsdom` deviation):**
  - When `jsdom` serializes a subtree, an attribute whose namespace is not declared on that subtree gets a generated prefix (`w14:paraId` becomes `ns1:paraId`, `r:id` becomes `ns1:id`). Chromium and xmldom keep the original prefix.
  - The library serializes paragraph, `pPr`, `rPr` and run fragments and strips every `xmlns:` declaration with a regex before re-embedding them (`pipeline/serialization.js` `stripNamespaceDeclarations`, `pipeline/patching.js`, `engine/rpr-helpers.js`, `engine/table-cell-context.js`). That is only safe when the serializer keeps the prefixes the surrounding document declares, as Chromium does.
  - Under `jsdom` the stripped `ns1` is unbound (parse error), or `r:id` stops being found (see below).
  - Proof that this is the only cause: a scratch shim that maps generated prefixes back to `w14`/`r` makes all four suites pass under `jsdom`.
- **Per-failure classification:**

| Suite | Class | Evidence |
| --- | --- | --- |
| `browser_document_session_tests` | B | A 2-op batch (redline + comment) on the Word-authored fixture, whose paragraphs carry `w14:paraId`, fails with `PACKAGE_OPERATION_FAILED` / unbound `ns1` under `jsdom`. xmldom and Chromium: `ok`, output well-formed, `w14:paraId` preserved. |
| `docx_cross_host_parity_tests` ("portable core") | B | A localized replace in a paragraph with `<w:hyperlink r:id>` is refused with `REVISION_FIDELITY_VIOLATION` under `jsdom` (`r:id` serialized as `ns1:id`, so the rel id reads back empty: `hyperlink:rIdLink` vs `hyperlink:`). xmldom and Chromium: `ok`, `r:id="rIdLink"` kept. The guard fails closed, so no corruption results. |
| `ooxml_formatting_visual_tests` | B | Same hyperlink `r:id` refusal, in the tab/line-break/hyperlink paragraph (`combined`, line 186). |
| `addin/integration_tests` ("Parties to Table") | B | Text-to-table over two NDA paragraphs: `PACKAGE_OPERATION_FAILED`, unbound `ns1` (the later `w:p` elements use `ns1:paraId` without a declaration). xmldom and Chromium: `ok`, Accept All resolves. |
| `phase4/golden-guardrail` | C | Correct output, provider-specific bytes. `table_reconcile`: xmldom and Chromium emit redundant `xmlns:w` on imported children, and their output is byte-identical after id/date normalization (3085 chars). `jsdom` omits them (2304). `comment_package_injection`: xmldom keeps the XML declaration plus a newline, Chromium keeps the declaration without the newline (1 char shorter), and `jsdom` drops the declaration. All are equivalent XML. The baseline is an xmldom-byte baseline, so any other provider (Chromium included) mismatches on `comment_package_injection`. Hashes were not changed. |

- **Class A (library defects that also occur in Chromium):** none found. No `docs/library-issues` report was filed.
- **Library robustness (optional, not a defect):** the `xmlns:` strip regexes assume the provider keeps prefixes. Re-using the original prefixes explicitly, or checking that every prefix in a stripped fragment is bound in the target context, would make the library tolerant of spec-style serializers such as `jsdom`'s. Only worth doing if a `jsdom`/Node-based consumer matters. It is not a user issue.
- **User impact:** none observed. Production runs on Chromium, where all five scenarios pass. Not tested: Word hosts that fall back to the legacy IE11/EdgeHTML webview, which this add-in does not target.
- **Exit condition (recommendation):**
  - Do not make the `jsdom` lane a required gate. It reports false failures, and a fix would force either a library change that no user needs or a prefix shim in the tests.
  - Keep xmldom as the required gate. Replace `jsdom` as the browser-like signal with a real-browser lane: headless Chromium or Edge loading `dist/docx-redline.bundle.js` and running the isomorphic scenarios (multi-op batch on a `w14:paraId` fixture, hyperlink `r:id` edit, text-to-table on the NDA sample, accept/reject). [`scripts/browser-document-validation.html`](../../scripts/browser-document-validation.html) is a manual starting point. Make it advisory first, then required once stable in CI.
  - Until that lane exists, keep `DOCX_TEST_XML_PROVIDER=jsdom` as an opt-in, non-gating diagnostic. Its 5 known failures are classified above.
  - Keep `jsdom` as a devDependency: `chat_ui_security_tests` and `model_settings_tests` use it for UI DOM tests, independent of the XML provider.

## Word JS API usage (host boundary)

Inventory of `src/` at `a707b85` plus the 0.9.0 working tree. Dev-only entries are excluded: `scripts/officejs-validation.js` and `scripts/golden/golden-driver.js`.

All document *transformation* happens in the library. The only Word call that changes content is a whole-body `insertOoxml(..., 'Replace')`. The tables below give every use, why it exists, and its exit condition.

**Necessary:** the library cannot reach the live document. **Expedient:** the library could do it, or the use is leftover.

### Source read (Flat OPC body)

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `chat/word-document-source.js` | `body.getOoxml()` | Library inspection needs the live body package | Necessary | Only a Word-supported way to read the document package directly |
| `docx-redline-js-integration/word-operation-runner.js` | `body.getOoxml()` | Source for the single batch | Necessary | Same |
| `chat/document-tools.js` | `getOoxml` (wrapped) | Freshness re-read after checkpoint, before commit | Necessary | Same |

### Commit

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `word-operation-runner.js` | `body.insertOoxml(payload, 'Replace')` | One whole-body write of the library-prepared package | Necessary | Only a different host API. The complete-DOCX facade and binary bridge were rejected by the owner |

### Tracking mode save, disable and restore

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `docx-redline-js-integration/word-ooxml.js`, `word-operation-runner.js`, `chat/checkpoint-restoration.js` | `document.changeTrackingMode` load, set to off, restore | Word would wrap the replacement in its own revisions. The library already produced the tracked output | Necessary until disproved | Evidence that `insertOoxml('Replace')` preserves the package's revisions unchanged with tracking on |

### Selection, identity and platform

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `chat/word-document-source.js` | `document.getSelection()`, `load('text')` | "User highlighted text" hint. The library has no selection concept | Necessary | None. A future improvement could map the selection to inspected paragraph handles instead of Word text |
| `taskpane.js` | `Office.context.document.url`, `Office.onReady`, `Office.context.platform` | Document-identity guard, host gate, and the platform passed to the library | Necessary | None |
| `chat/word-document-source.js` → `word-source-identity.js` `detectExportTailIndex` | `body.paragraphs.load('$none')`, `items.length` | Host paragraph count, used only to recognise the empty trailing paragraph Word adds to `getOoxml` exports | **Expedient** | The library recognises the Word export tail from Flat OPC alone, or the extra paragraph is shown never to occur. Then remove the count and `detectExportTailIndex` |

### Verification and checkpoint restore

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `chat/document-tools.js` | `body.getOoxml()` after commit | Evidence of what Word actually committed, compared through library inspection | Necessary | None |
| `chat/checkpoint-restoration.js` | `body.insertOoxml(saved, 'Replace')`, then `getOoxml`, then `paragraphs.load` | Whole-body restore and token comparison | Necessary (the paragraph count is the expedient tail check above) | None for the restore itself |

### Author lookup and glance checks

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `taskpane.js` (~L561) | `document.properties.load('lastAuthor, author')` | Pre-fills the redline author on refresh | Convenience, host-only | Optional. An explicit author setting exists |
| `taskpane.js` (glance checks) | none since WP4 (2026-10-16) | Prompt text is built from library inspection (`inspection.paragraphs[].exactText`) through the shared source reader, not Word's `body.text` | Retired | — |

### Local package logic

This is not Word API use, but it is outside the library:

- **`consumer-core.js` and `word-source-identity.js`:** Flat OPC part and relationship handling, comment and numbering merge, insertion-package assembly, and Word export-noise normalisation. **They stay in the plugin** by owner decision (2026-10-17): the library stays host-independent. Their remaining issue is imports of unexported library internals (see phase-2 WP1).
- Moved by phase-2 WP3 (library 0.10.0): the foreign-revision policy choice (now `existingRevisions: 'auto'`) and the post-commit review-segment comparison (now `compareDocumentInspections`).

### Expedient uses to retire (priority order)

1. Host paragraph count for the export tail: `detectExportTailIndex` needs `hostParagraphCount` to certify that Word's trailing empty paragraph is an export artefact and not real content. Removing it needs Word evidence that the tail is always present, or a Flat OPC marker that identifies it. This is plugin work now that the Word layer stays here.

Retired by phase-2 WP3 (2026-10-17): the `TRACKING_POLICY_MISMATCH` guard (KB-5).

Retired by phase-2 WP4 (2026-10-16): the glance-check `body.text` read, the unreachable checkpoint-capture `getOoxml` fallback in `taskpane.js`, the unused `baseTrackingMode` option, the unimported `src/taskpane/debug_structure.js`, and the Outlook-template `src/commands/commands.js` (the manifest's `FunctionFile` still points at the now script-free `commands.html`).
