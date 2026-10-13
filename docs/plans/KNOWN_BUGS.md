# Known bugs and host-boundary exceptions

**Updated:** 2026-10-12, after the `@ansonlai/docx-redline-js` 0.9.0 upgrade
([validation](../validation-reports/2026-10-12-docx-redline-v090.md)).

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

### KB-4: List conversion and list-format changes are refused despite library support

- 0.9.0 ships `list-format` (`convert-plain`, `change-format`, `join`) and `paragraph-format.numbering`. It also supports a composed `redline` plus `list-format` for changed-text headers, and returns explicit refusal codes.
- The agent's typed intents expose none of these. Header-to-list conversion and numbering-style changes are refused before writing.
- Tracked in the [canonical list migration follow-up](2026-09-30-canonical-list-migration-follow-up.md) (WP1 to WP3).
- **Exit condition:** that plan's acceptance criteria are met.

### KB-5: `TRACKING_POLICY_MISMATCH` guard is a 0.8.4 workaround

- [`document-tools.js`](../../src/taskpane/modules/chat/document-tools.js) (`onPrepared`) refuses an untracked batch whose prepared output carries new own-author revisions. It was added for the 0.8.4 surgical-insertion bug.
- 0.9.0 fixes that bug and enforces the same rule itself (`REVISION_FIDELITY_VIOLATION` / `UNTRACKED_CREATED_REVISION`).
- The guard is kept as a backstop and no longer fires on the certified cases.
- **Exit condition:** one release cycle on 0.9.x with the untracked capability gate passing. Then remove the guard and rely on the library's `UNTRACKED_CREATED_REVISION`.

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
| `taskpane.js` (checkpoint capture fallback, ~L925) | `body.getOoxml()` | Fallback when no source is supplied. The only caller always supplies one, so this branch looks unreachable | Expedient (dead) | Delete the branch |

### Commit

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `word-operation-runner.js` | `body.insertOoxml(payload, 'Replace')` | One whole-body write of the library-prepared package | Necessary | Only a different host API. The complete-DOCX facade and binary bridge were rejected by the owner |

### Tracking mode save, disable and restore

| Location | Calls | Why | Class | Exit condition |
| --- | --- | --- | --- | --- |
| `docx-redline-js-integration/word-ooxml.js`, `word-operation-runner.js`, `chat/checkpoint-restoration.js` | `document.changeTrackingMode` load, set to off, restore | Word would wrap the replacement in its own revisions. The library already produced the tracked output | Necessary until disproved | Evidence that `insertOoxml('Replace')` preserves the package's revisions unchanged with tracking on |
| `word-ooxml.js` `baseTrackingMode` parameter | none | No caller passes it | Expedient (dead) | Remove the parameter |

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
| `taskpane.js` (~L804, glance checks) | `body.load('text')`, `body.text` | Prompt text for glance checks. This is **Word's text, not library text**: it differs on soft breaks and deleted runs | **Expedient** | Build the text from `inspectWordSource(...).inspection.paragraphs[].exactText` through the shared source reader |

### Local package logic that should move upstream

This is not Word API use, but it is outside the library:

- **`consumer-core.js`:** Flat OPC part and relationship handling, comment-part merge (`mergeCommentPartsIntoPackage`), and `buildBatchInsertionPackage`. Exit: a supported library Flat OPC inspect/prepare surface. The plan already proposes one; its API is not yet agreed.
- **`word-source-identity.js`:** normalisation of Word rsid and paraId export noise, and export-tail detection. Exit: the same Flat OPC surface, or a library "Word export" normaliser.
- **`chat/document-tools.js`:** choosing `existingRevisions: 'slice-cross-author'` from the revision authors in a paragraph. Exit: a library auto policy for foreign revisions.
- **`chat/document-tools.js`:** the local review-segment merge used for post-commit verification. Exit: a library helper that verifies committed output against the prepared output.

### Expedient uses to retire (priority order)

1. Glance-check `body.text` read: switch to library inspection text.
2. Host paragraph count and `detectExportTailIndex`: needs a library export-tail detector.
3. `TRACKING_POLICY_MISMATCH` guard (KB-5).
4. Dead code: the `taskpane.js` checkpoint-capture fallback, the `baseTrackingMode` parameter, the unimported `src/taskpane/debug_structure.js` (`Word.run` + `paragraphs.load`), and the Outlook-only `src/commands/commands.js` template stub.
5. Upstream the `consumer-core.js` / `word-source-identity.js` package logic once a library surface is agreed.
