# Word-authored list fixture

`nested-lists-source.docx` was authored and save/reopened with Microsoft Word
desktop COM. It contains 24 body paragraphs, including a nine-level bullet
definition, nested decimal numbering, a continuation across plain paragraphs,
a separate restart instance, unmarked header candidates and manually marked
headers. The first paragraph is an unchanged bold-formatting sentinel.

`source-observations.json` records the independent Word COM build and the
source's list levels, labels, values, direct OOXML `numId` values and numbering
formats. Tests use those observations instead of inferring expected source
numbering from the library under test.

The current direct document-tool lane retains four tracked list insertion
cases: same-level insertion into a root bullet, level-1 decimal, a continued
decimal sequence and an independent restart. It also checks a tracking-off
insertion with tracking restored. The manifest preserves the original fixture,
expected source/accepted/rejected views, numbering, labels, formatting sentinel
and comments. `tests/document_edit_intents_tests.mjs` covers compiler boundaries;
`tests/direct_document_fidelity_manifest_tests.mjs` emits the Office.js cases.

```powershell
node tests/direct_document_fidelity_manifest_tests.mjs --export-dir .cache/reliability/direct-word-document-tools
node scripts/run-officejs-validation.mjs --launch --artifacts-dir .cache/reliability/direct-word-document-tools/officejs
```

The old native fallback route and its five-case generator were retired on
2026-10-02 when the direct typed list subset replaced that production route.
The September 30 Word and Office.js native-fallback reports are historical
evidence only; they do not gate the current direct tool path. The old general
list planner matrix was also retired. Package list generation independent of the add-in route is tested in the
library repo (`tests/same_text_list_conversion_tests.mjs`).

To recreate the Word-authored source fixture on Windows with Word installed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/create-agentic-list-fixtures.ps1
```

Generation logs and bounded-worker progress belong under
`.cache/reliability/agentic-list-fixture-generation`; only the source DOCX and
its observation record belong in this fixture directory.

## List-action sources

`scripts/create-agentic-list-action-fixtures.ps1` authors six small sources through Word COM
(`list-headers`, `list-letter-headers`, `list-bullets`, `list-two-lists`, `list-changed-header` and
`list-ambiguous-marker`), saves and reopens each, and records Word's observations (text, list state,
`ListString`, level, numId and numbering format) in `list-actions-source-observations.json`. They back the
`convert_to_list`, `change_list_format` and `continue_list` Office.js and Word-oracle cases.
