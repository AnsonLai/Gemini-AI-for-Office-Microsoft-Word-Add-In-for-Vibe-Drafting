# Agentic list fidelity source

`nested-lists-source.docx` is authored and reopened with Microsoft Word desktop
COM. It contains 24 direct body paragraphs, including a true nine-level bullet
definition, nested decimal numbering, a continuation across plain paragraphs,
a separate restart instance, unmarked header candidates, and manually marked
headers. The first paragraph is a bold formatting sentinel.

`source-observations.json` records the Word COM build and observations after
save/reopen. It includes Word's list level, label, value, direct OOXML `numId`,
and numbering-definition format so the tests do not infer the source fixture's
list semantics from the document library under test.

To recreate the source on Windows with Word installed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/create-agentic-list-fixtures.ps1
```

Generation logs and the bounded-worker progress file go under
`.cache/reliability/agentic-list-fixture-generation`; only the DOCX and its
observations belong in this fixture directory. The saved fixture is immutable
during tests.

`node tests/agentic_list_fidelity_tests.mjs` runs the offline fidelity matrix.
Add `--export-host-dir .cache/reliability/agentic-list-fixtures` to export
tracked, accepted, and rejected packages plus `manifest.json` for the Office.js
collector and independent Word validation. `known-defects-manifest.json`
contains separately classified library Reject All defects and is intentionally
not part of the passing host manifest.

The offline matrix includes ten cases: root/nested bullet insertion, decimal
indentation, insertion at a continuation anchor, insertion within an
independent restarted list, insertion before an existing root through the
candidate range mapper, and noncontiguous marked-header conversion. Continuation
and restart expectations are specified from the source and requested edit, not
copied from engine output.

Production `insert_list_item` native-fallback coverage is generated with
the following command:

```sh
node tests/agentic_list_native_fallback_manifest_tests.mjs --export-dir .cache/reliability/agentic-list-native-final
```

Its five Word cases cover a requested deep level, outdent from a prepared deep
source, upper and lower Roman numbering, and insertion with revision tracking
disabled while restoring the prior tracking mode. Each prepared source package
is frozen before the production call. `scripts/officejs-validation.js` verifies
the actual route (one `Paragraph.insertParagraph`, zero `body.insertOoxml`) and
tracking-mode transitions. The independent Word oracle checks source, tracked, Accept All,
and Reject All views against labels and levels declared in the manifest; these
native cases explicitly omit engine-resolved reference packages. All five cases
passed 20 Office.js checks and 20 applicable independent Word checks on Word
16.0.20430.20092. The Word report also records five engine-reference-package
checks as not applicable; source, tracked, Accept All, and Reject All passed
for every case. Reports: [Office.js](../../../docs/validation-reports/2026-09-30-agentic-native-list-officejs.json)
and [Word oracle](../../../docs/validation-reports/2026-09-30-agentic-native-list-word.json).
