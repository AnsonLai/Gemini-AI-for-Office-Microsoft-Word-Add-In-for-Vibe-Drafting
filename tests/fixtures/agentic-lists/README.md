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

The supported matrix now includes ten cases: root/nested bullet insertion,
decimal indentation, insertion at a continuation anchor, insertion within an
independent restarted list, insertion before an existing root through the
candidate range mapper, and noncontiguous marked-header conversion. Continuation
and restart expectations are specified from the source and requested edit, not
copied from engine output. The production insertion selector's deeper-level,
Roman-style and tracking-off cases are covered separately by mocked invocation
tests; their live Word fidelity gates remain open.
