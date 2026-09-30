# v0.8.2 upgrade validation

Date: 2026-09-30. Supersedes the open library-defect findings in the
[v0.8.1 report](2026-09-30-docx-redline-js-v0.8.1-wp6.md).

Both the root add-in and MCP server now pin and install exact
`@ansonlai/docx-redline-js@0.8.2`. Each install changed only the library package.
Published npm integrity:
`sha512-0j/y9jko8G0F0jT5hCBNjVHH2xN0mbYxnnOEEZgS1ETWAmoeZgWFQB5Bt7cOrBzfWfRZuvniYeMJCnUeoLUkFA==`.

No host operation/schema migration or consumer workaround was needed. The
existing v0.8.1 API compatibility suite continues to pass on 0.8.2; the release
notes retain CLI contract version 8.

## Checks

| Check | Result |
| --- | --- |
| Both `npm ls` inventories | Exact 0.8.2 |
| `npm test` | 33 suites passed, 0 failed |
| Portable library issue reproducer | Both defects now pass |
| `npm run test:word` | 69 live Word checks passed, 0 failed |
| Development build | Passed |
| Production build | Passed; three existing bundle-size warnings, taskpane approximately 747 KiB |
| Existing golden guardrail | Passed with unchanged hashes |

Four non-offline entrypoints remain explicitly excluded from `npm test`.
The historical v0.5.4 suite reports its version-specific package behavior as
skipped while its consumer guards pass. Build telemetry requests remain blocked;
webpack succeeds without Node polyfill errors.

## Required fidelity regressions

Library issues [#3](https://github.com/AnsonLai/docx-redline-js/issues/3) and
[#4](https://github.com/AnsonLai/docx-redline-js/issues/4) are fixed in the library.
The default consumer fidelity suite now requires **four** cases: localized
`replacements` and full-paragraph `modified` forms of each edit. These are no
longer optional known failures or excluded from live Word testing.

Independent assertions require:

- Source `example.org` hyperlink plus plain `.` accepts to a hyperlink displaying
  only `example.net`, with the period still outside it. Rejection restores the
  original boundary; the hyperlink relationship remains unchanged.
- Source `\ttabbed\nLine with ` accepts to `\taligned\nLine with ` and rejects
  exactly to source, including the tab, manual break and trailing space.
- Package validation succeeds and Word recognizes revision markup.

Expected text is built from fixed source literals and the requested replacement,
not from the library's own resolver output. Generated packages include tracked,
engine-accepted and engine-rejected views for independent Word comparison.

## Real desktop Word

Word version `16.0`, build `16.0.20326`. Command:

```powershell
npm run test:word -- -ArtifactsDir .cache/wp6/v082/live-word -TimeoutSeconds 120
```

The 69 checks comprise ten fixtures × six differential checks, two native
insertion roundtrips × four checks, and the invalid-content-type negative control.
Word's own Accept All/Reject All and its hyperlink ranges confirm both fixes in
both operation forms. Engine-resolved packages are checked for leftover revisions
before Word performs any additional resolution.

Existing live checks also cover formatting, font inheritance, tab alignment,
section orientation/columns, first-page footer content, thread parent identities
and resolved state. Both actual production-bridge native insertion roundtrips
still pass: plain replacement and threaded-comment reply. Word exports live
scope XML, the bridge produces one payload, native `InsertXML` imports it, and
the saved package reopens with correct tracked/accepted/rejected states.

Detailed reports and generated fixtures remain ignored under `.cache/wp6/v082/`.
No v0.8.2 golden baseline refresh was necessary.

## Remaining scope

The dependency upgrade and both library-defect gates are verified. Actual
Office.js `insertOoxml` dispatch remains a separate outstanding host integration
check; COM native insertion does not certify that transport. PDF export remains
optional diagnostic evidence. The earlier measured 1,000-paragraph performance
is accepted by the user and was not re-benchmarked for this patch upgrade.
