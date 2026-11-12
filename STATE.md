# Project state: Gemini Word add-in

**Updated:** 2026-11-11, at release v2.4.0.0 ([release notes](docs/release/v2.4.0.0.md); previous: [v2.3.0.1](docs/release/v2.3.0.1.md), [v2.3.0.0](docs/release/v2.3.0.0.md)). Open bugs, deferred work and the Word JS API inventory are in [KNOWN_BUGS.md](docs/plans/KNOWN_BUGS.md).

## What the add-in does

The add-in gives Gemini three tools: `inspect_document`, which returns focused text and opaque edit handles; `apply_document_edits`, which applies one typed batch; and `search_web` (since v2.4.0.0), a separate grounded Gemini search for facts the document does not contain.

**Supported edits:**

- text replacement (several per paragraph) and paragraph rewrites;
- paragraph insertion, appending and deletion, where each line of inserted text becomes its own paragraph (since 2026-10-31);
- list item insertion at the first two levels, one or more items at once;
- turning paragraphs into a list, changing list numbering, and continuing an earlier list;
- new bulleted or numbered lists from new text, in two stages (insert plain paragraphs, then convert them; since v2.4.0.0);
- comments, several per paragraph;
- bold, italic, underline, strikethrough and yellow highlight, on several separate pieces of one paragraph;
- table-cell text.

Anything else is refused with no change. A batch with an unsupported edit is refused as a whole; when only part of a request is supported, the assistant either refuses or does that part and says what it left out. The [tool contract](docs/agentic-tool-contracts.md) and the editing guide (v12) give the exact rules.

`@ansonlai/docx-redline-js` 0.11.1 (exact pin) does every document transformation. Office.js only does the following:

- reads the body as Flat OPC;
- writes one prepared package with `insertOoxml('Replace')`, with tracking turned off during the write;
- re-reads the body to verify what Word saved.

## Safety

- **Before each write:** the add-in saves a checkpoint of the exact source and re-checks that the document hasn't changed. Revert restores the checkpoint and verifies the restore. If the document changed after the edit, Revert asks first ("Revert anyway" or "Keep document"); it is refused while a request runs.
- **After each write:** the result is checked against Word's saved output. A refused, uncertain or unverified write is reported as it is and is never replayed.
- **Agent loop limits:**
  - six edit turns, plus up to eight inspection-only turns;
  - two failed edit cycles before a stop that says why;
  - one request for a closing summary if the model ends silently after an edit;
  - a 120-second limit per model call and 300 seconds per request;
  - after a failure that follows an edit or an attempted edit, the request stops with the edit reported and no Retry.

## Evidence for v2.4.0.0 (2026-11-11, on the 0.11.1 pin)

Run on the release tree (release marker v2.4.0.0) unless a line says otherwise.

- Offline suite: 41/41. The count is lower than 48 because library-only tests moved to the docx-redline-js repository on 2026-11-09.
- Real-Word replays: NDA 16/16 steps (53 checks; step 11 is now a supported three-stage edit); privacy notice 16/16 (65 checks).
- Taskpane checks lane (Revert, Cancel, Retry, Refresh, Settings, scripted model): 14 pass, 2 skipped by design.
- Office.js validation: 20 cases passed. Independent desktop Word Accept All/Reject All oracle on its exports: 88/0 (22 not applicable). `npm run test:word` passed.
- Live runs on guide v12, before the release marker changed, one run each: `gemini-flash-latest` NDA 16/16 (53 of 53) and privacy notice 16/16 (65 of 65); `gemini-4-carbon` first 11 NDA requests 35 of 35.
- Not run: a full 16-request live run on the default slow model (`gemini-4-carbon`), and any live run on `gemini-4-argon`.
- Open defects recorded in this cycle: KB-20 ("Revert anyway" discards later changes) and KB-21 (library: delete refused after same-author pending insertions). KB-17 (model judgement) stays open.

## Evidence for v2.3.0.1 (2026-10-31, on the 0.11.1 pin)

- Offline suite: 48/48.
- Real-Word replays: privacy notice 16/16 steps (65 checks, six new steps for multi-paragraph insertion, several formats, replacements and comments in one paragraph); NDA 16/16 (52 checks).
- Live run (privacy notice, configured models): every new capability worked; the two failed checks were wording-strict checks, since loosened and rechecked against the live output.
- Office.js validation: 20 cases passed. Independent desktop Word Accept All/Reject All oracle: 88/0.
- Open library defect found in this release: KB-19 (a replacement drops a pending same-author bold).

## Evidence for v2.3.0.0 (2026-10-30, on the 0.11.1 pin)

- Offline suite: 48/48.
- Office.js validation: 20 cases passed.
- Independent desktop Word Accept All/Reject All oracle: 88/0.
- Real-Word replays:
  - NDA, 16/16 steps (52 checks);
  - privacy notice, 10/10 steps (35 checks). Its fixture has a footnote, and step 04 checks that the footnote survives a rewrite.
- Live runs: slow-model NDA (gemini-2.5-pro) passed 45 of 52 checks, and the privacy notice passed 34 of 35. The add-in bugs found were fixed before release. The model-judgement gaps are tracked as KB-17.

Not certified: Word Online, which wasn't re-checked for this release.
