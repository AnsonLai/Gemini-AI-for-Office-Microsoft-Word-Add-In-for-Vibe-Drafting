# Project state: Gemini Word add-in

**Updated:** 2026-10-31, at release v2.3.0.1 ([release notes](docs/release/v2.3.0.1.md); previous: [v2.3.0.0](docs/release/v2.3.0.0.md)). Open bugs, deferred work and the Word JS API inventory are in [KNOWN_BUGS.md](docs/plans/KNOWN_BUGS.md).

## What the add-in does

The add-in gives Gemini two tools: `inspect_document`, which returns focused text and opaque edit handles, and `apply_document_edits`, which applies one typed batch.

**Supported edits:**

- text replacement (several per paragraph) and paragraph rewrites;
- paragraph insertion, appending and deletion, where each line of inserted text becomes its own paragraph (since 2026-10-31);
- list item insertion at the first two levels, one or more items at once;
- turning paragraphs into a list, changing list numbering, and continuing an earlier list;
- comments, several per paragraph;
- bold, italic, underline, strikethrough and yellow highlight, on several separate pieces of one paragraph;
- table-cell text.

Anything else is refused with no change, and a mixed request with an unsupported part is refused as a whole. The [tool contract](docs/agentic-tool-contracts.md) and the editing guide (v10) give the exact rules.

`@ansonlai/docx-redline-js` 0.11.1 (exact pin) does every document transformation. Office.js only does the following:

- reads the body as Flat OPC;
- writes one prepared package with `insertOoxml('Replace')`, with tracking turned off during the write;
- re-reads the body to verify what Word saved.

## Safety

- **Before each write:** the add-in saves a checkpoint of the exact source and re-checks that the document hasn't changed. Revert restores the checkpoint and verifies the restore.
- **After each write:** the result is checked against Word's saved output. A refused, uncertain or unverified write is reported as it is and is never replayed.
- **Agent loop limits:**
  - six edit turns, plus up to eight inspection-only turns;
  - two failed edit cycles before a stop that says why;
  - one request for a closing summary if the model ends silently after an edit;
  - a 120-second limit per model call and 300 seconds per request.

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
