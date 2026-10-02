# Installed 0.8.3 loses tracked formatting when a table is appended

> **Resolved:** Fixed in 0.8.4 (verified 2026-10-02 offline and in desktop Word Accept All: underline kept, 3×3 table present). The public-library script now asserts the corrected behavior and exits successfully on 0.8.4. Word's Reject All extra paragraph remains open: [Word Reject All report](../2026-10-01-word-reject-table-append-paragraph.md).

**Consumer scope update (2026-10-02, WP5):** The direct Word agent now refuses
structural table creation and has no legacy planner or special
`UNSUPPORTED_TABLE_FORMATTING` command path. Consumer behavior below is
historical incident evidence. The independent repro remains a regression guard;
retiring the consumer route does not change the separate open Word incident.

The add-in planner represents a P8 table append as a replacement of the final source paragraph, P7, with the retained paragraph text followed by the Markdown table. A plain P7 edit and P8 table append can be combined into one source-relative operation; that path retains the P7 text and creates all nine table cells. The installed engine's Reject All resolver restores the original seven paragraphs, but Word's native Reject All leaves an extra empty paragraph; see the separate [Word report](../2026-10-01-word-reject-table-append-paragraph.md).

The installed `@ansonlai/docx-redline-js` 0.8.3 engine has two separate fidelity failures around inline formatting:

- If P7 is underlined in one same-author batch, then a later P8 table append is accepted and committed, but Accept All removes P7's underline. The installed engine's Reject All resolver restores the original seven paragraphs and removes the table; Word's native Reject All leaves one additional empty paragraph after the original text. See [the separate Word Reject All report](../2026-10-01-word-reject-table-append-paragraph.md).
- If a new Markdown underline and table are sent together in one replacement string, the engine reports success and creates the table, but Accept All does not apply the requested underline.

The add-in refuses formatting-plus-table combinations with `UNSUPPORTED_TABLE_FORMATTING` before attempting a Word write. It does not change the installed library. Plain-text paragraph edits plus a single appended Markdown table are still coalesced into one canonical operation, and both AI changes map to that operation's engine receipt; Word's native Reject All fidelity issue remains open for that plain append path.

Run the standalone reproduction from the repository root:

```powershell
node docs/library-issues/completed/2026-10-01-table-append-tracked-formatting-reproducer.mjs
```

The standalone script uses only the public `applyOperationsToDocumentXml`, inspection, Accept All, and Reject All APIs. It guards the 0.8.4 fix: same-author table append keeps the earlier underline on Accept All, Reject All restores the original seven paragraphs, and a new underline combined with the table is retained. It exits successfully when these fixed behaviors hold and prints both resolved views. The former consumer-planner reproduction has been retired because it depended on add-in-only behavior rather than an independently callable library API.
