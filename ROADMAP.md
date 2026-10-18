# Project roadmap

**Updated:** 2026-10-12. The [library-first lightweight harness plan](docs/plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md)
was closed on 2026-10-12 after the `@ansonlai/docx-redline-js` 0.9.0 upgrade
([validation](docs/validation-reports/2026-10-12-docx-redline-v090.md)); its linked reports and commits are the implementation record. [STATE.md](STATE.md) describes current scope.

Keep the product an in-place Word plugin with a small inspect/apply harness.
The library owns document transformations; the Word adapter reads, commits and
verifies one Flat OPC body batch while preserving tracking. Preserve ordinary
formatting, numbering, comments and other reviewers' revisions. Refuse requests
outside the tested retained actions. Maintain exact matching library pins in
the add-in and MCP.

Before shipping the changed tool protocol/prompt, complete the documented live
fast/slow model gate using the original NDA requests, intended edits, refusal,
turn-count and host-commit evidence. Desktop Word is the initial validated host;
Word Online requires separate affected-operation checks before being advertised
with equivalent editing guarantees. These are validation requirements, rather
than new product features.

No library issues are open as of 0.10.0 (see the [closeout](docs/validation-reports/2026-10-17-library-issues-closeout.md));
their independent reproducers and Word fixtures stay runnable. Remaining
consumer-side bugs and the Word JavaScript API inventory are in
[KNOWN_BUGS.md](docs/plans/KNOWN_BUGS.md). A published library fix permits a tested
version upgrade; it does not authorize expanding scope. Broader list
conversion/depth/styles and table creation/rows remain refused unless the owner
later approves a capability and its source/accepted/rejected fidelity passes.
The [September list follow-up](docs/plans/2026-09-30-canonical-list-migration-follow-up.md)
is open for consumer adoption of the released `list-format` capabilities; the
[table plan](docs/plans/completed/2026-09-30-table-creation-reliability.md) is closed.

Local Flat OPC assembly remains a measured consumer responsibility. A future
supported upstream inspection/preparation API may reduce it if negotiated and
verified. No complete-DOCX Word facade, binary bridge, Pi runtime adoption,
general skills, additional tool catalogue, font/layout controls or checkpoint
branching is scheduled. Browser/MCP facade behavior remains unchanged.

Earlier engine extraction, package integration, reliability and performance
milestones are preserved in [completed plans](docs/plans/completed/) and dated
[validation reports](docs/validation-reports/). Their native-route evidence is
historical and does not imply that those commands remain available.
