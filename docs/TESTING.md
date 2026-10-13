# Consumer verification

The library owns supported document transformations; this repository verifies
request scope, source-bound handles, atomic host transport and truthful results.
Current supported/refused actions are in the [tool contract](agentic-tool-contracts.md).
The [dated plan](plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md)
links frozen baseline and each milestone's actual evidence.

| Lane | Evidence | Limit |
| --- | --- | --- |
| `npm test` | Compiler allowlist and whole-batch refusal; handles/freshness; complete history and budgets; no replay after uncertain outcomes; checkpoint failure; library/public package, browser and MCP regression contracts | Offline mocks do not certify Word transport. Version-specific skips, manual entrypoints and `KNOWN_LIBRARY_DEFECT:` lines are reported separately. The default lane configures no XML provider (the library falls back to its own `@xmldom/xmldom` in Node) unless `DOCX_TEST_XML_PROVIDER=jsdom` selects the browser-like lane (see KNOWN_BUGS KB-8); it is never chosen by which DOM package happens to be installed. |
| Actual Office.js collector | Genuine Word proxies through production inspect/apply tools; source/write/verification counts; no native editing; persisted exact-source IndexedDB checkpoint; tracking restoration; DOCX exports | Engine preparation alone is insufficient. Check exports independently. |
| Independent desktop Word | Save/reopen without repair; source/tracked/Accept All/Reject All text, numbering identity/continuation/restart, comments, inline formatting and table cells; restored body checkpoint | Scoped fixtures and Word build, rather than universal document fidelity. |
| Direct NDA taskpane replay | The original sixteen user prompts through the real UI and loop; intended edits, explicit unchanged refusals, author attribution, paired tool receipts, turns and actual commits | Mocked model output does not establish drafting quality. |
| Live direct NDA gate | Actual configured fast/slow provider proposals, validity/turns, retained edits, refusals, over-editing/regressions and host outcomes | Opt-in credentials and provider usage; mandatory before shipping the changed prompt. |
| Browser/MCP/public package | Existing open/inspect/atomic edit/save/reopen and cross-host parity | Separate consumers; their facade does not certify Word body transport. |
| Benchmark/startup | Observational direct compile/preparation and package phases; deferred module loading | No universal latency threshold or correctness claim. |

Use existing lanes and Word-authored fixtures, with expectations independent of
the engine under test. A retired capability is an explicit unchanged refusal,
rather than a renamed passing edit. Native-route September reports remain
historical. The smaller retained subset needs fresh direct-tool evidence.
Checkpoints cover body Flat OPC, rather than complete-document backups.

The initial WP4 redline list recipe failed independent accepted list identity.
The documented complete `list-change` fixed it and passed fresh Word checks.
The [WP4 report](validation-reports/2026-10-02-lightweight-harness-wp4.md)
records 11 collector cases, 44 applicable Word checks and 49 NDA replay checks.
The [final WP5 report](validation-reports/2026-10-02-lightweight-harness-wp5.md) records reruns after bridge thinning and actual checkpoint
restoration. Excluded/inapplicable checks are never added to pass totals.

The earlier table-append and Undo incidents retain dated evidence in their
[incident plan](plans/completed/2026-09-30-table-creation-reliability.md),
[Word report](validation-reports/2026-10-01-table-creation-reliability.md) and
[library issue index](library-issues/README.md) (no open issues as of 0.9.0; remaining consumer bugs in [KNOWN_BUGS.md](plans/KNOWN_BUGS.md)). Structural table creation
is now refused by the Word agent. Keep independent library reproducers runnable;
retirement does not close upstream bugs. Desktop is the initial target;
Word Online requires separate affected-operation checks before equivalent
editing guarantees are advertised.

Commands, artifacts and isolation/cleanup behavior are documented in
[scripts/README.md](../scripts/README.md). Provider keys are supplied through
the environment and never written to model logs. PDF export is an optional
diagnostic, not a visual certification substitute.
