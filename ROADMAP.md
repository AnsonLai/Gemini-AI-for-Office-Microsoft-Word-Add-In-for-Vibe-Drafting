# Project roadmap

**Updated:** 2026-10-30, after v2.3.0.0 shipped on `@ansonlai/docx-redline-js` 0.11.1 (exact pin). [STATE.md](STATE.md) describes the current scope.

## Direction

Keep the product an in-place Word add-in with a small inspect/apply agent loop, and keep shrinking it. The library owns every document transformation. The add-in keeps three thin parts: the taskpane UI and settings, the agent loop, and the Word host adapter, which reads the body as Flat OPC, writes one prepared package, and verifies what Word saved. Word-specific glue (Flat OPC handling and Word export noise) stays in the add-in, and the library stays host-independent.

A new document capability arrives as a library release plus a thin intent mapping, never as add-in logic. A library fix permits a tested version upgrade; it does not by itself widen what the assistant offers.

## Done

- **Phase 1, library-first lightweight harness** (closed 2026-10-12): [plan](docs/plans/completed/2026-10-02-library-first-lightweight-agent-reliability.md).
- **Phase 2, library-first simplification** (closed 2026-10-30, shipped in v2.3.0.0): [plan](docs/plans/completed/2026-10-12-library-first-simplification-phase-2.md). Production code is 4,362 lines in 19 files, and `src/` imports only the library root and `/standalone-runner`.

## Open and deferred

Everything still open is in [KNOWN_BUGS.md](docs/plans/KNOWN_BUGS.md):

- KB-1, body-final table append. It isn't exposed, but it blocks bringing table append back.
- KB-17, live-model judgement gaps.
- The deferred list: automatic model fallback (issue #3), Word Online, the last expedient Word API use, and lint.

## Not scheduled

These stay refused unless the owner approves a capability and it passes the source, Accept All and Reject All checks in Word:

- table creation and rows;
- list depth, indent and range changes;
- font and layout controls;
- revision acceptance.

Also not planned: a complete-DOCX Word facade or binary bridge, and a third-party harness (the [harness plan](docs/plans/2026-10-12-harness-optional.md) is parked).

Earlier milestones are kept in [completed plans](docs/plans/completed/) and dated [validation reports](docs/validation-reports/).
