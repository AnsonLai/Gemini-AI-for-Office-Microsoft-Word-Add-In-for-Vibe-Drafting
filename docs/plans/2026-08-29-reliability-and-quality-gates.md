# Reliability, Quality Gates, and Accuracy Verification Plan

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** Partially implemented by the upgrade; remaining work ready to start.

**Recommended order:** 1 of 4 August plans.

## Baseline and priorities

Both consumers pin exact `@ansonlai/docx-redline-js@0.8.2`. The [September upgrade](2026-09-09-docx-redline-js-v0.5.4-upgrade.md) is **near complete**, with actual Office.js insertion verification outstanding. This plan owns ongoing quality-gate maintenance; record that final upgrade check in both plans when performed.

The [current validation report](../validation-reports/2026-09-30-docx-redline-js-v0.8.2.md) records 33 offline suites, 69 live Word checks and successful development/production builds. These are dated results, not a guarantee that future changes pass.

1. Exact targeting, independent Accept All/Reject All expectations, preserved formatting/relationships, and fail-closed engine failures.
2. Portable OOXML core tests with host tests in separate lanes.
3. Measured performance after correctness. The user accepts the observed slower 1,000-paragraph workload.

Offline tests must not require Word, Office.js or provider credentials. Transport adapter tests may use mocks; mocks do not certify actual Word behavior. Real Word remains a required host-validation lane for relevant changes. PDF export is optional.

## Completed / inherited work

| Original work | Current state | Evidence |
| --- | --- | --- |
| WP1: unified offline runner | Implemented; maintain rather than recreate | `scripts/run-all-tests.mjs`, `npm test` |
| Fidelity and revision regressions | Implemented, including both v0.8.2 fixes | `tests/ooxml_formatting_visual_tests.mjs` |
| Live Word lane | Implemented: supervised COM, production bridge, native insertion, save/reopen and revision resolution | `scripts/verify-wp6-word.ps1`, `npm run test:word` |
| Golden provenance | Reviewed/corrected; verification does not rewrite baselines | `tests/phase4/golden-guardrail.mjs` |
| Generic error/result propagation | Present in batch bridge and MCP; broader tool/UI consistency needs audit | Result contract, batch and MCP suites |

Four entrypoints are explicitly excluded from the offline runner, and historical version-specific skips are reported. Do not relabel excluded/skipped checks as passes.

## Remaining work packages

### WP1 — Close actual Office.js transport gate and maintain the lanes

- Exercise the production add-in through real `getOoxml` / `insertOoxml`, including localized body replacement and a threaded-comment package.
- Check no repair prompt, exact accepted/rejected text, comment identities and successful write reporting after save/reopen.
- Check no write on no-op or failed preparation; count one successful scope read and one successful insertion per batch. Tracking-mode management may require additional synchronization.
- Record Word/build, operations, artifacts and results in a validation report; close September WP6 only after this evidence exists.
- Document repeatable offline, live Word, provider-evaluation and observational benchmark commands. Extend discovery/exclusions deliberately as suites are added.

### WP2 — Centralize provider requests and bounded recovery

Provider requests still appear in multiple paths in `commands/agentic-tools.js`; `chat/gemini-client.js` has not been implemented.

- Inventory taskpane, browser-demo and evaluation callers before choosing the shared client boundary.
- Centralize request construction, cancellation, timeout cleanup, response parsing and bounded transient HTTP/network retries with jitter.
- Keep the configured model; withdraw the old plan's hardcoded model default.
- Separate provider-request retries from edit recovery. Never replay an applied mutation because a later response/UI step failed.
- Authentication, invalid requests, ambiguous targets and unsafe engine boundaries require actionable failure. At most one context refresh for a stale target, followed by revalidation.
- Test cancellation, timeout, transient recovery and exhausted retries with deterministic stub responses; live provider evaluation stays opt-in.

### WP3 — Consistent mutation outcomes and diagnostics

- Audit tool/UI boundaries for preservation of actual library status, errors, receipts and bridge `written` state.
- Distinguish prepared engine success, successful host insertion, no-op, engine rollback and host-write failure. Engine atomicity does not prove rollback after a failed host write.
- Preserve unknown error codes. Establish policy from the installed package and exercised cases; the old fixed error table was not a complete retry contract.
- Do not invent receipt fields such as `revisionIds`, `disposition` or `validationSummary` unless supplied by the library or derived by a documented consumer adapter.
- Keep diagnostics free of API keys and document text by default; expose concise actionable errors.

## Acceptance

- [x] Offline runner and live Word harness exist; v0.8.2 validation recorded.
- [ ] Actual Office.js transport gate has a reproducible passing report.
- [ ] Provider cancellation/retry policy centralized and tested without replaying mutations.
- [ ] Tool/UI outcomes distinguish engine and host completion, including unknown errors.
- [ ] Required checks and optional lanes have documented execution criteria.

**Next:** [Agentic tools and list reliability](2026-08-29-agentic-tools-and-list-reliability.md), after host and outcome policies are established.
