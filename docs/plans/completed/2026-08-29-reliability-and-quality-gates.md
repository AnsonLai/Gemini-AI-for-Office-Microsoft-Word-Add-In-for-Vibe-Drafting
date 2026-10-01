# Reliability, Quality Gates, and Accuracy Verification Plan

> **Retrospective ownership review (2026-09-30):** See the [library offload review](../../library-offload-review.md) for measured code changes, responsibilities delegated to the library, and remaining consumer/native routes. Version pins and validation counts in this closure record describe their dated checkpoints; current consumers pin exact 0.8.3. All five original plans are closed; remaining canonical list work is in the [fresh follow-up](../2026-09-30-canonical-list-migration-follow-up.md).

**Date:** 2026-08-29

**Last Updated:** 2026-09-30

**Status:** Complete — WP1–WP3 implemented and verified on 2026-09-30. Ongoing maintenance and broader tool/list coverage continue in the next plan.

**Recommended order:** 1 of 4 August plans.

## Baseline and priorities

Both consumers pin exact `@ansonlai/docx-redline-js@0.8.2`. The [September upgrade](2026-09-09-docx-redline-js-v0.5.4-upgrade.md) is **complete**: its actual Office.js insertion gate passed during this plan. This plan owns ongoing quality-gate maintenance.

The [current validation report](../../validation-reports/2026-09-30-docx-redline-js-v0.8.2.md) records 33 offline suites, 69 live Word checks and successful development/production builds. These are dated results, not a guarantee that future changes pass.

The [reliability validation report](../../validation-reports/2026-09-30-reliability-and-quality-gates.md) records the final 35-suite offline run, seven actual Office.js checks, 12 independent Word package/revision checks, both builds, outcome semantics and reproduction commands. Prior upgrade/plan changes were committed as `68da8a4` before implementation began; reliability work is saved as a separate milestone commit before starting the agentic-tools plan.

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
| Generic error/result propagation | Audited and extended across redline/comment/highlight/native mutating tools and UI | Result contract, batch, mutation-outcome and MCP suites |

Four entrypoints are explicitly excluded from the offline runner, and historical version-specific skips are reported. Do not relabel excluded/skipped checks as passes.

## Work packages and implementation records

### WP1 — Close actual Office.js transport gate and maintain the lanes

**Complete.** Added optional webpack validation entry, local collector and external-manifest oracle support. Word `16.0.20326.20158` passed seven genuine Office.js checks and 12 independent reopen/no-repair/revision checks. Body and threaded-comment edits each used one real scope read/insertion; unchanged paragraph, empty batches and missing targets inserted nothing. The temporary validation registration was removed. Commands and lane criteria are documented in `scripts/README.md`.

- Exercise the production add-in through real `getOoxml` / `insertOoxml`, including localized body replacement and a threaded-comment package.
- Check no repair prompt, exact accepted/rejected text, comment identities and successful write reporting after save/reopen.
- Check no write on no-op or failed preparation; count one successful scope read and one successful insertion per batch. Tracking-mode management may require additional synchronization.
- Record Word/build, operations, artifacts and results in a validation report; close September WP6 only after this evidence exists.
- Document repeatable offline, live Word, provider-evaluation and observational benchmark commands. Extend discovery/exclusions deliberately as suites are added.

### WP2 — Centralize provider requests and bounded recovery

**Complete.** `chat/gemini-client.js` now centralizes all eight request paths in the taskpane, specialized tools, browser demo and eval harness. `tests/gemini_client_tests.mjs` covers bounded retries, authentication/invalid-request refusal, cancellation, timeout through response parsing, abortable backoff and cleanup. Configured models/settings are retained. No live provider calls were made.

HTTP-400 history-reset recovery was removed. Engine refusal stops immediately; only pre-engine malformed model proposals may receive one corrective generation. No automatic document-context refresh is added; stale targets fail closed, within the at-most-one-refresh budget.

- Inventory taskpane, browser-demo and evaluation callers before choosing the shared client boundary.
- Centralize request construction, cancellation, timeout cleanup, response parsing and bounded transient HTTP/network retries with jitter.
- Keep the configured model; withdraw the old plan's hardcoded model default.
- Separate provider-request retries from edit recovery. Never replay an applied mutation because a later response/UI step failed.
- Authentication, invalid requests, ambiguous targets and unsafe engine boundaries require actionable failure. At most one context refresh for a stale target, followed by revalidation.
- Test cancellation, timeout, transient recovery and exhausted retries with deterministic stub responses; live provider evaluation stays opt-in.

### WP3 — Consistent mutation outcomes and diagnostics

**Complete.** Engine receipts remain untouched; Word outcomes separately report `written`, `writeAttempted`, confirmed writes and operation results. Failed synchronization never triggers insertion replay. Tracking-restoration errors preserve earlier confirmed writes. Comment/highlight sequences stop on first failure; native list/table/section commands report observed partial/uncertain writes. The UI stops further editing when inspection is required. Default diagnostics omit raw request/document content and host error stacks. Unknown codes remain in structured results.

Regression coverage includes first-comment success followed by an unknown host failure (third comment never attempted), uncertain native insertion, tracking-restoration failure after committed insertion and one provider request/zero writes on engine refusal. Native commands retain their current architecture; migration to canonical atomic list operations remains in the agentic plan.

- Audit tool/UI boundaries for preservation of actual library status, errors, receipts and bridge `written` state.
- Distinguish prepared engine success, successful host insertion, no-op, engine rollback and host-write failure. Engine atomicity does not prove rollback after a failed host write.
- Preserve unknown error codes. Establish policy from the installed package and exercised cases; the old fixed error table was not a complete retry contract.
- Do not invent receipt fields such as `revisionIds`, `disposition` or `validationSummary` unless supplied by the library or derived by a documented consumer adapter.
- Keep diagnostics free of API keys and document text by default; expose concise actionable errors.

## Acceptance

- [x] Offline runner and live Word harness exist; v0.8.2 validation recorded.
- [x] Actual Office.js transport gate has a reproducible passing report.
- [x] Provider cancellation/retry policy centralized and tested without replaying mutations.
- [x] Tool/UI outcomes distinguish engine and host completion, including unknown errors.
- [x] Required checks and optional lanes have documented execution criteria.

**Next:** [Agentic tools and list reliability](2026-08-29-agentic-tools-and-list-reliability.md), after host and outcome policies are established.
