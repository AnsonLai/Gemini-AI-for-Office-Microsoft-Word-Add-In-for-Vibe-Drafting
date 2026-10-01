# Reliability and quality gates validation

Date: 2026-09-30. Baseline commit: `68da8a4`; both consumers use exact
`@ansonlai/docx-redline-js@0.8.2`.

## Changes

- Added a development-only actual Office.js lane using the production atomic
  batch bridge and genuine Word proxies. Kept native COM validation separate.
- Centralized Gemini HTTP requests across taskpane, specialized tools, browser
  demo and live evals. Preserved configured models and request settings.
- Bounded transient transport retries to three attempts; diff generation keeps
  its single-request timeout ceiling. Authentication/invalid requests and invalid
  JSON fail without transport retry. Cancellation and timeout cover parsing and
  backoff, with timer/listener cleanup. Transport retries never execute tools.
- Removed HTTP-400 history-reset recovery and raw prompt/provider response logs.
- Separated engine receipts from host observations; removed insertion replay
  after a failed Word synchronization. A confirmed write followed by tracking
  restoration failure remains a confirmed write.
- Added conservative outcomes for native mutating commands and stop-after-failure
  behavior for comment/highlight sequences. These commands have not all migrated
  to one atomic batch; their outcome contract reports that limitation.

## Mutation outcome contract

| Outcome | Meaning | Recovery |
| --- | --- | --- |
| `noop` | No document insertion was needed | No edit retry |
| `refused` / `rolled_back` | Validation/engine rejected before host insertion | Report code; preserve source |
| `prepared` | XML prepared but package/host preparation failed | Report failure; no claimed application |
| `applied` | Host synchronization confirmed insertion | Do not replay |
| `indeterminate` | Host insertion attempted but not confirmed | Stop tools; inspect Word |
| `partial` | Sequence stopped after earlier confirmed writes | Stop tools; inspect Word |
| `applied_with_host_error` | Confirmed write followed by a host error | Stop tools; preserve confirmed outcome |

`written`, `writeAttempted` and `confirmedHostWrites` are consumer observations;
they are not fabricated engine receipt fields. Library receipts retain their
original meaning. Unknown engine/host error codes remain available in the
structured results. No adapter claims host rollback on an unconfirmed write.

## Actual Office.js methodology

The local collector uses existing trusted development certificates and a separate
validation add-in identity. It loads a disposable document without provider
credentials. Fixture seeding and compressed-file export are harness operations;
the production edit uses `getOoxml` → atomic library batch → `insertOoxml`.

Cases cover a localized `beta` → `BETA` edit and a reply to an existing modern
comment thread. The proxy wrapper counts forwarded real Word read/write calls.
Empty batches, an unchanged paragraph and invalid targets must perform no insert.
The resulting DOCX bytes are independently reopened without repair and resolved
with Word Accept All/Reject All. The oracle also checks engine-resolved reference
packages separately, so expected text is not derived from Word's own result.

Commands and lane selection are documented in [scripts/README.md](../../scripts/README.md).
Raw collector artifacts live under `.cache/reliability/officejs/`; independent
oracle evidence lives under `.cache/reliability/officejs-oracle/`.

## Scope

No paid provider evaluations are required for deterministic request/error tests.
Native list/table/section editing architecture and browser facade migration remain
in their August follow-on plans. Actual Office.js coverage certifies the exercised
body/comment paths, not every possible Word version or document structure.
PDF export remains optional; accepted 1,000-paragraph performance is unchanged.

## Final verification

| Check | Result |
| --- | --- |
| `npm test` | 35 suites passed, 0 failed; four entrypoints excluded and historical version-specific skip reported |
| Actual Office.js checks | Seven passed; Word `16.0.20326.20158`, PC |
| Independent Word oracle | 12 passed, no failed checks; Word 16.0/build 16.0.20326 |
| Development build | Passed |
| Production build | Passed; three bundle warnings, taskpane approximately 762 KiB |
| Ordinary build contents | Validation entry/page absent |
| `git diff --check` | Passed |

Evidence snapshots: [Office.js results](2026-09-30-reliability-officejs.json) and
[Word oracle](2026-09-30-reliability-word-oracle.json). Collector run completed
at `2026-09-30T23:32:03.831Z`. The temporary validation add-in registration was
removed after verification. Blocked build telemetry requests were nonblocking.

SHA-256 for actual Office.js exported DOCX artifacts:

- `word-addin-plain-replacement-officejs.docx`:
  `F573B82C50CF453A7AA2135C74566C792C638E851E5180C52D614341FD147942`
- `word-addin-sibling-reply-officejs.docx`:
  `D74A914DF3A444BF62245F4F94AF2DAB9BD9836FB66595B18F6752EADFAA53E6`

The static tool-cutover guard was updated for the observed batch adapter, using
an AST to handle destructured function parameters. Behavioral outcome tests
independently check write counts and stop-after-failure behavior.

The September upgrade's final Office.js gate and reliability WP1–WP3 are closed.
Broader native tool migrations and nested-list host coverage remain in the next
agentic-tools plan.
