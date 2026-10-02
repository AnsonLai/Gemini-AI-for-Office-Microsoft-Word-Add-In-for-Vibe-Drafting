# Library offload review — August plans and September upgrade

**Reviewed:** 2026-09-30  
**Installed library:** Exact `@ansonlai/docx-redline-js@0.8.4` in root and MCP.  
**Purpose:** Internal ownership and maintenance review, separate from release notes.

## What moved out of the consumers

The reconciliation engine was already extracted before this work (`aa55e57`),
and the consumers already used its public package. This cycle removed additional
consumer orchestration and duplicated document services by adopting newer
library operations and the document-session facade. It did not extract the
entire engine for the first time.

For historical scale, the earlier `aa55e57` extraction deleted 14,975 lines
under `src/taskpane/modules/reconciliation`. That is an earlier gross deletion
count, not part of the 1,471-line net reduction measured below.

| Completed plan | Additional delegation to the library | Work retained in this repository |
| --- | --- | --- |
| [September 9 upgrade](plans/completed/2026-09-09-docx-redline-js-v0.5.4-upgrade.md) | Atomic redline execution, localized replacements, structured list/table output, document inspection and MCP open/edit/comment/serialize lifecycle. Library helpers supply comment sibling reconciliation and part specifications; the add-in still applies the resulting parts and relationships to Word's Flat OPC package. Removed iterative Word redline execution, manual table recovery, an unused structured-list module and four MCP services. | AI proposal validation and operation mapping; Word scope reads, Flat OPC preparation/transport, tracking and host outcome reporting; MCP file/session/tool contracts and blank template. |
| [Reliability and quality gates](plans/completed/2026-08-29-reliability-and-quality-gates.md) | No additional document engine extraction. Library errors and receipts propagate through consistent consumer outcomes. | Shared Gemini client, bounded provider retries, mutation outcome handling, offline gates and actual Office.js/Word verification. |
| [Agentic tools and list reliability](plans/completed/2026-08-29-agentic-tools-and-list-reliability.md) | Supported active bullet/decimal insertion with redlining enabled uses one canonical library body batch at source/resolved levels 0–1. Canonical root outdent requires a following same-list root sibling. | Validation, capability selection and supplied source-baseline guards. Native insertion covers plain anchors, deeper levels, other formats, tracking off and unsupported outdent contexts. General list editing retains a range helper; header conversion remains native. |
| [Package boundaries and integrations](plans/completed/2026-08-29-package-boundaries-and-integrations.md) | Browser editing adopts `openDocx` / inspect / atomic apply / serialize instead of rebuilding edited ZIP/XML packages. | `consumer-core.js` is a local portable extraction, not a new library module. Browser prompt projection, marker seeding, preview, upload/download and UI remain local. Preview still uses JSZip. |
| [OOXML engine and performance](plans/completed/2026-08-29-oxml-engine-and-performance.md) | No additional library implementation was moved during this plan. It profiled the settled library boundary. | Benchmarking, real Word instrumentation and loading editing modules on first use. Deferred code is still shipped. |

## How much local code became smaller

The comparison is `1b52ec0` (updated plan, before implementation) to `664f7d9`
(0.8.3 verified, original plans closed, before this documentation review).
Counts are physical lines including comments and blank lines, not logical lines
or a measure of feature coverage. Git additions/deletions include rewritten and
relocated lines; they do not measure how many lines were copied upstream.

| Measured production surface | Before → after / reduction | Interpretation |
| --- | --- | --- |
| `word-redline-runner.js` | 1,102 → 74 lines; 1,028 fewer, about 93% | The former iterative execution path is now a small planner/adapter caller. Some planning was moved to a new local module. |
| Four deleted MCP services | 876 lines removed | Package service 303, redline service 237, paragraph targeting 152, XML utilities 184. A new 138-line facade service replaces their consumer responsibilities. |
| MCP `src/` overall | 784 fewer lines | Includes the new facade service and server/session-store changes, rather than counting deleted files alone. |
| `browser-demo/demo.js` plus new `document-session.js` | 1,583 → 1,547 lines; 36 fewer | The main file shrank by 232, but the new local session/prompt helper adds 196. Responsibility consolidation is greater than the small net reduction suggests. |
| Add-in integration directory overall | 651 fewer lines | Includes the new 623-line portable core and 165-line redline planner. These remain local; the older operation runner shrank from 534 to 308 lines. |
| Above production surfaces combined | 3,069 deleted, 1,598 added; **1,471 fewer lines net** | Integration JS, MCP source MJS and the two browser editing JS files only. Excludes tests, documentation, manifests, generated files, browser HTML and other application code. |

These numbers support a substantial reduction in consumer execution and package
maintenance, especially in Word redlining and MCP. They do not establish a
percentage of the whole application's code or imply that every document tool
now uses canonical operations. New local validation, testing and integration
code elsewhere is outside this selected-surface total.

For example, `src/taskpane/modules/commands/` grew by 891 lines net in the same
range as validation, mapping and tool handling expanded. The repository added
substantial regression and host-verification evidence too. The smaller editing
surfaces therefore do not imply that the whole repository shrank by 1,471 lines.

### Reproduce the counts

Run from the repository root:

```powershell
git diff 1b52ec0..664f7d9 --numstat -- src/taskpane/modules/docx-redline-js-integration/*.js mcp/docx-server/src/*.mjs mcp/docx-server/src/services/*.mjs browser-demo/demo.js browser-demo/document-session.js
git show 1b52ec0:src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js
git show 664f7d9:src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js
```

Implementation milestones: `1a235db` (Word batch and workaround removal),
`68da8a4` (MCP facade and upgrade completion), `96b6d92` (supported canonical
list insertion), `3018409` (local portable preparation extraction), `b8f23b9`
(browser facade/parity), `42a62df` (deferred startup loading), and `148c96b`
(published 0.8.3 consumer upgrade).

## Current ownership and remaining migration

The library owns document reconciliation, supported canonical operations,
tracked-change markup, document-session packaging, revision resolution and
package inspection. The add-in still has local XML/Flat OPC bridge logic and
uses public library service subpaths as well as top-level exports; it is not
solely a `DocxDocument` wrapper. The browser and MCP use complete DOCX sessions;
Word uses scope/body Flat OPC transport. Whole-document binary Word insertion
is a deferred option.

Library 0.8.2 fixed hyperlink punctuation and manual-break rejection, and 0.8.3
fixed reported list rejection/inspection/numbering defects; 0.8.4 fixed the
soft-break, bullet-numbering, format-occurrence and table-append formatting reports. Those fixes belong
to the shared engine; the add-in carries regression evidence rather than
parallel fixes. Remaining canonical plain-to-list/header conversion and list
format changes belong to the [fresh migration follow-up](plans/2026-09-30-canonical-list-migration-follow-up.md).
The original agentic plan is closed with deferred scope, not full migration.
Once all issues are resolved, that follow-up requires retaining necessary
validation evidence, deleting `docs/library-issues/` and updating incoming links.

## Evidence and testing ownership

The [published 0.8.3 validation record](validation-reports/2026-09-30-docx-redline-v083.md)
(the [0.8.4 validation](validation-reports/2026-10-02-docx-redline-v084.md) report records the later upgrade)
records 51 passing offline suites, four excluded entrypoints, 68 actual Office.js
checks, 92 independent Word checks on exported list/native cases and 12 separate
facade Word checks. The five inapplicable engine-reference views are excluded
from pass totals. These are dated results for tested cases, not universal list
coverage or new runs performed for this review.

Consumer tests remain necessary for mapping, source freshness, no-write refusal,
package preservation, browser/MCP parity and uncertain host outcomes. Atomic
engine success cannot prove rollback after Word attempts a write. Actual
Office.js and independent Word save/reopen/Accept All/Reject All checks cover
that host boundary. Native routes require their own verification.
See the [testing overview](TESTING.md), [verification commands](../scripts/README.md),
[architecture](../ARCHITECTURE.md), [package boundaries](package-boundaries.md),
and [tool contracts](agentic-tool-contracts.md).
