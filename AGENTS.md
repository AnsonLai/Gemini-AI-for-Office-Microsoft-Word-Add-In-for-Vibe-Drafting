# Notes for coding agents

This repository is a Word task pane add-in (Office.js) that lets a Gemini model edit the open document. The model gets three tools (`inspect_document`, `apply_document_edits`, `search_web`); every document change is prepared by the external library `@ansonlai/docx-redline-js` (exact pin in `package.json`) and written to Word in one step. Start with [STATE.md](STATE.md), [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) and [docs/plans/KNOWN_BUGS.md](docs/plans/KNOWN_BUGS.md).

## Rules that are easy to get wrong

- **Library bugs are reported, not worked around.** The owner also owns the library (source at `..\..\Docx Redline JS`). Reproduce the defect with the library alone, record it in `docs/plans/KNOWN_BUGS.md`, and leave the add-in as it is.
- **Word-specific code stays here.** Flat OPC handling and Word export noise live in `src/taskpane/modules/docx-redline-js-integration/`; only host-independent DOCX behaviour belongs in the library.
- **Do not push, tag or deploy** unless the owner asks in that conversation. Releases follow `.anson/publish instructions.md`.
- **Stage files by name.** Never `git add -A`; the owner often has uncommitted work in the tree.
- **A changed editing guide needs live runs.** `src/taskpane/modules/chat/document-editing-guide.js` is the model's prompt. Bump its version line and run both scenarios live on both default models before it ships.

## Testing in real Word

Offline tests (`npm test`) mock Word. They do not show that an edit survives Word's own save, tracked-change handling or Accept All / Reject All, so anything that touches editing, Revert, the agent loop or the task pane UI is also run in desktop Word. All of the lanes below are scripted: none needs clicking in Word, and none needs computer-use tools.

### What you need

- Windows with desktop Word installed and signed in.
- A trusted development certificate (`~/.office-addin-dev-certs/localhost.crt`, valid 30 days). Renewing it (`npx office-addin-dev-certs install`) changes the machine's trusted root store, so **ask the owner to run it**; do not run it yourself.
- For live runs only: `GEMINI_API_KEY` in the user environment. Read it into the process without printing it, and never ask the owner to paste the key into chat:

```powershell
$env:GEMINI_API_KEY = [Environment]::GetEnvironmentVariable('GEMINI_API_KEY','User')
```

### How every lane works

Each runner builds its own development bundle, serves it on a loopback port, registers a temporary add-in, starts **its own** Word instance (`winword /x`) on a per-run copy of the fixture, drives the task pane from a script inside the page, exports the document after each step, and then closes only the Word it started and removes its registration. The owner's own Word windows are never touched: do not close, save or type into them.

Ports: 3000 is the owner's dev server, 3001 is shared by the golden scenario and the taskpane checks (run one at a time), 3002 is the Office.js validation lane.

### The lanes

Run them one after another, not in parallel. Details and flags are in [scripts/README.md](scripts/README.md).

| Lane | Command | What it shows | Time |
| --- | --- | --- | --- |
| NDA scenario, scripted model | `node scripts/run-golden-scenario.mjs --launch --artifacts-dir .cache/golden/nda` | The owner's 16 real requests through the real task pane and loop, each scored on the exported document | about 2 min |
| Privacy notice scenario, scripted model | add `--scenario privacy` | A second document (lists, table, footnote), 16 requests | about 2 min |
| Taskpane checks | `node scripts/run-taskpane-checks.mjs --launch --artifacts-dir .cache/golden/checks` | Cancel, Retry, Revert, Refresh and Settings behaviour with a scripted model | about 4 min |
| Office.js validation | build with `npx webpack --mode development --env WORD_HOST_VALIDATION=1 --output-path .cache/reliability/direct-dist`, then `node scripts/run-officejs-validation.mjs --launch --dist-dir .cache/reliability/direct-dist --artifacts-dir .cache/reliability/direct-officejs` | 20 edit cases through the production inspect and apply tools, with checkpoints | about 3 min |
| Word oracle on those exports | `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 -FixtureManifest .cache/reliability/direct-officejs/officejs-fixtures.json -ArtifactsDir .cache/reliability/direct-oracle -SkipNativeInsert -TimeoutSeconds 600` | Word itself opens each export and runs Accept All and Reject All; text, lists, comments and tables must match | about 3 min |
| Library fixtures in Word | `npm run test:word` | The same oracle on the public library fixtures | about 3 min |
| Live model | add `--mode live --profile fast --model <id>` or `--profile slow --slow-model <id>` to the scenario command | Real Gemini output on the same scenarios. Costs provider usage | 5 to 25 min per run |

`run-golden-scenario.mjs` has no `--help`; an unknown flag starts a run. `--until N` stops after N requests, `--skip-build` reuses the last build, `--cleanup` removes a registration left by an interrupted run.

### Reading the results

- Each run writes `report.md` and `report.json` in its artifacts directory, plus a `.docx` and `.json` per step and, for scenarios, `model-log.jsonl` with every model request and tool result. Read the report first, then the step JSON, then the log. The log is large: filter it with a small script instead of reading it whole.
- A scenario step is scored on the exported document (accepted view), on the tool receipts (one verified write per batch, no nested model calls) and on the tracked-change author. A refusal step passes only if nothing was written.
- Run the lanes in the background and wait for them to finish; a full pass of every lane takes 15 to 20 minutes, a live pass on two models about 20 more.

### When a lane misbehaves

- **Hangs after "Started a separate Word instance", no progress:** check the certificate's expiry first. Then look for a stray `WINWORD.EXE /Automation -Embedding` left by an earlier COM run and end only that process.
- **Every edit refused with "The Word source changed after inspection":** Word regenerated some ids on export and the freshness token saw it as a change. The scenario driver prints a `[StaleSource]` diff per step; the fix belongs in `word-source-identity.js`.
- **Revert reports "could not be verified":** same family. Word gives paragraphs and table rows new `w14:paraId` values when a body is re-inserted.
- **Reject All leaves an extra paragraph in a COM check:** use `Document.Revisions.RejectAll()` (what the ribbon does), not `Document.RejectAllRevisions()`.
- **A live step fails:** decide which of three things it is before changing anything. A model judgement gap goes under KB-17 and may need a guide sentence. A stale scenario expectation (the model did something correct that the scenario did not allow) is fixed in `scripts/golden/`. A library refusal is reproduced with the library alone and reported upstream.

### What to record

State results with their counts and what was not run. `STATE.md` holds the evidence for the current release; `docs/plans/KNOWN_BUGS.md` holds open defects and the live-run history. A lane that ran before a later source change is reported as such, not as a result on the final tree.
