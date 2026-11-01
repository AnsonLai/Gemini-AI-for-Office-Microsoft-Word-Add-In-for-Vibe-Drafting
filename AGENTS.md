# AGENTS.md: launch card

Use this file to route work on the Gemini for Word add-in. Don't explore the
whole repository before acting. The add-in is an in-place Word taskpane with a
small inspect/apply agent loop. `@ansonlai/docx-redline-js` (exact pin in
`package.json`) does every document transformation. This repo owns the UI, the
agent loop and the Word host adapter.

## Pick the route

| Task | Start here |
|---|---|
| The model misread or under-used a tool | `src/taskpane/modules/chat/document-editing-guide.js` (versioned guide), then `tool-registry.js` (tool descriptions) |
| Add, lift or change an edit capability | `document-edit-intents.js` (validation, combination rules, library operations), then the [tool contract](docs/agentic-tool-contracts.md) |
| Handles, inspection, freshness | `document-inspection-session.js`, `docx-redline-js-integration/word-source-identity.js` |
| Checkpoint, write, after-write check | `document-tools.js`, `docx-redline-js-integration/word-operation-runner.js`, `modules/storage/checkpoint-store.js` |
| Turn limits, stop reasons, closing summary | `document-agent-loop.js`, `mutation-loop-guard.js` |
| Gemini requests, retries, errors | `gemini-client.js`, `modules/config/model-profiles.js` |
| Taskpane UI and settings | `src/taskpane/taskpane.js`, `chat-ui.js`, `taskpane.html` |
| Known bugs, deferred work, Word API uses | [KNOWN_BUGS](docs/plans/KNOWN_BUGS.md) |
| Which test lane proves what | [Testing](docs/TESTING.md) |

## Rules

- **Library first.** A new document capability is a library release plus a thin
  intent mapping, never add-in document logic. Word-specific glue (Flat OPC
  handling, Word export noise) stays here; the library stays host-independent.
- **Library bugs are reported, not patched.** Add a KNOWN_BUGS entry with a
  reproducer that uses only public APIs, plus a test that prints a
  `KNOWN_LIBRARY_DEFECT:` line. No add-in workarounds. A clear validation
  message in front of an unclear library refusal is fine.
- **Probe before you refuse.** Before adding or keeping an add-in limit, run
  the same operations through the library and check Accept All and Reject All.
  A limit the library doesn't need pushes the model into degraded output: lines
  joined with slashes, paragraphs run together, a request refused outright.
- **Every capability change touches five places:** the compiler, the guide
  (bump its version), the tool description, the contract doc, and a test.
  Add a golden replay step when real Word behaviour matters.
- **Refusals tell the model what to do next.** Name the alternative (another
  target, one combined action). Use `requiresReinspection` when a fresh
  inspection can fix it.
- **Truthful results.** Only verified writes count. Never replay an uncertain
  write, and never claim work from a proposal alone.

## Verification

Run the closest suite first (`node tests/<suite>.mjs`), then `npm test` (all
offline suites must pass). Real-Word lanes, slowest last:

```bash
node scripts/run-golden-scenario.mjs --scenario privacy --launch   # replay, privacy notice
node scripts/run-golden-scenario.mjs --launch                      # replay, original 16-prompt NDA
npx webpack --mode development --env WORD_HOST_VALIDATION=1 --output-path .cache/reliability/direct-dist
node scripts/run-officejs-validation.mjs --launch --dist-dir .cache/reliability/direct-dist --artifacts-dir .cache/reliability/direct-officejs
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-wp6-word.ps1 -FixtureManifest .cache/reliability/direct-officejs/officejs-fixtures.json -ArtifactsDir <dir> -SkipNativeInsert -TimeoutSeconds 600
```

- Live model runs add `--mode live`. They need `GEMINI_API_KEY`, which the
  owner sets in their own terminal. Never ask for the key or write it anywhere.
- Each run's `.cache/golden/<UTC time>/` folder holds `model-log.jsonl` (every
  request and response) and a `.docx` after each step; read those before
  guessing.
- The NDA scenario replays the owner's original prompts word for word; add new
  cases to the privacy scenario instead.
- **If a lane hangs:** check the dev certificate (`~/.office-addin-dev-certs`,
  30-day expiry), and look for leftover `WINWORD` processes from earlier lanes.
  Never close the owner's own Word windows.

## Releases

The release marker is in `src/taskpane/taskpane.html` (meta `release-marker`)
and `src/taskpane/taskpane.js` (`RELEASE_MARKER`). Notes go in
`docs/release/`. `npm run build` produces `dist/` for the Azure `$web`
container. The owner pushes, tags, publishes and deploys; prepare everything,
then ask.
