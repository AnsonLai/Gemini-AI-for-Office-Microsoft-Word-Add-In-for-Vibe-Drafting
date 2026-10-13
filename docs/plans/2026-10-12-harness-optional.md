# Harness (optional): a Pi-style agent harness

**Date:** 2026-10-12 (America/Vancouver)
**Status:** Optional and parked. Not scheduled; start only when a trigger below becomes real.
**Related:** [Library-first simplification, phase 2](2026-10-12-library-first-simplification-phase-2.md), which this was split out of. [Phase-1 plan owner decisions](completed/2026-10-02-library-first-lightweight-agent-reliability.md#approved-owner-decisions-2026-10-02).

> **Recommendation (Claude, 2026-10-12): adopting a Pi harness is a bad idea for now.**
>
> - **The loop is not the problem.** The plugin's agent loop is about 1,000 lines. It passes the 16-prompt Word replay with zero auxiliary model calls, and its safety rules are all enforced and tested. The "unwieldy" problem came from document logic in the plugin, and the library-first plan fixes that. The loop is already the small part.
> - **The saving is smaller than it looks.** The safety policy, Gemini-specific handling and a harness-to-UI adapter stay ours. The likely net cut is 400–600 lines, not 1,000, and the bundle grows: `pi-ai` brings `@google/genai` in place of an 86-line fetch client.
> - **We'd be fighting the framework's defaults at the most safety-critical point.** Pi assumes parallel tools, automatic continuation and a `terminate` hint that only stops when every result agrees. Our invariants are the opposite: sequential tools, at most one write per response, and a hard stop after an uncertain write. They would live in hooks around someone else's loop, and that loop decides whether another write happens.
> - **It's the opposite ownership bet from the library.** Relying on `docx-redline-js` works because the owner controls it. Pi is a third-party, fast-moving 1.x project. Browser support for `pi-agent-core` is unconfirmed, and debugging a bad edit would run through code we don't own.
>
> **What would change this:** a real need for other model providers (Claude, OpenAI), steering mid-run, queued follow-ups, persistent sessions with compaction, or on-demand skills. Pi provides those, and building them ourselves would eventually recreate it. If that need appears, start with WP-H0 below.
>
> **Cheap in the meantime:** when the loop is touched for other reasons, borrow Pi's *shape* without its runtime: named lifecycle events, schema-defined tools, and the safety policy in one module. That keeps a later adoption cheap. If multiple providers is the only want, consider `pi-ai` alone; it does not touch the safety loop.

## Owner direction and history

- **Phase 1 (2026-10-02):** "tighten a small local loop; no Pi runtime adoption or experiment". Pi informed the design only. A general skill system was excluded.
- **2026-10-12:** the owner asked for the agent loop to be "more like a harness, something like [pi.dev](https://pi.dev/docs/latest)". After the recommendation above, the owner chose to keep the idea as this separate optional plan rather than schedule it.

## Triggers to start

Start WP-H0 only when at least one of these is a committed product requirement:

1. **Providers.** Users need non-Gemini providers, or automatic model fallback beyond today's transient retries (issue #3's remaining request).
2. **Steering.** Users need to redirect the agent mid-run, or queue follow-up requests while one runs.
3. **Sessions.** Conversations must persist across taskpane reloads, with compaction.
4. **Guidance.** Editing guidance must load on demand, Pi-skill style, because one compact guide no longer fits.

## Plugin code a harness would replace

| Module | Lines | Role today | Under a Pi-style harness |
| --- | ---: | --- | --- |
| `chat/document-agent-loop.js` | 343 | Turn loop, budgets, stop reasons, stale replan | `Agent` loop plus a small policy module of hooks |
| `chat/chat-history.js` | 359 | Bounded history, pairing, sanitising | Harness messages plus `transformContext` / `convertToLlm` |
| `chat/gemini-client.js` | 86 | Gemini fetch, timeout, retry | Harness provider layer (`pi-ai` Google provider), keeping our retry policy |
| `chat/tool-registry.js` | 75 | Declarations, validation, mutability | TypeBox `AgentTool` definitions |
| `chat/function-call-recovery.js` | 45 | Malformed-call recovery | Harness tool-argument validation |
| `chat/mutation-loop-guard.js` | 60 | Failed-mutation budget | `beforeToolCall` / `afterToolCall` policy |
| `chat/refreshed-document-context.js` | 31 | Context refresh after a write | `transformContext` |

That totals about 1,000 lines (2026-10-12). The safety policy, Gemini specifics and a UI adapter survive any migration.

## What Pi offers

Checked 2026-10-12 against the published READMEs and npm metadata, not yet tested in Word.

| Layer | Package | Runs in the Word add-in? | Relevance |
| --- | --- | --- | --- |
| Coding-agent SDK (`createAgentSession`, JSONL sessions, extensions, skills, compaction, project trust) | `@earendil-works/pi-coding-agent` | **No.** Its SDK docs say it embeds Pi in a Node.js or Bun process | Concepts only: session tree, compaction, skills, extensions |
| Agent core | `@earendil-works/pi-agent-core` 1.2.0, MIT; dependencies `typebox` and `pi-ai` only | **Unverified.** No browser statement; the README examples use Node APIs, but there is no Node-only dependency | `Agent`, `agentLoop`, an event stream, TypeBox `AgentTool`s, `transformContext`, `convertToLlm`, steering and follow-up queues, `toolExecution: 'sequential'`, `beforeToolCall` (can block), `afterToolCall`, `abort()`, a `terminate` hint |
| Providers and models | `@earendil-works/pi-ai` 1.2.0, MIT | **Yes, per its README** ("supports browser environments"). API keys are passed explicitly or through a `CredentialStore`. Providers are lazy chunks; Google uses `@google/genai`, which has a web build | One model catalog across many providers, plus token and cost tracking |

## Safety invariants any harness must keep

These are enforced in code, never by prompt:

| Invariant | Pi mechanism |
| --- | --- |
| Document tools run sequentially | `toolExecution: 'sequential'`; the mutation tool also sets `executionMode: 'sequential'` |
| At most one mutation call per model response; reject every call in a response that contains several | `beforeToolCall` blocks every call in such a batch |
| Stop immediately after a partial, unconfirmed or uncertain write; no replay | `afterToolCall` records the outcome, then `agent.abort()`. The `terminate` hint alone is insufficient, because it stops only when *every* result in the batch terminates |
| Two failed-mutation cycles; read-only calls don't reset the count; one fresh-context replan after a proven no-write stale refusal | Policy state consulted by `beforeToolCall` and `afterToolCall` |
| Refresh document context after a confirmed write; handles expire | `transformContext` injects the fresh inspection; the handle store stays authoritative |
| Checkpoint saved before a write, or no write | Inside the `apply_document_edits` tool, unchanged |
| Cancellation never implies rollback; the UI always unlocks with a terminal state | Harness events (`agent_end`, abort) drive the existing UI states |

## Work packages (only once triggered)

### WP-H0: Spike

Timeboxed, on a throwaway branch. Bundle `pi-agent-core` and `pi-ai` (Google provider only) with the add-in's webpack build, and run one inspect-then-apply turn in desktop Word against the NDA fixture. Answer:

- Does it build? At what bundle size? Do the lazy provider chunks stay out of the main bundle?
- Does it run in Word's WebView2 under the current CSP? `connect-src 'self' https://generativelanguage.googleapis.com` must stay unchanged for Gemini-only; lazy chunks load under `script-src 'self'`.
- Does Gemini tool calling through `pi-ai` keep our schemas? Check `StringEnum` versus `Type.Enum`, and Gemini 3 strict mode.
- Can the API key come from the existing settings store?
- Can every safety invariant above be expressed with `beforeToolCall`, `afterToolCall` and `abort()` without patching the harness?
- Is there any Node-only import or `process`/`fs` reference at runtime?

**Exit:** an owner go/no-go decision. If the answer is no-go, record the blocker. The fallback is to reshape our own loop around the same concepts; never fork or patch Pi. Report harness defects upstream.

### WP-H1: Migrate (if go)

- Replace `document-agent-loop.js`, `gemini-client.js`, `tool-registry.js`, `function-call-recovery.js`, `mutation-loop-guard.js` and most of `chat-history.js` with:
  - two TypeBox `AgentTool`s;
  - one policy module implementing the invariants table;
  - `transformContext` / `convertToLlm`;
  - a harness-event-to-UI adapter.
- Keep our retry and backoff semantics. This is a provider or `streamFn` concern; check `pi-ai` error events.
- Pin exact versions.

**Exit:**
- Deterministic offline contracts pass unchanged: multiple-mutation refusal, uncertain-write stop, no replay, failure budget, stale replan, cancellation, paired history, checkpoint refusal.
- The NDA replay passes. The direct replay lane may need a scripted `pi-ai` custom provider.
- The Office.js and oracle gates pass.
- Net plugin lines fall.

### WP-H2: Harness features

Each feature is its own owner decision:

- steering and follow-up in the taskpane;
- IndexedDB conversation sessions with compaction (Pi's JSONL sessions are Node-only);
- on-demand guidance (reopens phase 1's skill-system exclusion);
- a token and cost display;
- multi-provider selection, which widens the CSP `connect-src`, and automatic model fallback.

## Open owner decisions (only once triggered)

1. **Dependency.** Adopt `pi-agent-core` + `pi-ai` after a passing spike, or build our own Pi-shaped loop?
2. **Skills and sessions.** Reopen phase 1's skill-system exclusion? Add persistent sessions?
3. **Providers.** Stay Gemini-only, or allow others? Automatic fallback on transient failures?
