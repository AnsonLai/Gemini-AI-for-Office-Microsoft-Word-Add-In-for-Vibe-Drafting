# Reliability and Quality Gates Plan

**Date:** 2026-08-29  
**Status:** Active — consolidated plan; no work has been executed from this document.

## Objective

Turn the reliability lessons from the first hardening pass into repeatable
tests, observability, bridge verification, and CI gates.

## Carry-forward verification

The first reliability plan records WP1–WP8 implementation as complete, but
still calls out manual Word smoke tests, a real-model eval run, and remaining
overall-acceptance verification. Close those items as part of this plan rather
than treating implementation notes as proof of release readiness.

## Work packages

1. **Test runner and CI.** Add one offline runner for the current test layout,
   fix the known migrated-tool test mismatch, and run the build/test gate in CI.
2. **Unified Gemini client.** Centralize timeout, cancellation, retry, quota,
   authentication, and error classification behavior behind one injectable
   client.
3. **Reliability event log.** Record machine-readable defense/failure events and
   expose a concise user-copyable diagnostic summary.
4. **Testable taskpane seams.** Extract pure routing, history, and dispatch logic
   from the Office.js-coupled chat loop.
5. **Record/replay evals.** Add offline replay and synthetic regression cases
   for thought leakage and repeated/truncated model output; keep live runs
   explicitly opt-in.
6. **Word bridge verification kit.** Share mock helpers, cover fallback and
   cleanup paths, and record the Desktop/Web manual smoke matrix.
7. **Package-boundary contracts.** Assert every consumer import and one real
   engine smoke path against the pinned `@ansonlai/docx-redline-js` surface.
8. **Drift control.** Single-source version markers, add docs freshness checks,
   and establish the lint ratchet without masking existing debt.

## Release gates

- Offline tests, replay evals, package contracts, docs checks, build, and lint
  pass in the supported environment.
- Word Desktop/Web smoke results are recorded for tracked edits, fallback
  insertion, undo/revert, invalid targets, and network failure recovery.
- Any live-model limitation or unverified behavior is recorded explicitly.

## Migrated source plans

- `docs/plans/migrated/reliability-hardening-plan.md`
- `docs/PLAN.md` (retained at its historical path; migrated by banner)

