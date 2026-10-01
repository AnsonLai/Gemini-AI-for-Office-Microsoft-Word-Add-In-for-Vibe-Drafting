# Historical Plan: Documentation Consolidation

> Role: **Atomic task execution plan** using XML-structured task blocks and explicit verification steps.

## Planning Standard

- Plans should be written as atomic tasks.
- Task definitions should use XML-like structure to make state transitions unambiguous.
- Every task block must include verification steps.

### Recommended Task Block Format

```xml
<task id="T-001" status="todo|in_progress|done">
  <goal>Clear, testable objective.</goal>
  <changes>
    <file path="relative/path.ext">What will change.</file>
  </changes>
  <verification>
    <step>Command, assertion, or manual check.</step>
  </verification>
</task>
```

This plan records the one-time transition to the GSD documentation structure, completed on 2026-02-13. It is not the current product roadmap. See [ROADMAP.md](ROADMAP.md), [STATE.md](STATE.md), and the [dated execution plans](docs/plans/) for current project status.

For the current split between library-owned DOCX behavior and consumer-owned
mapping, host bridges, reliability and UI/session code, see the
[library offload review](docs/library-offload-review.md).

## Completed Planning Record: [2026-02-13]

### [DOCUMENTATION] [Cleaning up legacy documentation]

- [x] Merge `STANDALONE_USAGE.md` into `SPEC.md`.
- [x] Merge `RECONCILIATION_WALKTHROUGH.md` and `OOXML_ARCHITECTURE.md` into `ARCHITECTURE.md`.
- [x] Generate `ROADMAP.md` from `WORD_JS_API_DEPENDENCIES.md`.
- [x] Prepare `STATE.md`, `PLAN.md`, and `SUMMARY.md`.
- [x] Delete `OOXML_ARCHITECTURE.md`.
- [x] Delete `RECONCILIATION_WALKTHROUGH.md`.
- [x] Delete `STANDALONE_USAGE.md`.
- [x] Delete `WORD_JS_API_DEPENDENCIES.md`.

**Verification**:
- Verify all links in `SPEC.md` work.
- Verify `pkg:package` usage examples are preserved.
- Verify migration phases are documented in `ROADMAP.md`.
