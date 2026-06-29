# Documentation Sync Gate

Date: 2026-06-25

## Why

The project needs documentation that stays current, but a blanket "update all docs every time" rule creates noise. Agents may rewrite unaffected docs just to satisfy process, making future reviews harder and diluting the useful signal.

The rule is therefore a gate on judgment, not a mandate to edit every document.

## Rule

Every project work session must start by reading current instructions, not by relying on memory or prior chat summaries:

- Run `git status --short --branch`.
- Read `AGENTS.md` and `HANDOFF.md`.
- Read the owning document for the area being changed.
- Repeat this reading gate after context compaction, a long pause, or a change in task direction.

Every code, config, script, packaging, business-rule, UI-behavior, test-workflow, or task-state change must end with a documentation impact judgment:

- If docs were affected, list the documents updated.
- If docs were not affected, explain why.
- If the judgment is missing, the task is not complete.

Development final responses should also include validation performed and any intentionally unfinished work.

## Document Responsibilities

- `README.md`: user-facing scope, setup, build, packaging, release, and notable behavior.
- `HANDOFF.md`: current branch state, release status, validation results, and next steps.
- `CONTEXT.md`: stable business vocabulary, settlement rules, and data safety constraints.
- `AGENTS.md`: agent operating rules, engineering constraints, and workflow gates.
- `docs/architecture.md`: module layering, stage boundaries, and migration strategy.
- `docs/RELEASE_CHECKLIST.md`: merge, release, packaging, and documentation gates.
- `docs/dev-notes/`: dated investigations, architecture notes, and lightweight governance decisions.

If a future `TASK_BOARD.md` or `AI_HANDOFF.md` exists, use it for active task tracking or agent-specific handoff notes rather than expanding user-facing docs.

## Common Cases

- UI copy or visible behavior: usually check `README.md`, `HANDOFF.md`, and task tracking docs if present.
- Business settlement rule: check `CONTEXT.md` and update it when the rule changes.
- Module seam, layering, or adapter responsibility: check `AGENTS.md` and record the decision in an ADR or dated dev note when useful.
- Release or packaging flow: check `README.md` and `docs/RELEASE_CHECKLIST.md`.
- Temporary local setup or new-machine exploration that does not enter the project mainline: no project document update is required; state that explicitly.
