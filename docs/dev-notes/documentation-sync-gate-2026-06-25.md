# Documentation Sync Gate

Date: 2026-06-25

Status: current process note. The live document map is `docs/README.md`; use this file for the documentation impact gate.

## Why

The project needs documentation that stays current, but a blanket "update all docs every time" rule creates noise. Agents may rewrite unaffected docs just to satisfy process, making future reviews harder and diluting the useful signal.

The rule is therefore a gate on judgment, not a mandate to edit every document.

## Rule

Every project work session must start by reading current instructions, not by relying on memory or prior chat summaries:

- Run `git status --short --branch`.
- Read `AGENTS.md`, `HANDOFF.md`, and `docs/README.md`.
- Use `docs/README.md` to identify and read the owning document for the area being changed.
- Repeat this reading gate after context compaction, a long pause, or a change in task direction.

Every code, config, script, packaging, business-rule, UI-behavior, test-workflow, or task-state change must end with a documentation impact judgment:

- If docs were affected, list the documents updated.
- If docs were not affected, explain why.
- If the judgment is missing, the task is not complete.

Development final responses should also include validation performed and any intentionally unfinished work.

## Document Responsibilities

- `README.md`: user-facing scope, setup, build, packaging, release, and notable behavior.
- `docs/README.md`: document map, source-of-truth index, and current/historical doc status.
- `HANDOFF.md`: concise current branch state, release status, latest validation summary, package pointers, safety boundary, and next steps.
- `CONTEXT.md`: stable business vocabulary, settlement rules, and data safety constraints.
- `AGENTS.md`: agent operating rules, engineering constraints, and workflow gates.
- `docs/architecture.md`: module layering, stage boundaries, and migration strategy.
- `docs/*-current-behavior.md`: detailed current behavior for mature modules that would make `CONTEXT.md` too noisy.
- `docs/CHANGELOG.md`: completed high-signal milestones that should remain discoverable without expanding `HANDOFF.md`.
- `docs/RELEASE_CHECKLIST.md`: merge, release, packaging, and documentation gates.
- `docs/dev-notes/`: dated investigations, architecture notes, and lightweight governance decisions.

If a future `TASK_BOARD.md` or `AI_HANDOFF.md` exists, use it for active task tracking or agent-specific handoff notes rather than expanding user-facing docs.

Do not append long historical build/test logs or investigation narratives to `HANDOFF.md`. Keep the current result there, move durable analysis to `docs/dev-notes/`, mature behavior to `docs/*-current-behavior.md`, and completed milestones to `docs/CHANGELOG.md`.

## Common Cases

- UI copy or visible behavior: usually check `README.md`, `HANDOFF.md`, and task tracking docs if present.
- Business settlement rule: check `CONTEXT.md` and update it when the rule changes.
- Module seam, layering, or adapter responsibility: check `AGENTS.md` and record the decision in an ADR or dated dev note when useful.
- Release or packaging flow: check `README.md` and `docs/RELEASE_CHECKLIST.md`.
- Script or repeatable validation workflow: check `README.md`, `docs/RELEASE_CHECKLIST.md` if release-related, and a dev note when the script has safety constraints.
- Test workflow or validation policy: check `docs/architecture.md` and update the owning dev note or checklist when repeatable behavior changes.
- Task-state, latest package, current branch, or immediate next-step change: update `HANDOFF.md`.
- Completed milestone that matters historically but is no longer the current handoff: update `docs/CHANGELOG.md`.
- Documentation-only cleanup: update `docs/README.md` when ownership, routing, status taxonomy, or canonical links change.
- Real-data analysis or smoke: update the relevant dev note or `HANDOFF.md` with authorization scope, safety boundary, output location, and whether the authorization is historical-only or still active.
- Temporary local setup or new-machine exploration that does not enter the project mainline: no project document update is required; state that explicitly.
