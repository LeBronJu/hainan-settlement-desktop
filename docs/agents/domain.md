# Domain Docs

How engineering skills should consume this repo's domain documentation when exploring the codebase.

## Layout

This is a single-context repo.

Expected domain docs:

- `docs/README.md` for the current documentation map and source-of-truth index
- `docs/CHANGELOG.md` for completed high-signal milestone history
- `CONTEXT.md` at the repo root, if present
- `docs/adr/` for architectural decisions, if present
- `docs/architecture.md` for current layering and migration boundaries
- `docs/*-current-behavior.md` for detailed current behavior of mature modules
- `AGENTS.md` and `HANDOFF.md` for current working rules and handoff state

If `CONTEXT.md` or `docs/adr/` do not exist, proceed silently. They can be created later when a domain term or architecture decision needs to be recorded.

## Domain Areas

- Stage 1: raw retail detail cleaning and ledger update.
- Stage 2: manually reviewed ledger to split workbooks, summary workbook, preflight report, and validation report.
- Excel layer: workbook reading, raw detail cleaning, template copying, and generated workbook writing.
- Win7/8 UI: WinForms maintenance desktop entry.
- Win10/11 UI: WPF primary desktop entry for new UI work.

## ADR Conflicts

If future work contradicts an ADR, surface it explicitly rather than silently overriding it.
