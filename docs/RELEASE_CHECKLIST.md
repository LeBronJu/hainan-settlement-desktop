# Release Checklist

Use this checklist before merging release-bound changes or publishing packages.

## Documentation Gate

- Confirm every changed behavior, business rule, workflow, package, script, test workflow, or task state has a documentation impact judgment.
- Update only the documents whose responsibility changed.
- For release or packaging changes, check `README.md` and this file.
- For business-rule changes, check `CONTEXT.md`.
- For module-boundary or architecture changes, check `AGENTS.md` and add an ADR or dated dev note when the decision should guide future agents.
- For branch state, release status, validation results, or next steps, check `HANDOFF.md`.
- If no project docs are needed, record the reason in the final response or PR notes.

## Validation Gate

- Run relevant tests, builds, and packaging checks, or explain why they are not needed.
- Confirm no real ledgers, customer data, settlement outputs, screenshots, or sensitive finance data are included.
- Confirm generated packages contain all required `.dll`, `.config`, and executable files.
