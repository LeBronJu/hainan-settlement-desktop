# Real Smoke Runner

Date: 2026-06-29

Status: historical tool note. The safety model and intent remain useful, but `scripts/run_real_smoke.ps1` currently references pre-rename service/gateway class names and must be refreshed before it is treated as a runnable validation entry. Current document ownership is listed in `docs/README.md`.

## Why

The project now needs repeatable local validation against user-authorized real work copies, but the previous process required copying a long PowerShell snippet into the terminal. That made the validation interface shallow: every run exposed assembly loading, output directory creation, stage orchestration, power comparison, workbook counting, and formula-error scanning to the caller.

`scripts/run_real_smoke.ps1` makes that workflow a deeper Module. The caller provides the settlement month and explicit input paths; the script keeps the implementation details local and writes a compact `smoke-summary.json`.

## Safety

- The script does not contain real workbook paths.
- The script writes generated files only under the caller-provided `OutputRoot`, or under `dist/` when no output root is provided.
- Real workbooks still require explicit user authorization before running the script.
- Real workbooks, generated smoke outputs, and summary JSON files remain ignored by git through existing file rules.

## Scope

The runner covers the current practical smoke path:

- optional raw detail cleaning from `.xlsx`, `.xls`, or `.csv`
- optional comparison with an existing cleaned power workbook
- optional stage 1 ledger update into an isolated output folder
- optional stage 2 generation into an isolated output folder
- generated `.xlsx` scan for common formula error text
- machine-readable JSON summary

The runner is not an automated regression fixture. It is a local validation tool for authorized real data or future sanitized work copies.
