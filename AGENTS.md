# Agent Instructions

This repository is the standalone C# desktop project for Hainan retail electricity settlement automation.

## Safety Rules

- Never commit real ledgers, settlement outputs, customer names, screenshots with sensitive data, or finance/payment data.
- Do not overwrite user workbooks in place. Generated files must be written to an output folder or a clearly named copy.
- Preserve user edits. If local files are dirty, inspect them and work with them.
- Prefer explicit validation reports and user-visible logs over silent guesses.
- If a business rule is unclear and cannot be inferred from existing documentation, ask the user before encoding it.

## Project Scope

This repository contains only the C# rewrite. The earlier Python project remains in the original repository and remains the full historical behavior reference:

- `https://github.com/LeBronJu/hainan-settlement-tool`

The C# version is being built as a maintainable Windows desktop app. It should eventually replace or reduce reliance on the Python app, but only after staged validation.

## Repository Layout

- `HainanSettlementTool.sln`: solution file.
- `src/HainanSettlementTool.WinForms/`: Win7/8 maintenance UI. Keep it buildable and fix blocking bugs, but do not add new features or UX improvements unless the user explicitly asks.
- `src/HainanSettlementTool.Wpf/`: Win10/11 WPF main UI. New UI features and UX improvements belong here by default.
- `src/HainanSettlementTool.Core/`: business models, services, and interfaces.
- `src/HainanSettlementTool.Excel/`: ClosedXML workbook reading/writing.
- `docs/architecture.md`: layering and migration boundary.
- `CONTEXT.md`: domain vocabulary and settlement rules.
- `docs/dev-notes/`: architecture reviews, robustness priorities, and one-off technical notes.
- `HANDOFF.md`: current state for future sessions.

## Agent skills

### Issue tracker

Issues and PRDs are tracked in GitHub Issues for `LeBronJu/hainan-settlement-desktop`. See `docs/agents/issue-tracker.md`.

### Triage labels

Use the default five-role triage label vocabulary. See `docs/agents/triage-labels.md`.

### Domain docs

This is a single-context repo. See `docs/agents/domain.md`.

## Engineering Rules

- Do not make development changes directly on `main` or `master`. Create a development branch first, using the `codex/` prefix unless the user requests another branch name.
- If an issue is uncertain, ambiguous, or risky, especially when it may affect settlement correctness, workbook safety, or user-visible business rules, stop and analyze it explicitly for the user. Do not encode a guess; ask the user to decide.
- UI must not contain Excel parsing, matching, amount calculation, or workbook template rules.
- Core must not reference ClosedXML, WinForms, WPF, or file-format implementation details.
- Excel layer owns workbook reading/writing and template copying.
- Keep stage boundaries explicit.
- Win7/8 WinForms is in maintenance mode. Do not spend quality/refactor work on WinForms parity unless needed for compilation, packaging, blocking bugfixes, shared Core/Excel behavior changes, or explicit user authorization.
- Keep documentation current without creating noise. Each code, config, script, packaging, release, workflow, architecture, business-rule, UI-behavior, test-process, or task-state change must end with a documentation impact judgment.
- Final responses for development work must include documentation impact, validation performed, and work intentionally not done when applicable. Missing the documentation impact judgment means the task is not complete.
- Update only documents whose responsibility is affected. User-visible behavior usually affects `README.md` and `HANDOFF.md`; business rules affect `CONTEXT.md`; module boundaries affect `AGENTS.md` plus an ADR or dated dev-note; release and packaging changes affect `README.md`, `HANDOFF.md`, and `docs/RELEASE_CHECKLIST.md`; branch state, validation results, or next steps affect `HANDOFF.md`.
- Temporary local setup or new-machine exploration that does not enter the project mainline can state that no project docs were needed.
- For temporary investigations or one-off architecture notes, add or update a dated file under `docs/dev-notes/`.

## Current Functional Boundary

Stage 1 currently supports:

- Read an existing power workbook, or build one from `.xlsx`/`.xls`/`.csv` raw detail.
- Copy/update the ledger with current month power.
- Add newly discovered customer names and customer codes where possible.
- Emit a JSON report.
- Run "clean power data only" without updating the ledger.

Stage 2 currently supports:

- Read the manually reviewed current-month ledger.
- Copy previous-month agent/intermediary split sheets as templates and write current-month input values.
- Generate the current-month summary workbook from the previous/corrected summary template.
- Emit a JSON settlement report and a text validation report.
- Show a detailed preflight confirmation for key changes such as new split entities, new split customers, unit price changes, and tax-rate changes.

Stage 1 and Stage 2 still do not:

- Auto-fill volatile business fields such as负责人 or 项目开发人.
- Change ledger customer names to match summary/payment-account names.
- Treat irregular January/February 2026 data as generic rules.

## Business Rules To Preserve

- Ledger power unit should be `万千瓦时`; older wording saying `兆瓦时` was wrong.
- Do not rename ledger customer names only to match payment-account names in summary workbooks.
- `项目开发人` is an agent/intermediary relationship under a负责人, not the salesperson themselves.
- Historical January/February 2026 data may be irregular. Do not generalize those quirks.
- New customers can be left with blank负责人/项目开发人 for manual review in stage 1.
- If a new customer's invoice/payment note is unknown, default to `走平台扣13%`, except when that agent already has a historical rule.

## Build Command

```powershell
dotnet msbuild ".\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

## Build Portability Check

```powershell
.\scripts\check_build_portability.ps1
```

## UI Support Policy

- Win10/11 WPF is the primary user experience and default target for new UI features.
- Win7/8 WinForms remains available as a maintenance compatibility entry.
- Shared settlement correctness, workbook safety, and report generation fixes still belong in Core/Excel and benefit both entries.
- WinForms-only changes should be limited to blocking bugfixes, build/package compatibility, or explicitly requested support.

## Compatibility Target

- Target framework: `.NET Framework 4.7.2`
- Intended runtime: Windows 7 SP1 and later, with .NET Framework 4.7.2 or newer installed.
- Development requires .NET SDK 8 or newer, or equivalent MSBuild that can resolve SDK-style .NET Framework projects.
- Packaging scripts should prefer `dotnet msbuild`; if `dotnet` is unavailable, use `vswhere` to discover MSBuild.exe instead of hard-coding a Visual Studio version path.
