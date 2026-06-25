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
- `src/HainanSettlementTool.WinForms/`: Win7/8 desktop UI only.
- `src/HainanSettlementTool.Wpf/`: Win10/11 WPF desktop UI shell only.
- `src/HainanSettlementTool.Core/`: business models, services, and interfaces.
- `src/HainanSettlementTool.Excel/`: ClosedXML workbook reading/writing.
- `docs/architecture.md`: layering and migration boundary.
- `HANDOFF.md`: current state for future sessions.

## Engineering Rules

- UI must not contain Excel parsing, matching, amount calculation, or workbook template rules.
- Core must not reference ClosedXML, WinForms, or file-format implementation details.
- Excel layer owns workbook reading/writing and template copying.
- Keep stage boundaries explicit.
- Do not migrate stage 2 until stage 1 is stable against real working copies or sanitized samples.

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

## Compatibility Target

- Target framework: `.NET Framework 4.7.2`
- Intended runtime: Windows 7 SP1 and later, with .NET Framework 4.7.2 or newer installed.
- Development requires .NET SDK 8 or newer, or equivalent MSBuild that can resolve SDK-style .NET Framework projects.
- Packaging scripts should prefer `dotnet msbuild`; if `dotnet` is unavailable, use `vswhere` to discover MSBuild.exe instead of hard-coding a Visual Studio version path.
