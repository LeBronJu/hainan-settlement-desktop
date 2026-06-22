# Handoff

Last updated: 2026-06-18

## Project

Standalone C# desktop rewrite of the Hainan retail electricity settlement automation tool.

Local path:

```text
D:\Document\文件处理\hainan-settlement-desktop
```

GitHub:

```text
https://github.com/LeBronJu/hainan-settlement-desktop
```

Original mixed Python/C# repository:

```text
https://github.com/LeBronJu/hainan-settlement-tool
```

This repository was split out from the original `csharp/` subdirectory so future C# work can proceed as a dedicated project.

## Current State

- C# solution exists and builds.
- UI has been modernized from the first WinForms draft.
- Month selector is a fixed dropdown from `2026年2月` through `2026年12月`.
- Input layout alignment bug has been fixed by explicit `TableLayoutPanel` row creation.
- The app implements stage 1 and an initial C# stage 2 migration.

## Stage 1 Current Capability

Inputs:

- Base ledger workbook.
- Existing power workbook, or `.xlsx`/`.csv` raw retail detail.
- Optional reference ledger.
- Output folder.

Outputs:

- Updated ledger copy.
- JSON report.

Current limitations:

- No direct `.xls` raw-detail cleaning.
- No formula recalculation through Excel automation.

## Stage 2 Current Capability

Inputs:

- Manually reviewed current-month ledger.
- Previous-month agent split workbook folder.
- Previous-month intermediary split workbook folder.
- Previous or corrected summary workbook.
- Output folder.

Outputs:

- Agent split workbooks under `2026年代理 - 海南`.
- Intermediary split workbooks under `2026年居间 - 海南`.
- Monthly agent-fee summary workbook.
- JSON settlement report.

Validation notes:

- C# stage 2 was smoke-tested against the local stable March 2026 settlement reference folder.
- C# and the Python baseline matched on ledger-derived counts and totals:
  - proxy rows: 59
  - intermediary rows: 7
  - proxy groups: 16
  - intermediary groups: 3
  - proxy total: 4.6098
  - intermediary total: 2.6526

## Environment

Development machine has:

- .NET SDK 8/9.
- Visual Studio Build Tools 2022.
- .NET Framework 4.7.2 targeting pack / SDK.

Verified build command:

```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" ".\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

Expected result: `0 个警告 / 0 个错误`.

## Suggested Next Steps

1. Verify the standalone repository builds after clone/restore.
2. Add sanitized sample workbooks for repeatable stage 1 tests.
3. Add Core tests around matching and report semantics.
4. Run broader stage 2 acceptance on working copies for later months.
5. Add optional Excel automation recalculation if required.
6. Continue hardening summary/split formatting edge cases as real templates surface.

## Useful Skills For Future Sessions

- `spreadsheets:Spreadsheets` for workbook behavior.
- `diagnose` for debugging UI/build/workbook failures.
- `handoff` when compacting context again.
