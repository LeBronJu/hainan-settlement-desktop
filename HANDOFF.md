# Handoff

Last updated: 2026-06-23

## Project

Standalone C# desktop rewrite of the Hainan retail electricity settlement automation tool.

- Local path: `D:\Document\文件处理\hainan-settlement-desktop`
- GitHub: `https://github.com/LeBronJu/hainan-settlement-desktop`
- Original Python baseline: `D:\Document\文件处理\hainan-settlement-tool`
- Original Python GitHub: `https://github.com/LeBronJu/hainan-settlement-tool`

The Python project is the historical full-function reference. New C# work should stay in this standalone repository.

## Current Branch And Worktree

- Current branch: `codex/ui-modernization`
- This branch contains the Win10/11 WPF UI work plus fixes that should also remain in the Win7/8 WinForms UI.
- Latest feature commit: `215dd81 Modernize settlement UI and validation flow`.
- Current product naming convention: call the two desktop entries `Win7/8版` and `Win10/11版`; use Windows-version labels in user-facing docs, package names, and handoffs.
- Do not add real ledgers, customer data, settlement outputs, screenshots, or finance/payment data to git.

Latest committed work includes:

- Stage 2 preflight/audit support.
- Win10/11 WPF UI refinements and preflight dialog.
- Stage 1 stale-clean-table fix.
- Stage 1 "clean power data only" action.
- Direct `.xls` raw-detail support for cleaning and customer-code reading.
- Release packaging script fix for WPF README encoding.

Use `git status --short --branch` before editing; the worktree is intentionally dirty.

## Functional State

### Stage 1

Inputs:

- Base ledger workbook (`.xlsx`).
- Existing power workbook (`.xlsx`), or raw retail detail (`.xlsx`, `.xls`, `.csv`).
- Optional reference ledger (`.xlsx`).
- Shared output folder.

Outputs:

- Updated ledger copy.
- JSON report.
- Optional cleaned power workbook: `零售侧用户电量数据处理表.xlsx`.

Recent behavior changes:

- If raw retail detail is selected, Stage 1 always regenerates the cleaned power workbook in the selected output folder. This avoids reusing a stale `零售侧用户电量数据处理表.xlsx` from the raw-source folder.
- A new "只清洗电量数据" action was added to both Win10/11 and Win7/8 versions. It cleans raw retail detail without updating the ledger.
- Newly inserted ledger rows are inserted above existing footer/formula rows instead of overwriting the first row after the last customer.
- Raw `.xls` detail is now supported via `ExcelDataReader`. The cleaned workbook output remains `.xlsx`.
- Customer-code lookup for newly added customers also supports `.xls` raw detail.

### Stage 2

Inputs:

- Manually reviewed current-month ledger.
- Previous-month agent split workbook folder.
- Previous-month intermediary split workbook folder.
- Previous or corrected summary workbook.
- Shared output folder.

Outputs:

- Agent split workbooks under `2026年代理 - 海南`.
- Intermediary split workbooks under `2026年居间 - 海南`.
- Monthly agent-fee summary workbook.
- JSON settlement report.
- `阶段二校验报告.txt`.

Current behavior:

- Stage 2 is template-driven. It copies prior-month sheets and writes current-month input/value cells while preserving template formatting, hidden columns, merged headers, blank cells, date display formats, and non-current formulas.
- Before generation, Stage 2 analyzes key changes and asks the user to confirm detailed items, not just counts.
- Preflight items include new agent/intermediary relationship, new customer in a split sheet, profit unit-price change, tax-rate change, and previous template read failures.
- During generation, Stage 2 audits split-table calculated results against ledger-derived values and writes differences to the validation report.
- The right-side Win10/11 progress panel and top status indicator are now used during Stage 2 preflight and execution.

Known limits:

- The app still uses ClosedXML formula writing/caches; it does not automate desktop Excel to force recalculation.
- January/February 2026 historical irregularities should not be generalized into new rules.
- Do not modify ledger customer names to match summary/payment-account names.
- `项目开发人` is an agent/intermediary relationship under a负责人, not the负责人 themselves.

## UI State

Two desktop entries exist:

- `src/HainanSettlementTool.WinForms`: Win7/8 UI on `.NET Framework 4.7.2`.
- `src/HainanSettlementTool.Wpf`: Win10/11 UI shell on `.NET Framework 4.7.2`.

The user prefers the Win10/11 UI direction, especially the right-side progress display. The left navigation from the earlier design mockup was intentionally not implemented.

Minimum-size protection has been added in the Win10/11 design so the main layout should remain visible instead of clipping when resized.

## Latest Verification

Debug build command:

```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" "D:\Document\文件处理\hainan-settlement-desktop\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

Latest observed result: `0 个警告 / 0 个错误`.

Release packages generated:

- Win10/11 build: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Wpf-Release-20260622-155838.zip`
- Win7/8 build: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Release-20260622-155855.zip`

Additional smoke check:

- A temporary `.xls` raw detail workbook was generated locally.
- `Stage1Service.CleanPowerData(...)` read it successfully and produced a cleaned `.xlsx`.
- Smoke result: `PowerRows=1`, `MonthTotal=12.3456`.
- Temporary smoke files were deleted.

`git diff --check` currently reports only line-ending warnings (`LF will be replaced by CRLF`), not whitespace errors.

## Important Files Changed

Core:

- `src/HainanSettlementTool.Core/Services/Stage1Service.cs`
- `src/HainanSettlementTool.Core/Services/Stage2Service.cs`
- `src/HainanSettlementTool.Core/Services/IStage2ExcelGateway.cs`
- `src/HainanSettlementTool.Core/Models/PowerCleanReport.cs`
- `src/HainanSettlementTool.Core/Models/Stage2CheckIssue.cs`
- `src/HainanSettlementTool.Core/Models/Stage2PreflightReport.cs`
- `src/HainanSettlementTool.Core/Models/Stage2Report.cs`
- `src/HainanSettlementTool.Core/Models/DetailSettlementRow.cs`

Excel:

- `src/HainanSettlementTool.Excel/HainanSettlementTool.Excel.csproj`
- `src/HainanSettlementTool.Excel/RawDetailReader.cs`
- `src/HainanSettlementTool.Excel/CustomerCodeReader.cs`
- `src/HainanSettlementTool.Excel/LedgerStage1Updater.cs`
- `src/HainanSettlementTool.Excel/Stage2SettlementGenerator.cs`
- `src/HainanSettlementTool.Excel/ClosedXmlStage1ExcelGateway.cs`

UI:

- `src/HainanSettlementTool.Wpf/MainWindow.xaml`
- `src/HainanSettlementTool.Wpf/MainWindow.xaml.cs`
- `src/HainanSettlementTool.Wpf/Stage2PreflightWindow.xaml`
- `src/HainanSettlementTool.Wpf/Stage2PreflightWindow.xaml.cs`
- `src/HainanSettlementTool.WinForms/MainForm.cs`

Packaging/docs:

- `scripts/package_wpf_release.ps1`
- `README.md`
- `AGENTS.md`
- `docs/architecture.md`
- `HANDOFF.md`

## Next Steps

1. Ask the user to verify the latest Win10/11 package with a real `.xls` raw retail detail file.
2. If accepted, commit the current branch intentionally. Suggested commit scope: Win10/11 UI, stage 2 validation, stage 1 clean-table fixes, `.xls` raw-detail support, docs.
3. Push `codex/ui-modernization` only after the user asks for it.
4. If the user wants main updated, merge only after confirming the long-term release strategy for Win7/8 and Win10/11 builds.
5. Longer-term: add sanitized fixture workbooks and automated regression tests for Stage 1 cleaning and Stage 2 audit/preflight behavior.

## Suggested Skills For Next Session

- `diagnose`: use for reported workbook/UI bugs.
- `spreadsheets:Spreadsheets`: use when inspecting workbook formatting or generated Excel outputs.
- `handoff`: use again before context compaction or branch handoff.
