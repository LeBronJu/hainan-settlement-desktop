# Handoff

Last updated: 2026-06-26

## Project

Standalone C# desktop rewrite of the Hainan retail electricity settlement automation tool.

- Local path: `D:\Document\文件处理\hainan-settlement-desktop`
- GitHub: `https://github.com/LeBronJu/hainan-settlement-desktop`
- Original Python baseline: `D:\Document\文件处理\hainan-settlement-tool`
- Original Python GitHub: `https://github.com/LeBronJu/hainan-settlement-tool`

The Python project remains the historical full-function reference. The C# desktop project is now versioned and released independently.

## Current Git State

- Current branch: `codex/stage2-summary-detail-template-fixes`
- Latest main/origin main commit: `7d43274 Merge branch 'codex/document-branch-rule'`
- Branch purpose: fix Stage 2 generated workbook template correctness around split-sheet totals/styles/dates and summary-sheet footer handling.
- Current release tag: `v1.0`
- Release page: `https://github.com/LeBronJu/hainan-settlement-desktop/releases/tag/v1.0`
- Win7/8 and Win10/11 entries are both part of `main`; they share Core/Excel logic but remain separate desktop apps.
- Do not add real ledgers, customer data, settlement outputs, screenshots, or finance/payment data to git.

Use `git status --short --branch` before editing. The expected post-release worktree is clean.

## Release 1.0

GitHub Release title:

- `海南售电结算自动化工具 v1.0`

Release assets:

- `HainanSettlementTool-Win7-8-v1.0.zip`
- `HainanSettlementTool-Win10-11-v1.0.zip`

Local package copies from the v1.0 build:

- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win7-8-v1.0.zip`
- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-v1.0.zip`

Notes:

- GitHub Release assets use ASCII names because one upload attempt with Chinese filename prefixes was normalized incorrectly by the GitHub tooling.
- Packaging scripts still generate timestamped package names under `dist/`; release assets can be copied/renamed from those timestamped zips.

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

Current behavior:

- If raw retail detail is selected, Stage 1 regenerates the cleaned power workbook in the selected output folder. It does not reuse a stale cleaned workbook from the raw-source folder.
- "只清洗电量数据" exists in both Win7/8 and Win10/11 versions. It cleans raw retail detail without updating the ledger.
- Raw `.xlsx`, `.xls`, and `.csv` detail are supported for power cleaning.
- Customer-code lookup uses the same raw detail row adapter for `.xlsx`, `.xls`, and `.csv`.
- If the same customer appears with multiple different customer codes, the app does not choose one automatically; the Stage 1 report should leave it for manual completion.
- Newly inserted ledger rows are inserted above existing footer/formula rows instead of overwriting the first row after the last customer.

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
- Split workbooks rewrite total-row formula ranges after row insert/delete. If the copied total row lacks table border/alignment formatting, generation repairs it from a prior month total row or, as a fallback, the last detail row.
- Split workbooks only update an existing bottom signature date after the total row; the date is shifted forward one month and no new bottom date is created when the template lacks one.
- Summary workbooks treat only rows before the `合计` row as subject data. Footer/signature/audit rows after `合计` must not be interpreted as summary subjects.
- Summary signature date defaults to the settlement month plus 2 months, day 8.
- Before generation, Stage 2 analyzes key changes and asks the user to confirm detailed items.
- Preflight items include new agent/intermediary relationship, new customer in a split sheet, profit unit-price change, tax-rate change, and previous template read failures.
- During generation, Stage 2 audits split-table calculated results against ledger-derived values and writes differences to the validation report.
- Stage 2 amount calculation and ledger/split-sheet difference issue generation are centralized in Core for testability.

Known limits:

- The app uses ClosedXML formula writing/caches; it does not automate desktop Excel to force recalculation.
- January/February 2026 historical irregularities should not be generalized into new rules.
- Do not modify ledger customer names to match summary/payment-account names.
- `项目开发人` is an agent/intermediary relationship under a负责人, not the负责人 themselves.

## UI State

Two desktop entries exist and should continue to coexist:

- `src/HainanSettlementTool.WinForms`: Win7/8 UI on `.NET Framework 4.7.2`.
- `src/HainanSettlementTool.Wpf`: Win10/11 UI shell on `.NET Framework 4.7.2`.

Both UI entries should remain thin shells for file selection, parameter input, confirmation, progress/log display, and error messages. Shared business behavior belongs in Core/Excel.

## Latest Verification

Commands run before `v1.0` release:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m
.\scripts\check_build_portability.ps1
.\scripts\package_release.ps1
.\scripts\package_wpf_release.ps1
```

Observed result:

- Core tests: 6 passed.
- Excel tests: 3 passed.
- Release build passed for Win7/8 and Win10/11.
- Packaging scripts produced both release zips.
- Build portability check passed.

Current Stage 2 template-fix branch verification on 2026-06-26:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m
.\scripts\check_build_portability.ps1
git diff --check
```

Observed result:

- Core tests: 6 passed.
- Excel tests: 5 passed, including synthetic Stage 2 workbook regressions for split-sheet total/date/style handling and summary footer protection.
- Release build passed for Win7/8 and Win10/11.
- Build portability check passed.
- `git diff --check` passed; Git only printed CRLF normalization warnings.

Authorized real `.xls` smoke check:

- File used locally: `D:\Document\文件处理\海南2026-4月代理费结算\零售侧明细结果.xls`
- The file has multiple sheets and includes `零售户号电量` and `零售主体电量`.
- `ReadRawPowerRows` read 123 raw power rows.
- `ReadCustomerCodes` read 64 customer-code mappings.
- `Stage1Service.CleanPowerData(...)` generated a temporary cleaned `.xlsx` successfully.
- Temporary smoke output was deleted; no real workbook was committed.

## Documentation Rule

Documentation is now part of the development contract:

- Update relevant docs in the same work session as behavior, packaging, release, workflow, architecture, or business-rule changes.
- Check `README.md`, `HANDOFF.md`, `CONTEXT.md`, and `docs/architecture.md` before finishing user-visible work.
- Put temporary investigations and architecture decisions in dated files under `docs/dev-notes/`.

## Useful Files

Core:

- `src/HainanSettlementTool.Core/Services/Stage1Service.cs`
- `src/HainanSettlementTool.Core/Services/Stage2Service.cs`
- `src/HainanSettlementTool.Core/Services/Stage2SettlementCalculator.cs`
- `src/HainanSettlementTool.Core/Services/FileAccessGuard.cs`

Excel:

- `src/HainanSettlementTool.Excel/RawDetailRowReader.cs`
- `src/HainanSettlementTool.Excel/RawDetailReader.cs`
- `src/HainanSettlementTool.Excel/CustomerCodeReader.cs`
- `src/HainanSettlementTool.Excel/LedgerStage1Updater.cs`
- `src/HainanSettlementTool.Excel/Stage2SettlementGenerator.cs`

UI:

- `src/HainanSettlementTool.WinForms/MainForm.cs`
- `src/HainanSettlementTool.Wpf/MainWindow.xaml`
- `src/HainanSettlementTool.Wpf/MainWindow.xaml.cs`
- `src/HainanSettlementTool.Wpf/Stage2PreflightWindow.xaml`

Packaging/docs:

- `scripts/package_release.ps1`
- `scripts/package_wpf_release.ps1`
- `scripts/check_build_portability.ps1`
- `README.md`
- `AGENTS.md`
- `CONTEXT.md`
- `docs/architecture.md`

## Next Steps

1. Review and merge `codex/stage2-summary-detail-template-fixes` before continuing Stage 2 workbook-quality work.
2. Use copied/sanitized workbooks for any real-data smoke; never write to the production `C:\Users\juqx2\Desktop\2026海南` tree.
3. Next architecture slice, if desired: extract a shared workflow module so Win7/8 and Win10/11 do not duplicate stage execution flow.
4. Consider adding sanitized Stage 2 fixture workbooks later; current regressions use dynamically generated synthetic workbooks.
