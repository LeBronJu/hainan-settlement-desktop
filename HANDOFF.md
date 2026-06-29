# Handoff

Last updated: 2026-06-29

## Project

Standalone C# desktop rewrite of the Hainan retail electricity settlement automation tool.

- Local path: `D:\Document\文件处理\hainan-settlement-desktop`
- GitHub: `https://github.com/LeBronJu/hainan-settlement-desktop`
- Original Python baseline: `D:\Document\文件处理\hainan-settlement-tool`
- Original Python GitHub: `https://github.com/LeBronJu/hainan-settlement-tool`

The Python project remains the historical full-function reference. The C# desktop project is now versioned and released independently.

## Current Git State

- Current branch: `main`
- Stage 2 workbook template fixes have been merged from `codex/stage2-summary-detail-template-fixes`.
- Latest merged Stage 2 fix commit before this handoff update: `d8cefbd Document stage two real comparison outcome`
- Current release tag: `v1.0.1`
- Release page: `https://github.com/LeBronJu/hainan-settlement-desktop/releases/tag/v1.0.1`
- Stage 1 ledger workbook tests and documentation impact gate cleanup have been integrated from `codex/stage1-ledger-tests`.
- Win10/11 WPF is the primary UI entry for new features and UX work.
- Win7/8 WinForms remains part of `main` as a maintenance compatibility entry: keep it buildable, packageable, and fix blocking bugs only unless explicitly requested.
- Win7/8 and Win10/11 share Core/Excel logic but remain separate desktop apps.
- Do not add real ledgers, customer data, settlement outputs, screenshots, or finance/payment data to git.

Use `git status --short --branch` before editing. The expected handoff worktree is clean.

## Release 1.0.1

GitHub Release title:

- `海南售电结算自动化工具 v1.0.1`

Release assets:

- `HainanSettlementTool-Win7-8-v1.0.1.zip`
- `HainanSettlementTool-Win10-11-v1.0.1.zip`

Local package copies:

- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win7-8-v1.0.1.zip`
- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-v1.0.1.zip`

Notes:

- `v1.0.1` publishes the Stage 2 workbook template fixes from `main`.
- The release keeps the existing Win7/8 and Win10/11 split packages. Future releases may continue shipping Win7/8 as a maintenance compatibility package, while new UI features default to Win10/11.

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
- Split workbooks only update an existing bottom signature date after the total row; text dates and Excel date-valued cells are shifted forward one month, and no new bottom date is created when the template lacks one.
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

- `src/HainanSettlementTool.WinForms`: Win7/8 maintenance UI on `.NET Framework 4.7.2`.
- `src/HainanSettlementTool.Wpf`: Win10/11 primary UI shell on `.NET Framework 4.7.2`.

Both UI entries should remain thin shells for file selection, parameter input, confirmation, progress/log display, and error messages. Shared business behavior belongs in Core/Excel.
The current workflow extraction branch adds Core `SettlementWorkflow` so both UI entries reuse the same stage completion summary rules while keeping UI-specific confirmation, progress, and error display local.
Do not add WinForms-only features or UX improvements by default. New UI work should target WPF unless the user explicitly asks for Win7/8 support.
Current Stage 2 workflow deepening keeps WPF responsible for the confirmation dialog and progress UI, while Core `SettlementWorkflow` owns the preflight plan and the confirmed/cancelled generation decision.

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
- Excel tests: 3 passed in `v1.0`; later quality branches added Stage 1 ledger workbook tests and Stage 2 workbook regressions.
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
- Excel tests: 6 passed, including synthetic Stage 2 workbook regressions for split-sheet total/date/style handling, Excel date-valued signature dates, summary footer protection, and current-month actual-payment headers.
- Release build passed for Win7/8 and Win10/11.
- Build portability check passed.
- `git diff --check` passed; Git only printed CRLF normalization warnings.

Authorized real Stage 2 read-only comparison on 2026-06-26:

- New generated folder inspected: `C:\Users\juqx2\Desktop\2026海南\test`
- Manually corrected production folder inspected: `C:\Users\juqx2\Desktop\2026海南\海南2026-4月代理费结算`
- Scope: Stage 2 `.xlsx` outputs only; ledger, raw detail, confirmation form, PDF, backups, and reports were not modified.
- File set matched for Stage 2 workbooks: 19 split workbooks plus 1 summary workbook.
- Split formulas matched after formula normalization. The real comparison exposed one date-valued bottom signature date that was not shifted and one missing summary `当月实际支付` header; both are now covered by tests and fixed in code.

Follow-up comparison after rebuilding and regenerating Stage 2 output:

- New generated folder inspected: `C:\Users\juqx2\Desktop\2026海南\test\test2`
- The Stage 2 workbook file set still matched production: 19 split workbooks plus 1 summary workbook.
- Visible summary headers and summary total values matched the manually corrected production workbook.
- Remaining differences were reviewed with the user: formula-backed cells in the new output versus static values in the hand-edited workbook, more complete borders/styles in the new output, and bottom split-sheet dates shifted to `2026年6月8日` where the template already had a date. The user confirmed the monthly date shift is correct because only some agents require dated settlement sheets.
- No further blocking Stage 2 output issue is known on this branch.

Main branch package build after merging Stage 2 fixes:

- `dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug` passed: Core 6 tests, Excel 6 tests.
- `dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m` passed.
- `.\scripts\check_build_portability.ps1` passed.
- Win10/11 package: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260626-141009.zip`
- Win7/8 package: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win7-8-Release-20260626-141017.zip`

Release `v1.0.1` verification on 2026-06-29:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m
.\scripts\check_build_portability.ps1
git diff --check
.\scripts\package_release.ps1
.\scripts\package_wpf_release.ps1
```

Observed result:

- Core tests: 6 passed.
- Excel tests: 6 passed.
- Release build passes for Win7/8 and Win10/11.
- Build portability check passes.
- `git diff --check` passes; Git only prints CRLF normalization warnings.
- Both release zips are copied to stable `v1.0.1` asset names under `dist/`.

Stage 1 ledger test branch integration on 2026-06-29:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
```

Observed result:

- Core tests: 6 passed.
- Excel tests: 9 passed, including the Stage 1 ledger workbook tests for inserting new customers before footer rows, filling only unambiguous customer codes, and preserving the base ledger by writing an output copy.

Workflow extraction branch verification on 2026-06-29:

```powershell
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
```

Observed result:

- Debug build passes for Core, Excel, WinForms, and WPF.
- Core tests: 9 passed, including shared `SettlementWorkflow` summary tests.
- Excel tests: 9 passed.

Stage 2 workflow deepening verification on 2026-06-29:

```powershell
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
```

Observed result:

- Debug build passes for Core, Excel, WinForms, and WPF.
- Core tests: 12 passed, including Stage 2 preflight continue/cancel workflow tests.
- Excel tests: 9 passed.

Authorized real `.xls` smoke check:

- File used locally: `D:\Document\文件处理\海南2026-4月代理费结算\零售侧明细结果.xls`
- The file has multiple sheets and includes `零售户号电量` and `零售主体电量`.
- `ReadRawPowerRows` read 123 raw power rows.
- `ReadCustomerCodes` read 64 customer-code mappings.
- `Stage1Service.CleanPowerData(...)` generated a temporary cleaned `.xlsx` successfully.
- Temporary smoke output was deleted; no real workbook was committed.

## Documentation Rule

Documentation is now part of the development contract:

- At the start of each project work session, run `git status --short --branch` and read `AGENTS.md` plus this `HANDOFF.md` before editing.
- Read the owning document before changing a responsibility area: `CONTEXT.md` for settlement rules, `docs/architecture.md` for module seams or workflow structure, `README.md` for user-facing setup/package status, and `docs/RELEASE_CHECKLIST.md` for release or packaging.
- Repeat the relevant reading gate after context compaction, a long pause, or a task direction change.
- Every code, config, script, packaging, business-rule, UI-behavior, test-workflow, or task-state change must end with a documentation impact judgment.
- Update only documents whose responsibility is affected; do not rewrite unaffected docs for process compliance.
- Business-rule changes must check `CONTEXT.md`; module-boundary changes must check `AGENTS.md` and an ADR or dev note; release/packaging changes must check `README.md` and `docs/RELEASE_CHECKLIST.md`; branch state, validation results, or next steps must check this `HANDOFF.md`.
- Before finishing a change, record whether docs were affected. If yes, list the updated docs; if no, state why no doc update was needed.
- Keep this `HANDOFF.md` current whenever branch state, release status, validation results, or next steps change.
- Use `.github/PULL_REQUEST_TEMPLATE.md` as the merge-time documentation checklist.
- Put temporary investigations and architecture decisions in dated files under `docs/dev-notes/`.

## Useful Files

Core:

- `src/HainanSettlementTool.Core/Services/Stage1Service.cs`
- `src/HainanSettlementTool.Core/Services/Stage2Service.cs`
- `src/HainanSettlementTool.Core/Services/SettlementWorkflow.cs`
- `src/HainanSettlementTool.Core/Services/Stage2WorkflowPlan.cs`
- `src/HainanSettlementTool.Core/Services/Stage2WorkflowResult.cs`
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
- `docs/RELEASE_CHECKLIST.md`
- `README.md`
- `AGENTS.md`
- `CONTEXT.md`
- `docs/architecture.md`
- `docs/RELEASE_CHECKLIST.md`

## Next Steps

1. Continue quality work with WPF as the default UI target; avoid WinForms parity work unless it is a bugfix, build/package compatibility issue, or explicitly requested.
2. Use the `v1.0.1` release packages for business-side acceptance.
3. Consider adding sanitized Stage 2 fixture workbooks later; current regressions use dynamically generated synthetic workbooks.
4. Consider adding a sanitized `.xls` fixture later; real `.xls` smoke passed, but the repository still has no committed `.xls` regression fixture.
