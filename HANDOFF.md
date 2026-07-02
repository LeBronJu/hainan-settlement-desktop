# Handoff

Last updated: 2026-07-02

## Project

Standalone C# desktop rewrite of the Hainan retail electricity settlement automation tool.

- Local path: `D:\Document\文件处理\hainan-settlement-desktop`
- GitHub: `https://github.com/LeBronJu/hainan-settlement-desktop`
- Original Python baseline: `D:\Document\文件处理\hainan-settlement-tool`
- Original Python GitHub: `https://github.com/LeBronJu/hainan-settlement-tool`
- Stable local reference folder: `D:\Document\文件处理\稳定参考版海南结算`

The Python project remains the historical full-function reference. The C# desktop project is now versioned and released independently.
The stable local reference folder is for future comparison/orientation only; do not read its real workbook contents without explicit user authorization.

## Real Data And Local Reference Paths

- Real production environment root: `C:\Users\juqx2\Desktop\2026海南`
- Real production workbooks include issued ledgers, settlement sheets, manually corrected outputs, customer data, screenshots, and sensitive financial results. Treat every workbook under this root as real production data.
- On 2026-07-02 the user authorized read-only inspection of real production and reference files for the employee reward analysis. This does not grant write permission.
- Do not read files under the real production root unless the user explicitly authorizes the specific file, folder, broad read-only analysis scope, or smoke scope for the current task. Prior authorizations in old conversation context should not be treated as write permission.
- Never modify files in the real production root in place. If the user authorizes a real smoke or comparison, write generated outputs only to an explicitly selected test/output folder and do not overwrite production files.
- Known test/output area used by the user: `C:\Users\juqx2\Desktop\2026海南\test`. Files there may still contain real or sensitive data; do not commit them.
- Stable local reference folder: `D:\Document\文件处理\稳定参考版海南结算`. It is for future comparison/orientation only; do not read real workbook contents there without explicit user authorization for the current task.
- Employee reward reference folder recorded by the user on 2026-07-02: `D:\Document\文件处理\稳定参考版海南结算\电量奖励参考`. It was inspected read-only for the employee reward design; the implemented module uses the files as output-shape references, not runtime inputs.
- Repository-safe real smoke entry point: `scripts/run_real_smoke.ps1`. It accepts paths as parameters and must not hard-code real production paths.

## Current Git State

- Current branch: `codex/employee-reward-module`
- Active development branch `codex/employee-reward-module` implements the independent `员工电量奖励` module for Win10/11 WPF, Core, and Excel. The user completed practical testing on 2026-07-02 and reported no blocking issues. Do not merge or release until the user explicitly asks for that next step.
- Quality branch `codex/real-smoke-runner` has been reviewed and merged.
- Stage 2 workbook template fixes have been merged from `codex/stage2-summary-detail-template-fixes`.
- Stage 2 special-row/template cleanup has been merged to `main`. It fixes new-subject borrowed-template history, adds preflight handling for previous-month special detail rows, and improves the WPF preflight confirmation layout.
- Latest merged Stage 2 fix commit before this handoff update: `d8cefbd Document stage two real comparison outcome`
- Current release tag: `v1.0.1`
- Release page: `https://github.com/LeBronJu/hainan-settlement-desktop/releases/tag/v1.0.1`
- Stage 1 ledger workbook tests and documentation impact gate cleanup have been integrated from `codex/stage1-ledger-tests`.
- Win10/11 WPF is the primary UI entry for new features and UX work.
- Win7/8 WinForms remains part of `main` as a maintenance compatibility entry: keep it buildable, packageable, and fix blocking bugs only unless explicitly requested.
- Win7/8 and Win10/11 share Core/Excel logic but remain separate desktop apps.
- Do not add real ledgers, customer data, settlement outputs, screenshots, or finance/payment data to git.

Use `git status --short --branch` before editing. The expected handoff worktree may be dirty on `codex/employee-reward-module` with employee reward code, tests, and documentation changes; no real Excel files or generated settlement outputs should be tracked.

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
- When a new agent/intermediary subject has no matching prior workbook and must borrow a same-kind template, the output copy keeps only the generated current-month sheet so donor history is not carried into the new subject workbook.
- Split workbooks rewrite total-row formula ranges after row insert/delete. If the copied total row lacks table border/alignment formatting, generation repairs it from a prior month total row or, as a fallback, the last detail row.
- Split workbooks only update an existing bottom signature date after the total row; text dates and Excel date-valued cells are shifted forward one month, and no new bottom date is created when the template lacks one.
- Summary workbooks treat only rows before the `合计` row as subject data. Footer/signature/audit rows after `合计` must not be interpreted as summary subjects.
- Summary signature date defaults to the settlement month plus 2 months, day 8.
- Before generation, Stage 2 analyzes key changes and asks the user to confirm detailed items.
- Preflight items include new agent/intermediary relationship, new customer in a split sheet, previous split-sheet detail rows absent from the current ledger, profit unit-price change, tax-rate change, and previous template read failures.
- Previous split-sheet special detail rows, such as manual refund/withholding/correction lines, remain in the historical sheet but are not inherited into the generated current-month sheet. If the current month still needs the adjustment, the user must manually update the generated split sheet and summary.
- During generation, Stage 2 audits split-table calculated results against ledger-derived values and writes differences to the validation report.
- Stage 2 amount calculation and ledger/split-sheet difference issue generation are centralized in Core for testability.

Known limits:

- The app uses ClosedXML formula writing/caches; it does not automate desktop Excel to force recalculation.
- January/February 2026 historical irregularities should not be generalized into new rules.
- Do not modify ledger customer names to match summary/payment-account names.
- `项目开发人` is an agent/intermediary relationship under a负责人, not the负责人 themselves.

### Employee Reward

Inputs:

- Latest Hainan settlement ledger workbook (`.xlsx`).
- Start month and end month.
- Shared output folder.

Outputs:

- Employee reward summary workbook.
- One employee power confirmation workbook per负责人.
- JSON validation/report file.

Current behavior:

- The module is exposed in the Win10/11 WPF app as the `员工电量奖励` tab.
- The module only needs the latest ledger and month range; the reference workbooks under `稳定参考版海南结算\电量奖励参考` are output-shape references, not runtime inputs.
- The ledger sheet is identified by `海南2026年售电结算台账` first, then by required row-2 headers as fallback.
- Fixed columns are found by row-2 headers. Month power columns are found by row-1 `X月` plus row-2 `总实际电量（万千瓦时）`, including hidden columns.
- Rewards aggregate by `负责人`, not by `项目开发人`.
- Blank helper/check rows with no customer code, no customer name, and no负责人 are excluded.
- Missing负责人, duplicate customer code, and empty企业名称 with selected-period power stop generation as serious ledger errors.
- Outputs use formulas for detail totals, employee summary totals, and reward amount (`电量合计 * 10000 * 0.0001`).
- Existing output files are not overwritten; timestamped unique filenames are used when needed.

Known limits:

- The first implementation generates an internal layout based on the reference workbooks instead of copying the reference templates. User visual review is still needed before release.
- Win7/8 WinForms has no employee reward UI entry; this follows the maintenance-only WinForms policy.

## UI State

Two desktop entries exist and should continue to coexist:

- `src/HainanSettlementTool.WinForms`: Win7/8 maintenance UI on `.NET Framework 4.7.2`.
- `src/HainanSettlementTool.Wpf`: Win10/11 primary UI shell on `.NET Framework 4.7.2`.

Both UI entries should remain thin shells for file selection, parameter input, confirmation, progress/log display, and error messages. Shared business behavior belongs in Core/Excel.
The current workflow extraction branch adds Core `SettlementWorkflow` so both UI entries reuse the same stage completion summary rules while keeping UI-specific confirmation, progress, and error display local.
Do not add WinForms-only features or UX improvements by default. New UI work should target WPF unless the user explicitly asks for Win7/8 support.
Current Stage 2 workflow deepening keeps WPF responsible for the confirmation dialog and progress UI, while Core `SettlementWorkflow` owns the preflight plan and the confirmed/cancelled generation decision.
The employee reward module currently exists only in the WPF app as a separate tab. It reuses the shared output folder and has its own start/end month selectors.
The WPF app supports UI theme selection in the custom title bar: `跟随系统`, `浅色`, and `深色`. The setting is stored in the existing WPF input snapshot XML. This is UI-only; generated Excel workbooks remain theme-independent, light, and print-safe.

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

Authorized real 2026-05 smoke rerun on 2026-06-29:

- Output folder: `C:\Users\juqx2\Desktop\2026海南\test\codex-real-smoke-20260629-120157`
- Scope: read real 2026-05 production inputs and write generated files only under the test smoke folder.
- Stage 1 raw `.xls` cleaning matched the existing production cleaned workbook: 69 rows, no missing/extra customer keys, and no power value differences.
- Stage 1 ledger update smoke completed: 69 power rows, 59 matched rows, 10 new rows, 10 missing manual-info rows, and 0 missing codes.
- Stage 2 smoke completed: 16 proxy workbooks, 3 intermediary workbooks, 1 summary workbook, and 0 formula error text hits across generated `.xlsx` files.
- Stage 2 validation report still contains 1 new-template prompt and 7 amount audit differences at `0.0001` 万元 scale. The user confirmed those differences are rounding differences and not blocking; the current generator writes summary amounts from split-sheet self-calculation.
- Stable local reference folder recorded for future comparison: `D:\Document\文件处理\稳定参考版海南结算`. Do not read real workbook contents there without explicit user authorization.

Real smoke runner branch verification on 2026-06-29:

- Added `scripts/run_real_smoke.ps1`.
- The runner accepts explicit input paths and does not hard-code real workbook directories.
- Smoke output from script validation: `C:\Users\juqx2\Desktop\2026海南\test\real-smoke-20260629-141431`
- Result matched the previous manual 2026-05 smoke: 69 cleaned rows, 0 power comparison differences, 59 matched stage 1 rows, 10 new stage 1 rows, 16 proxy workbooks, 3 intermediary workbooks, 1 summary workbook, 8 stage 2 audit issues including the known rounding/template prompts, and 0 formula error text hits.

Stage 2 special-row/template cleanup branch verification on 2026-06-30:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\check_build_portability.ps1
git diff --check
.\scripts\package_wpf_release.ps1
```

Observed result:

- Core tests: 12 passed.
- Excel tests: 12 passed, including new synthetic regressions for borrowed-template history cleanup, previous-month special detail row preflight, and not carrying special rows into the generated current month.
- Debug build passes for Core, Excel, WinForms, and WPF.
- Build portability check passes.
- `git diff --check` passes; Git only prints CRLF normalization warnings.
- Win10/11 business-acceptance package: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260630-115250.zip`.
- User ran practical testing on 2026-07-01 and reported no issues found.

Final local verification before committing the Stage 2 special-row/template cleanup branch on 2026-07-01:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\check_build_portability.ps1
git diff --check
.\scripts\package_wpf_release.ps1
```

Observed result:

- Core tests: 12 passed.
- Excel tests: 12 passed.
- Debug build passes for Core, Excel, WinForms, and WPF.
- Build portability check passes.
- `git diff --check` passes; Git only prints CRLF normalization warnings.
- Win10/11 final local package: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260701-100344.zip`.

Main branch verification after merging Stage 2 special-row/template cleanup on 2026-07-01:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m
.\scripts\check_build_portability.ps1
git diff --check
.\scripts\package_wpf_release.ps1
```

Observed result:

- Core tests: 12 passed.
- Excel tests: 12 passed.
- Release build passes for Core, Excel, WinForms, and WPF.
- Build portability check passes.
- `git diff --check` passes; Git only prints a CRLF normalization warning for `HANDOFF.md`.
- Win10/11 main-branch local package: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260701-102428.zip`.

Employee reward module branch verification on 2026-07-02:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m
.\scripts\check_build_portability.ps1
git diff --check
```

Observed result:

- Core tests: 16 passed, including employee reward aggregation, serious ledger error handling, single-month range support, and workflow summary lines.
- Excel tests: 14 passed, including employee reward ledger sheet/month-column detection, hidden month column reading, helper row exclusion, non-overwrite output naming, formula generation, and personal confirmation workbook creation.
- Debug and Release builds pass for Core, Excel, WinForms, and WPF.
- Build portability check passes.
- `git diff --check` passes; Git only prints CRLF normalization warnings.
- A temporary real-data smoke test was run against the read-only 2026-05 production ledger. It generated employee reward outputs in a temporary directory, verified summary/report/personal workbook creation, and deleted the temporary output. No production workbook was modified and no real path test file was kept.

WPF employee reward tab and dropdown polish on 2026-07-02:

```powershell
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
.\scripts\package_wpf_release.ps1
```

Observed result:

- Debug build passes for Core, Excel, WinForms, and WPF.
- Core tests: 16 passed.
- Excel tests: 14 passed.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-121703`.
- Zip package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-121703.zip`.
- The WPF tab selector was restyled from the default WPF tab chrome to a project-native segmented control style.
- All WPF `ModernComboBox` dropdowns now use a custom modern template instead of the default system ComboBox chrome.

WPF small-window layout fix on 2026-07-02:

```powershell
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
.\scripts\package_wpf_release.ps1
```

Observed result:

- Debug build passes for Core, Excel, WinForms, and WPF.
- Core tests: 16 passed.
- Excel tests: 14 passed.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-122403`.
- Zip package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-122403.zip`.
- The main feature tab contents and the right-side completion card now have vertical scroll containers, so controls are not clipped when available window height is smaller or display scaling is high.

WPF shared-month selector guard on 2026-07-02:

```powershell
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
.\scripts\package_wpf_release.ps1
```

Observed result:

- Debug build passes for Core, Excel, WinForms, and WPF.
- Core tests: 16 passed.
- Excel tests: 14 passed.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-123328`.
- Zip package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-123328.zip`.
- When the WPF `员工电量奖励` tab is selected, the top shared settlement-month selector is disabled and the shared output-folder selector remains enabled.

Employee reward workbook layout fix on 2026-07-02:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\package_wpf_release.ps1
```

Observed result:

- Core tests: 16 passed.
- Excel tests: 14 passed, including employee reward workbook layout checks for detail-sheet borders through the total row, personal confirmation borders through the total row, summary-sheet period names such as `1月-4月员工电量汇总`, and the summary total row before the reward note.
- Debug build passes for Core, Excel, WinForms, and WPF.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-130042`.
- Zip package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-130042.zip`.

Employee reward visible border follow-up on 2026-07-02:

```powershell
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter EmployeeRewardGeneratorTests
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\package_wpf_release.ps1
```

Observed result:

- The real user test output from `C:\Users\juqx2\Desktop\2026海南\电量奖\1-4月测试` was inspected read-only. Its workbook XML and Excel COM both showed borders existed, but they used the light `#8FA1A8` color and could look like no visible border in Excel.
- Employee reward table borders now write each cell's top, bottom, left, and right borders explicitly as black thin borders instead of relying on range inside/outside borders.
- Core tests: 16 passed.
- Excel tests: 14 passed, including strengthened assertions that employee reward table borders are black thin borders.
- Debug build passes for Core, Excel, WinForms, and WPF.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-130840`.
- Zip package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-130840.zip`.

WPF UI theme support on 2026-07-02:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\package_wpf_release.ps1
```

Observed result:

- Win10/11 WPF title bar now has a `主题` selector with `跟随系统`, `浅色`, and `深色`.
- `跟随系统` reads the Windows app theme registry value and re-applies the palette when system preferences change.
- Theme choice is saved in the existing WPF input snapshot XML.
- Theme resources cover the main WPF shell, title bar, cards, inputs, buttons, tabs, progress area, completion area, and log area.
- Generated Excel workbooks are not theme-aware and remain fixed light/print-safe outputs.
- Core tests: 16 passed.
- Excel tests: 14 passed.
- Debug build passes for Core, Excel, WinForms, and WPF.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-135657`.
- Zip package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260702-135657.zip`.
- The user practical-tested this WPF build and reported no issues with the implemented employee reward and theme behavior.

Final local closeout on 2026-07-02:

```powershell
git diff --check
git status --short | Select-String -Pattern '\.(xlsx|xls|csv|png|jpg|jpeg|pdf)$'
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
```

Observed result:

- `git diff --check` passed; Git only printed CRLF normalization warnings.
- No real Excel, CSV, image, PDF, or generated sensitive output file appears in `git status`.
- Core tests: 16 passed.
- Excel tests: 14 passed.
- Debug build passes for Core, Excel, WinForms, and WPF.

## Documentation Rule

Documentation is now part of the development contract:

- At the start of each project work session, run `git status --short --branch` and read `AGENTS.md` plus this `HANDOFF.md` before editing.
- Read the owning document before changing a responsibility area: `CONTEXT.md` for settlement rules, `docs/architecture.md` for module seams or workflow structure, `README.md` for user-facing setup/package status, and `docs/RELEASE_CHECKLIST.md` for release or packaging.
- Repeat the relevant reading gate after context compaction, a long pause, or a task direction change.
- This is a local single-developer project. Pull requests are optional; after a feature branch is committed and pushed, local merge to `main` is acceptable when the user authorizes it.
- Every code, config, script, packaging, business-rule, UI-behavior, test-workflow, or task-state change must end with a documentation impact judgment.
- Update only documents whose responsibility is affected; do not rewrite unaffected docs for process compliance.
- Business-rule changes must check `CONTEXT.md`; module-boundary changes must check `AGENTS.md` and an ADR or dev note; release/packaging changes must check `README.md` and `docs/RELEASE_CHECKLIST.md`; branch state, validation results, or next steps must check this `HANDOFF.md`.
- Before finishing a change, record whether docs were affected. If yes, list the updated docs; if no, state why no doc update was needed.
- Keep this `HANDOFF.md` current whenever branch state, release status, validation results, or next steps change.
- If a pull request is used, use `.github/PULL_REQUEST_TEMPLATE.md` as the merge-time documentation checklist. If merging locally without a PR, apply the same documentation-impact and validation checks manually.
- Put temporary investigations and architecture decisions in dated files under `docs/dev-notes/`.

## Useful Files

Core:

- `src/HainanSettlementTool.Core/Services/Stage1Service.cs`
- `src/HainanSettlementTool.Core/Services/Stage2Service.cs`
- `src/HainanSettlementTool.Core/Services/SettlementWorkflow.cs`
- `src/HainanSettlementTool.Core/Services/Stage2WorkflowPlan.cs`
- `src/HainanSettlementTool.Core/Services/Stage2WorkflowResult.cs`
- `src/HainanSettlementTool.Core/Services/Stage2SettlementCalculator.cs`
- `src/HainanSettlementTool.Core/Services/EmployeeRewardService.cs`
- `src/HainanSettlementTool.Core/Services/IEmployeeRewardExcelGateway.cs`
- `src/HainanSettlementTool.Core/Services/FileAccessGuard.cs`

Excel:

- `src/HainanSettlementTool.Excel/RawDetailRowReader.cs`
- `src/HainanSettlementTool.Excel/RawDetailReader.cs`
- `src/HainanSettlementTool.Excel/CustomerCodeReader.cs`
- `src/HainanSettlementTool.Excel/LedgerStage1Updater.cs`
- `src/HainanSettlementTool.Excel/Stage2SettlementGenerator.cs`
- `src/HainanSettlementTool.Excel/EmployeeRewardGenerator.cs`

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

1. Review the full branch diff, then commit and push `codex/employee-reward-module` when ready.
2. When the user authorizes merge, merge locally to `main` without requiring a PR, then run the documented validation checks.
3. Decide whether to cut a Win10/11 acceptance release from the accepted WPF package or rebuild a fresh release package after commit/merge.
4. Continue quality work with WPF as the default UI target; avoid WinForms parity work unless it is a bugfix, build/package compatibility issue, or explicitly requested.
5. Consider adding sanitized employee reward fixture workbooks later; current regressions use dynamically generated synthetic workbooks and a local temporary real smoke.
6. Consider adding sanitized Stage 2 fixture workbooks later; current regressions use dynamically generated synthetic workbooks.
7. Consider adding a sanitized `.xls` fixture later; real `.xls` smoke passed, but the repository still has no committed `.xls` regression fixture.
