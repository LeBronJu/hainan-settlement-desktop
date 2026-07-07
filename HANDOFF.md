# Handoff

Last updated: 2026-07-07

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
- Chongqing local working root mentioned by the user: `C:\Users\juqx2\Desktop\2026年-重庆`. Treat files under this root as real business data. On 2026-07-06 the user authorized read-only inspection of one Chongqing transaction-center source workbook, one manually cleaned comparison workbook, and the current Chongqing ledger workbook for designing the Chongqing Stage 1 cleaning and ledger-update module. Do not commit those files or generated outputs.

## Current Git State

- Current branch: `codex/wpf-log-controller`
- Branch purpose: follow-up WPF quality slice that extracts run-log append/clear/save behavior from `MainWindow.xaml.cs` into `MainWindowLogController`. This branch is based on local `main` after `0b5ef73 Document main acceptance package`.
- Local `main` is ahead of `origin/main` after merging `codex/wpf-path-picker-controller`; remote push has not been performed in this session.
- Merged branch: `codex/wpf-path-picker-controller` through `c2ce2b5 Add Chongqing customer decisions`, with local merge commit `4911e47 Merge Chongqing stage one WPF updates`.
- Branch purpose now complete: Chongqing Stage 1 power cleaning/ledger update, WPF province UI, WPF controller decomposition slices, Chongqing target-month block copy, customer-resolution decisions, and WPF display-title fixes are on local `main`.
- Previous dialog branch: `codex/wpf-dialog-controller`, pushed through `191b0ff Extract WPF dialog controller`.
- Previous naming branch: `codex/province-neutral-naming`, pushed through `2f93013 Clarify province-neutral naming boundaries`.
- The Chongqing target-month block fix is isolated on `codex/chongqing-month-block-copy` and was pushed through `e45d358 Document WPF Chongqing test package`.
- Multi-province readiness note: `docs/dev-notes/multi-province-readiness-2026-07-07.md`. Read it before new-province onboarding, WPF province UI, Core multi-province workflow, or Excel multi-province adapter work.
- Previous uncommitted WPF small-window work was reviewed on 2026-07-06. The action-row `DockPanel LastChildFill="False"` fixes were already present on the Chongqing branch, the remaining `MinHeight="720"` fix was reapplied, and the old stash was dropped.
- Employee reward module has been merged to `main` from `codex/employee-reward-module`. The user completed practical testing on 2026-07-02 and reported no blocking issues.
- Latest employee reward feature commit before this handoff update: `feb933f Add employee reward module`
- `main` now contains the employee reward module, WPF theme support, and Chongqing Stage 1 power cleaning/ledger update after `v1.0.1`, but no newer release tag has been cut yet.
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
- Latest local acceptance packages built from local `main`:
  - Win10/11 WPF: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-140505.zip`
  - Win10/11 WPF unpacked directory: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-140505`
  - Win7/8 WinForms maintenance package: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win7-8-Release-20260707-140515.zip`
  - Win7/8 WinForms unpacked directory: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win7-8-Release-20260707-140515`
- Previous local WPF test package from `codex/wpf-path-picker-controller`:
  - `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-115127.zip`
  - unpacked directory: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-115127`
  - built after implementing Chongqing customer-resolution decisions and WPF display-title fixes.
- Previous local WPF test package from the Chongqing month-block branch:
  - `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-105743.zip`
  - unpacked directory: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-105743`
  - built from `codex/chongqing-month-block-copy` after `074401a` and `f60d325`.
- Do not add real ledgers, customer data, settlement outputs, screenshots, or finance/payment data to git.

Use `git status --short --branch` before editing. The expected handoff worktree should be clean on local `main`; no real Excel files or generated settlement outputs should be tracked.

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

- The implementation generates an internal layout based on the reference workbooks instead of copying the reference templates. The user completed practical testing on 2026-07-02 and reported no blocking issues, but a formal release has not been cut yet.
- Win7/8 WinForms has no employee reward UI entry; this follows the maintenance-only WinForms policy.

### Chongqing Stage 1 Power Cleaning And Ledger Update

Inputs:

- Chongqing trading-center electricity confirmation statement (`.xlsx`, `.xls`, or `.csv`).
- Chongqing settlement ledger (`.xlsx`) when updating the ledger.
- Settlement month from the file title/name, with the UI month as fallback.
- Shared output folder.

Outputs:

- `x月重庆零售侧用户电量数据处理表.xlsx`.
- `x月重庆零售侧用户电量校验报告.json`.
- `x月重庆售电结算台账-阶段一更新.xlsx`.
- `x月重庆阶段一台账更新报告.json`.

Current behavior:

- The module is exposed only in the Win10/11 WPF app through the new province selector. Selecting `重庆` shows the Stage 1 account-update area and enables `清洗并更新台账` plus `只清洗电量数据`.
- Excel input prefers the `sheet1` worksheet and falls back to the first sheet when `sheet1` is absent. CSV input is read as a single table.
- Required headers are `用户名称`, `户号`, `时段`, and `用电量`.
- Unit is `兆瓦时`; this must remain separate from Hainan's `万千瓦时` Stage 1 ledger unit.
- Period values map as `尖峰` -> `尖`, `高峰` -> `峰`, `平段` -> `平`, `低谷` -> `谷`.
- Output main sheet `用户电量汇总` aggregates by customer name. Output detail sheet `户号明细` retains account-number rows for audit and later ledger-update work.
- Missing customer name, missing account number, invalid period, non-numeric power, and negative power stop generation as serious source-data errors.
- Ledger update matches by `电力用户名称`, not `电力用户编码`.
- WPF preflight collects one-time customer handling decisions for unmatched power customers. Each displayed unmatched power customer must explicitly choose `新增客户到台账`, `不匹配，本月不写入`, or one existing ledger customer. Existing ledger customer targets can be used only once in the same preflight; create/skip actions can be repeated.
- Ledger update does not fill `电力用户编码` / B-column account numbers; account numbers remain in the cleaned detail workbook and report for traceability.
- Ledger update writes only target-month `总实际电量（兆瓦时）` and `尖/峰/平/谷` into a copied ledger. It does not overwrite the source ledger.
- If the target month block is absent, ledger update creates it by copying the previous Chongqing 30-column month block, changing the month label, clearing the target-month `总实际电量/尖/峰/平/谷` power columns, and preserving template columns such as coefficients, refund, agent/intermediary formulas, and remarks.
- `代理或自营=自营` customers can legitimately have blank agent/intermediary fixed fields and blank monthly agent/intermediary revenue columns; Stage 1 must not fill those fields or treat the blanks as errors.
- When the user chooses `新增客户到台账`, Chongqing Stage 1 inserts a row only in the output ledger copy, writes the customer name and target-month `总实际电量/尖/峰/平/谷`, preserves copied row formatting/formulas, and leaves `电力用户编码` / B-column account number plus负责人/代理/居间 fields blank for manual completion.
- Missing/new/ledger-only customers, possible alias candidates, multiple-account customers, month mismatch, and existing target-month power differences are surfaced in a WPF confirmation before writing and also written to the JSON report.

Known limits / next required behavior:

- Chongqing Stage 2 settlement generation is not implemented.
- Customer-name alias mapping is not automatic or persistent; possible aliases are reported and the WPF preflight lets the user choose one-time manual matches for the current run.
- The new Chongqing customer-resolution behavior is implemented with synthetic regression coverage and packaged for user acceptance, but has not yet been validated against the user's real Chongqing workbooks in this session.
- Current regressions use synthetic workbooks; no real Chongqing workbook is committed.

## UI State

Two desktop entries exist and should continue to coexist:

- `src/HainanSettlementTool.WinForms`: Win7/8 maintenance UI on `.NET Framework 4.7.2`.
- `src/HainanSettlementTool.Wpf`: Win10/11 primary UI shell on `.NET Framework 4.7.2`.

Both UI entries should remain thin shells for file selection, parameter input, confirmation, progress/log display, and error messages. Shared business behavior belongs in Core/Excel.
Core `SettlementWorkflow` lets both UI entries reuse the same stage completion summary rules while keeping UI-specific confirmation, progress, and error display local.
Do not add WinForms-only features or UX improvements by default. New UI work should target WPF unless the user explicitly asks for Win7/8 support.
Stage 2 workflow handling keeps WPF responsible for the confirmation dialog and progress UI, while Core `SettlementWorkflow` owns the preflight plan and the confirmed/cancelled generation decision.
The employee reward module currently exists only in the WPF app as a separate tab. It reuses the shared output folder and has its own start/end month selectors.
The WPF app supports UI theme selection in the custom title bar: `跟随系统`, `浅色`, and `深色`. The setting is stored in the existing WPF input snapshot XML. This is UI-only; generated Excel workbooks remain theme-independent, light, and print-safe.
The WPF app now has a top-level settlement-province selector. It defaults to empty on startup so the user must explicitly choose the province before running business actions. Before a province is selected, the main business area shows only a neutral empty state; stale file paths, run buttons, and province-specific output rows must stay hidden. `海南` keeps the mature Stage 1, Stage 2, and employee reward entries. `重庆` currently exposes only Stage 1 power cleaning; Hainan-only Stage 2 and employee reward UI entries are hidden when Chongqing is selected.
WPF business confirmation, warning, and error prompts should use project-native modern WPF dialogs, not system `MessageBox`; OS-native dialogs are still acceptable for file and folder selection.

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

Main merge verification on 2026-07-02:

```powershell
git diff --check
git status --short | Select-String -Pattern '\.(xlsx|xls|csv|png|jpg|jpeg|pdf)$'
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
```

Observed result:

- `main` fast-forwarded from `164d9e5` to `feb933f` with the employee reward module.
- `git diff --check` passed; Git only printed CRLF normalization warnings for the handoff update.
- No real Excel, CSV, image, PDF, or generated sensitive output file appears in `git status`.
- Core tests: 16 passed.
- Excel tests: 14 passed.
- Debug build passes for Core, Excel, WinForms, and WPF.

Chongqing Stage 1 power-cleaning branch verification on 2026-07-06:

```powershell
dotnet test .\tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj /p:Configuration=Debug --filter CleanProvinceStage1PowerDataReturnsSharedSummaryLines
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter ChongqingPowerCleanGeneratorTests
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
```

Observed result:

- New Core workflow test passed.
- New Excel tests passed, including Chongqing customer/account aggregation and negative-power stop behavior.
- WPF Debug build passed.
- Full Debug test suite passed: Core 17 tests, Excel 16 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.

Authorized Chongqing Stage 1 real sample smoke on 2026-07-06:

- Input source workbook inspected read-only: `C:\Users\juqx2\Desktop\2026年-重庆\重庆\重庆2026年电量确认结算单\2026年05月售电公司电量确认结算单.xlsx`
- Manual comparison workbook inspected read-only: `C:\Users\juqx2\Desktop\2026年-重庆\数据清洗—5月用户电量统计表.xlsx`
- Output folder: `C:\Users\juqx2\Desktop\2026年-重庆\test\codex-chongqing-smoke-20260706-142318`
- The generator produced `5月重庆零售侧用户电量数据处理表.xlsx` and `5月重庆零售侧用户电量校验报告.json`.
- JSON report result: source sheet `sheet1`, raw rows 212, customer rows 26, account rows 46, total power 12887.548 MWh, and one skipped non-power tail row.
- Generated customer summary matched the manual cleaned workbook on customer count, total power, and every row-level numeric power vector. One customer-name text difference remained between the raw transaction-center source and the manual cleaned workbook; the matching numeric vector confirms it is a name/alias difference, not a power calculation difference.

WPF small-window stash cleanup on 2026-07-06:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
```

Observed result:

- The historical `wpf-small-window-quality before chongqing work` stash was reviewed instead of blindly popped.
- Current branch already contained the action-row `DockPanel LastChildFill="False"` fixes.
- The remaining WPF launch-window fix was reapplied by setting `MinHeight` to 720 while keeping the default launch height at 900.
- WPF Debug build passed.
- Full Debug test suite passed: Core 17 tests, Excel 16 tests.
- The reviewed stash was dropped after its remaining change was reapplied.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260706-143742.zip`.

WPF province selector and dialog polish on 2026-07-06:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
.\scripts\package_wpf_release.ps1
```

Observed result:

- The settlement-province selector was moved to its own top row in the shared settings card and no longer restores a saved/default province on startup.
- Business actions now require an explicit province selection; no-selection state disables execution and shows a modern WPF error dialog.
- Remaining WPF `MessageBox` confirmations/errors in `MainWindow` were replaced by a project-native modern dialog. Stage 2 preflight already uses a custom WPF confirmation window; file/folder pickers still use OS-native dialogs.
- The development rule now states that WPF business confirmation, warning, and error prompts must use project-native WPF dialogs instead of system `MessageBox`.
- WPF Debug build passed.
- Debug build passed for Core, Excel, WinForms, and WPF.
- Full Debug test suite passed: Core 17 tests, Excel 16 tests.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260706-151425.zip`.

Authorized Chongqing ledger read-only structure analysis on 2026-07-06:

- Ledger inspected read-only: `C:\Users\juqx2\Desktop\2026年-重庆\重庆2026年售电结算台账20260609.xlsx`
- Structure: one `Sheet1`, 30 rows, 194 columns, three-level header, customer data rows 4-29, 26 customer rows.
- Fixed fields found in row 2 include `电力用户编码`, `电力用户名称`, `合同年用电量（兆瓦时）`, `履约开始月份`, `履约结束月份`, `项目开发人`, `代理或自营`, and `负责人`.
- The `电力用户编码` column is blank for all 26 rows. Because Chongqing customers can have multiple account numbers, the app no longer auto-fills this column and does not use it as a matching key.
- Current customer rows have no blank customer names, no blank负责人, no blank项目开发人, and no duplicate customer names. Business type count is 7 self-operated and 19 proxy rows.
- Month blocks were identified by row-1 merged headers. Existing blocks cover January through May 2026. The May block starts at `FI`: total actual power in `FI`, periods in `FJ:FM` (`尖/峰/平/谷`), coefficient columns in `FN:FO`, and downstream benefit fields after that.
- The current May ledger power values already match the generated Chongqing cleaned power output on customer count, total power, and every numeric power vector. The same one customer-name text difference remains, so the next module needs an alias/mismatch report rather than silent fuzzy matching.

Chongqing Stage 1 ledger update implementation verification on 2026-07-06:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
```

Observed result:

- Added Core models for Chongqing ledger update options, preflight plan, issues, and result.
- Added `ChongqingLedgerStage1Updater` in the Excel layer.
- WPF `清洗并更新台账` is now enabled for Chongqing and runs a preflight before writing.
- WPF preflight uses the modern dialog and shows matching issues before generating a copied ledger.
- Synthetic tests passed for workflow summary and Chongqing ledger update writing.
- Core tests: 18 passed.
- Excel tests: 17 passed.
- WPF Debug build passed.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260706-154127.zip`.

Authorized Chongqing Stage 1 ledger update real smoke on 2026-07-06:

- Inputs inspected read-only: `C:\Users\juqx2\Desktop\2026年-重庆\重庆2026年售电结算台账20260609.xlsx` and `C:\Users\juqx2\Desktop\2026年-重庆\重庆\重庆2026年电量确认结算单\2026年05月售电公司电量确认结算单.xlsx`
- Output folder: `C:\Users\juqx2\Desktop\2026年-重庆\test\codex-chongqing-ledger-update-smoke-20260706-153825`
- Generated output ledger: `5月重庆售电结算台账-阶段一更新.xlsx`
- Generated report: `5月重庆阶段一台账更新报告.json`
- Real-smoke preflight required confirmation: 25 matched customer rows, 7 multi-account rows, 1 power customer not in ledger, 1 ledger customer not in power table, 1 possible alias candidate, and 0 existing target-month power differences.
- The smoke wrote 25 matched power rows into the copied ledger and did not modify the original ledger or fill B-column account codes.

Chongqing Stage 1 B-column account-code policy update on 2026-07-06:

```powershell
dotnet test .\tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj /p:Configuration=Debug --filter UpdateProvinceStage1LedgerReturnsSharedSummaryLines
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter ChongqingPowerCleanGeneratorTests
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
git diff --check
git status --short | Select-String -Pattern '\.(xlsx|xls|csv|png|jpg|jpeg|pdf)$'
```

Observed result:

- User decided Chongqing B-column `电力用户编码` should not be automatically maintained because one customer may have multiple account numbers.
- `ChongqingLedgerStage1Updater` now writes only target-month power columns and never writes or compares B-column account codes.
- WPF preflight now shows multi-account customers as a reminder only, with wording that B column is not written.
- Synthetic Chongqing ledger update test now verifies B-column values stay blank while power columns are written.
- Full Debug test suite passed: Core 18 tests and Excel 17 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.
- `git diff --check` passed with CRLF normalization warnings only.
- No real Excel, CSV, image, PDF, or generated sensitive output file appears in `git status`.

Chongqing Stage 1 one-time manual customer matching on 2026-07-06:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter ChongqingPowerCleanGeneratorTests
dotnet test .\tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj /p:Configuration=Debug --filter UpdateProvinceStage1LedgerReturnsSharedSummaryLines
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\package_wpf_release.ps1
```

Observed result:

- Added `ProvinceStage1CustomerMatch` and manual-match lists on Chongqing Stage 1 ledger update options, plans, and results.
- Added a WPF Stage 1 preflight window that lets the user map unmatched cleaned power customers to ledger-only customers for the current run only.
- Manual matches are validated as one source customer to one target ledger customer; duplicate target selections are blocked before writing.
- `ChongqingLedgerStage1Updater` now writes exact matches plus confirmed manual matches into the copied ledger and records manual matches in the JSON update report.
- Synthetic alias test verifies a cleaned power customer with an old name can be written into a ledger customer row with a different current name.
- Full Debug test suite passed: Core 18 tests and Excel 18 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260706-171820.zip`.
- Package contents checked for the WPF executable, config, Core/Excel DLLs, and ClosedXML DLL.

WPF no-province empty state polish on 2026-07-06:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\package_wpf_release.ps1
```

Observed result:

- When the WPF settlement province is empty, the main business tab now shows a neutral empty state instead of stale file paths, province-specific labels, and run buttons.
- The right-side completion/output card now shows a no-province waiting state and hides province-specific output rows and completion time until a province is selected.
- Selecting a province resets the right-side previous-result summary so old success counts from another province are not shown under the new province.
- Full Debug test suite passed: Core 18 tests and Excel 18 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.
- `git diff --check` passed with CRLF normalization warnings only.
- WPF `MessageBox` search had no matches.
- No real Excel, CSV, image, PDF, or generated sensitive output file appears in `git status`.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260706-173120.zip`.
- Package contents checked for the WPF executable, config, Core/Excel DLLs, and ClosedXML DLL.

WPF Chongqing preflight manual matching polish on 2026-07-06:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\package_wpf_release.ps1
```

Observed result:

- The preflight manual-match dialog now uses clearer labels: unmatched source rows are shown as `待匹配客户名称`, and the right side is `选择台账客户名称（必选）`.
- The ledger-customer dropdown now renders `CustomerTargetOption.DisplayText` through an item template and `ToString()` fallback, so it no longer displays the WPF view-model type name.
- Manual matching rows no longer default to silent no-write. The user must explicitly choose a ledger customer or choose `不匹配，本月不写入`; otherwise confirmation is blocked with a validation message.
- `电量客户不在台账` and `台账客户不在电量表` are presented inside the customer manual-matching area, while `其它预检项目` only contains the remaining non-matching warnings.
- Full Debug test suite passed: Core 18 tests and Excel 18 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.
- `git diff --check` passed with CRLF normalization warnings only.
- WPF `MessageBox` search had no matches.
- No real Excel, CSV, image, PDF, or generated sensitive output file appears in `git status`.
- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260706-173952.zip`.
- Package contents checked for the WPF executable, config, Core/Excel DLLs, and ClosedXML DLL.

Multi-province readiness architecture self-audit on 2026-07-07:

Observed result:

- Added `docs/dev-notes/multi-province-readiness-2026-07-07.md`.
- Identified the main third-province readiness debt: WPF province UI/profile logic, Excel province-stage adapter dispatch, Core province-stage validation, stable preflight issue codes, and province onboarding contract tests.
- Recorded P0/P1/P2 readiness order so future new-province work can first reduce the highest-risk coupling instead of adding more scattered province branches.
- Updated `AGENTS.md` and this handoff so future sessions read the readiness note before new-province onboarding, WPF province UI, Core multi-province workflow, or Excel multi-province adapter work.
- Recorded the project collaboration rule that Codex subagents/spawning/parallel exploration are allowed for safe efficiency, with user warning/confirmation reserved for high-risk operations.
- No build or test run was required; this was a documentation-only architecture audit.

Multi-province P0 readiness slice on 2026-07-07:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter ChongqingPowerCleanGeneratorTests
dotnet test .\tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj /p:Configuration=Debug --filter UpdateProvinceStage1LedgerReturnsSharedSummaryLines
```

Observed result:

- Added WPF `ProvinceUiProfile` to centralize province display name, stage availability, Stage 1 labels, button text, result labels, and file-picker titles for Hainan and Chongqing.
- `MainWindow` now binds the province dropdown to profile objects and uses profile capabilities for tab visibility, input-row visibility, result-row visibility, and Stage 1 button enablement.
- Stage 1 run/clean actions now explicitly dispatch by `ProvinceCode`, so a future newly listed province will not silently run the Hainan workflow.
- Added stable `ProvinceStage1LedgerUpdateIssue.Kind` plus `ProvinceStage1LedgerUpdateIssueKinds` while keeping Chinese `Category` for WPF display and JSON compatibility.
- WPF Stage 1 preflight manual-matching grouping now uses stable issue kinds, with Chinese category fallback for old or hand-built issue objects.
- Added Excel internal `IProvinceStage1Adapter` and `ChongqingProvinceStage1Adapter`; `ClosedXmlSettlementExcelGateway` now dispatches multi-province Stage 1 Excel work through an adapter dictionary instead of direct Chongqing if branches.
- Added a repository `.vscode/settings.json` pointing VSCode at `HainanSettlementTool.sln` to reduce false WPF code-behind diagnostics.
- Targeted WPF Debug build passed.
- Targeted Chongqing Excel tests passed: 4 tests.
- Targeted Core workflow test passed: 1 test.
- Full Debug test suite passed: Core 18 tests, Excel 18 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.
- `git diff --check` passed with CRLF normalization warnings only.
- WPF `MessageBox` search had no matches.
- No real Excel, CSV, image, PDF, or generated sensitive output file appears in `git status`.

WPF Release test package on 2026-07-07:

```powershell
.\scripts\package_wpf_release.ps1
```

Observed result:

- Win10/11 WPF test package generated at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-100359.zip`.
- Package contents checked for the WPF executable, config, Core/Excel DLLs, ClosedXML DLLs, and runtime dependency DLLs.
- This is a development test package from `codex/wpf-mainwindow-decomposition`, not a formal release tag.

WPF MainWindow progress decomposition on 2026-07-07:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
git diff --check
rg "MessageBox" -n src\HainanSettlementTool.Wpf
```

Observed result:

- Added `MainWindowProgressController` to own the WPF status pill, progress bar, progress description, and five-step status rendering.
- `MainWindow.xaml.cs` keeps the existing workflow call sites through thin helper methods, but no longer directly owns step text/status arrays or step rendering rules.
- This is an internal UI-architecture slice only; no settlement calculation, Excel read/write, customer matching, or generated workbook behavior changed.
- Targeted WPF Debug build passed.
- Full Debug test suite passed: Core 18 tests, Excel 18 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.
- `git diff --check` passed with CRLF normalization warnings only.
- WPF `MessageBox` search had no matches.

WPF MainWindow result decomposition on 2026-07-07:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug
rg -n "_lastOutputDirectory|Stage1ResultStatus\.Text|ProxyResultStatus\.Text|IntermediaryResultStatus\.Text|SummaryResultStatus\.Text|EmployeeRewardResultStatus\.Text|FinishedAtText\.Text|CompletionTitleText\.Text|CompletionOutputText\.Text|CompletionCard\.Visibility" .\src\HainanSettlementTool.Wpf\MainWindow.xaml.cs
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
git diff --check
rg "MessageBox" -n src\HainanSettlementTool.Wpf
```

Observed result:

- Added `MainWindowResultController` to own the WPF completion card, output item visibility, result status/count rows, finished-at timestamp, and most recent output directory.
- `MainWindow.xaml.cs` now reports stage success through result helper methods instead of writing result TextBlocks directly, while workflow orchestration remains in the window.
- This is an internal UI-architecture slice only; no settlement calculation, Excel read/write, customer matching, confirmation flow, or generated workbook behavior changed.
- Targeted WPF Debug build passed.
- Result-control residual scan had no matches for old direct result TextBlock writes or `_lastOutputDirectory`.
- Full Debug test suite passed: Core 18 tests, Excel 18 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.
- `git diff --check` passed with CRLF normalization warnings only.
- WPF `MessageBox` search had no matches.

Chongqing Stage 1 target-month block creation fix on 2026-07-07:

```powershell
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter UpdateLedgerCreatesTargetMonthBlockFromPreviousChongqingLedger
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter ChongqingPowerCleanGeneratorTests
dotnet test .\tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj /p:Configuration=Debug --filter ProvinceStage1
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
git diff --check
rg "MessageBox" -n src\HainanSettlementTool.Wpf
```

Observed result:

- Diagnosed the bug with three user-authorized real Chongqing ledgers inspected read-only: the 4-month ledger ended at the 4月 block, while the manually prepared 5-month ledgers had an added 30-column 5月 block.
- Added Chongqing ledger update behavior that creates a missing target month block from the previous month block, clears only the target month `总实际电量/尖/峰/平/谷` columns, and preserves the remaining monthly template columns/formulas.
- Added synthetic regression coverage for updating a 5月 ledger from a 4月-only Chongqing ledger, including a `代理或自营=自营` customer whose agent/intermediary revenue columns stay blank.
- The new regression failed before the fix with `重庆台账中未找到5月电量区块` and passed after the fix.
- Targeted Chongqing Excel tests passed: 5 tests.
- Targeted Core province-stage tests passed: 2 tests.
- Full Debug test suite passed: Core 18 tests, Excel 19 tests.
- Debug build passed for Core, Excel, WinForms, and WPF.
- `git diff --check` passed with CRLF normalization warnings only.
- WPF `MessageBox` search had no matches.
- The broader project/namespace name `HainanSettlementTool` remains unchanged for now. Province-neutral naming cleanup has been promoted into today's multi-province technical-debt mainline, but it should still be done through low-risk internal slices instead of a large all-at-once project/namespace rename.

WPF local Release test package on 2026-07-07:

```powershell
.\scripts\package_wpf_release.ps1
```

Observed result:

- Release build passed for Core, Excel, WinForms, WPF, and both test projects as part of the packaging script.
- Package directory created at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-105743`.
- Zip package created at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-105743.zip`.
- Package content was checked and includes the Win10/11 exe, `.exe.config`, required `.dll` files, and `README.txt`.
- `git status --short --branch` was clean after packaging; generated `dist/` artifacts are not tracked.

Province-neutral naming slice on 2026-07-07:

```powershell
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet test .\tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj /p:Configuration=Debug --filter SettlementWorkflowTests
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter "ChongqingPowerCleanGeneratorTests|Stage1LedgerUpdaterTests|Stage2SettlementGeneratorTests|EmployeeRewardGeneratorTests|RawDetailReaderTests"
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
git diff --check
rg "MessageBox" -n src\HainanSettlementTool.Wpf
```

Observed result:

- Renamed the Hainan-specific Stage 1/Stage 2 service and gateway interfaces to `HainanStage1Service`, `HainanStage2Service`, `IHainanStage1ExcelGateway`, and `IHainanStage2ExcelGateway`.
- Renamed the combined ClosedXML gateway to `ClosedXmlSettlementExcelGateway` because it now coordinates Hainan Stage 1, Hainan Stage 2, employee reward, and multi-province Stage 1 adapters.
- Added `ProvinceDisplayNames` so Core/Excel/WPF no longer call `ProvinceStage1Service` just to format province names.
- Updated the WPF shell title and saved-log default name to multi-province wording.
- `UserInputStore` now writes to the neutral AppData folder `SettlementAutomationTool` and still falls back to the old `HainanSettlementTool` folder on load.
- Debug build passed for Core, Excel, WinForms, WPF, and both test projects.
- Targeted Core workflow tests passed: 9 tests.
- Targeted Excel coverage for renamed gateway/service callers passed: 19 tests.
- Full Debug test suite passed: Core 18 tests, Excel 19 tests.
- `git diff --check` passed with CRLF normalization warnings only.
- WPF `MessageBox` search had no matches.

WPF dialog-controller slice on 2026-07-07:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m
rg -n "new ModernDialogWindow|new ConfirmRunWindow|MessageBox" src\HainanSettlementTool.Wpf\MainWindow.xaml.cs src\HainanSettlementTool.Wpf
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
git diff --check
```

Observed result:

- Added `MainWindowDialogController` to own `ModernDialogWindow` and `ConfirmRunWindow` creation for common main-window error, warning, and run-confirmation prompts.
- `MainWindow.xaml.cs` keeps thin helper methods, so existing workflow call sites and prompt behavior stay unchanged.
- Targeted WPF Debug build passed.
- Dialog residual scan shows window construction only inside `MainWindowDialogController`; no WPF `MessageBox` usage was introduced.
- Debug build passed for Core, Excel, WinForms, WPF, and both test projects.
- `git diff --check` passed with CRLF normalization warnings only.

WPF path-picker-controller slice on 2026-07-07:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m
rg -n "new OpenFileDialog|new VistaFolderBrowserDialog|Ookii\.Dialogs\.Wpf|CheckFileExists|UseDescriptionForTitle|ShowNewFolderButton" src\HainanSettlementTool.Wpf\MainWindow.xaml.cs src\HainanSettlementTool.Wpf
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
git diff --check
```

Observed result:

- Added `MainWindowPathPickerController` to own `OpenFileDialog` and `VistaFolderBrowserDialog` creation.
- `MainWindow.xaml.cs` now keeps thin browse helpers that call the controller and save inputs after a successful selection.
- Targeted WPF Debug build passed.
- Path-picker residual scan shows file/folder dialog construction only inside `MainWindowPathPickerController`.
- Debug build passed for Core, Excel, WinForms, WPF, and both test projects.
- `git diff --check` passed with CRLF normalization warnings only.

Chongqing Stage 1 customer-resolution and WPF display fixes on 2026-07-07:

```powershell
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter UpdateLedgerAppliesCreateAndSkipCustomerDecisions
dotnet test .\tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj /p:Configuration=Debug --filter ChongqingPowerCleanGeneratorTests
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
.\scripts\package_wpf_release.ps1
```

Observed result:

- Added Core customer-decision model for multi-province Stage 1: `MatchExisting`, `CreateNew`, and `SkipWrite`.
- Chongqing Stage 1 WPF preflight now requires every unmatched power customer to choose `新增客户到台账`, `不匹配，本月不写入`, or one existing ledger customer. Existing ledger targets are unique per preflight; create/skip actions can repeat.
- Chongqing ledger update now applies create/skip decisions. `CreateNew` inserts a row only in the output copy, writes customer name and target-month power, preserves copied formatting/formulas, and leaves B-column/customer-code plus负责人/代理/居间 fields blank for manual completion. `SkipWrite` leaves the customer out of the ledger update and records the decision in JSON.
- Hainan Stage 1 behavior was not changed.
- Fixed WPF display fallbacks so the province selector and preflight customer-target dropdown show reader-facing text instead of internal type names.
- WPF visible window title/header changed to `清能电力-结算自动化工具`; project/namespace/assembly/package names were not renamed.
- Added synthetic regression coverage for a Chongqing run with one created customer and one skipped customer. The targeted new test passed.
- Targeted Chongqing Excel tests passed: 6 tests.
- Targeted WPF Debug build passed.
- Full Debug test suite passed: Core 18 tests, Excel 20 tests.
- Debug build passed for Core, Excel, WinForms, WPF, and both test projects.
- Release packaging script passed and created `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-115127.zip`.
- Package content was checked and includes the Win10/11 exe, `.exe.config`, required `.dll` files, and `README.txt`.
- No real Excel, real ledger, customer data, screenshots, or settlement outputs were read for this implementation/validation.

Main merge and acceptance packages on 2026-07-07:

```powershell
git merge --no-ff codex/wpf-path-picker-controller -m "Merge Chongqing stage one WPF updates"
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m
git diff --check
rg "MessageBox" -n src\HainanSettlementTool.Wpf
.\scripts\check_build_portability.ps1
.\scripts\package_wpf_release.ps1
.\scripts\package_release.ps1
```

Observed result:

- Local `main` merged `codex/wpf-path-picker-controller` with merge commit `4911e47 Merge Chongqing stage one WPF updates`.
- Full Debug test suite passed on `main`: Core 18 tests, Excel 20 tests.
- A first concurrent Debug build attempt hit a transient file lock because tests were still using `HainanSettlementTool.Core.dll`; a sequential rerun passed for Core, Excel, WinForms, WPF, and both test projects.
- Release build passed for Core, Excel, WinForms, WPF, and both test projects.
- `git diff --check` passed.
- WPF `MessageBox` search had no matches.
- Build portability check passed.
- Win10/11 WPF acceptance package created at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-140505.zip`.
- Win7/8 maintenance package created at `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win7-8-Release-20260707-140515.zip`.
- Both package directories were checked for the executable, `.exe.config`, required `.dll` files, and `README.txt`.
- No real Excel, real ledger, customer data, screenshots, or settlement outputs were read for this merge/package validation.
- No formal tag or GitHub Release was created.
- Remote push was not performed.

WPF log-controller slice on 2026-07-07:

```powershell
dotnet msbuild .\src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Debug /m
git diff --check
rg "MessageBox" -n src\HainanSettlementTool.Wpf
```

Observed result:

- Added `MainWindowLogController` to own WPF run-log append, scroll-to-end, clear, and save-to-text-file behavior.
- `MainWindow.xaml.cs` keeps thin `AddLog`, `ClearLog_Click`, and `SaveLog_Click` wrappers so existing workflow call sites stay unchanged.
- This is an internal UI-architecture slice only; no settlement calculation, Excel read/write, customer matching, confirmation flow, generated workbook behavior, or user-visible log text changed.
- Targeted WPF Debug build passed.
- Full Debug test suite passed: Core 18 tests, Excel 20 tests.
- Debug build passed for Core, Excel, WinForms, WPF, and both test projects.
- `git diff --check` passed with CRLF normalization warnings only.
- WPF `MessageBox` search had no matches.

## Documentation Rule

Documentation is now part of the development contract:

- At the start of each project work session, run `git status --short --branch` and read `AGENTS.md` plus this `HANDOFF.md` before editing.
- Read the owning document before changing a responsibility area: `CONTEXT.md` for settlement rules, `docs/architecture.md` for module seams or workflow structure, `README.md` for user-facing setup/package status, and `docs/RELEASE_CHECKLIST.md` for release or packaging.
- Repeat the relevant reading gate after context compaction, a long pause, or a task direction change.
- This is a local single-developer project. Pull requests are optional; after a feature branch is committed and pushed, local merge to `main` is acceptable when the user authorizes it.
- Codex subagents, spawning, parallel exploration, and other efficiency tools are allowed for this project when they materially speed up safe work. Routine low-risk use does not require repeated confirmation; pause and warn or ask only before high-risk operations such as reading real business files outside an authorized scope, modifying production/user workbooks, destructive git commands, merging to `main`, tagging/releasing, deleting/moving large file trees, or actions that could affect settlement correctness or sensitive data.
- Every code, config, script, packaging, business-rule, UI-behavior, test-workflow, or task-state change must end with a documentation impact judgment.
- Update only documents whose responsibility is affected; do not rewrite unaffected docs for process compliance.
- Business-rule changes must check `CONTEXT.md`; module-boundary changes must check `AGENTS.md` and an ADR or dev note; release/packaging changes must check `README.md` and `docs/RELEASE_CHECKLIST.md`; branch state, validation results, or next steps must check this `HANDOFF.md`.
- Before finishing a change, record whether docs were affected. If yes, list the updated docs; if no, state why no doc update was needed.
- Keep this `HANDOFF.md` current whenever branch state, release status, validation results, or next steps change.
- If a pull request is used, use `.github/PULL_REQUEST_TEMPLATE.md` as the merge-time documentation checklist. If merging locally without a PR, apply the same documentation-impact and validation checks manually.
- Put temporary investigations and architecture decisions in dated files under `docs/dev-notes/`.

## Useful Files

Core:

- `src/HainanSettlementTool.Core/Services/HainanStage1Service.cs`
- `src/HainanSettlementTool.Core/Services/HainanStage2Service.cs`
- `src/HainanSettlementTool.Core/Services/IHainanStage1ExcelGateway.cs`
- `src/HainanSettlementTool.Core/Services/IHainanStage2ExcelGateway.cs`
- `src/HainanSettlementTool.Core/Services/SettlementWorkflow.cs`
- `src/HainanSettlementTool.Core/Services/Stage2WorkflowPlan.cs`
- `src/HainanSettlementTool.Core/Services/Stage2WorkflowResult.cs`
- `src/HainanSettlementTool.Core/Services/Stage2SettlementCalculator.cs`
- `src/HainanSettlementTool.Core/Services/EmployeeRewardService.cs`
- `src/HainanSettlementTool.Core/Services/IEmployeeRewardExcelGateway.cs`
- `src/HainanSettlementTool.Core/Services/FileAccessGuard.cs`
- `src/HainanSettlementTool.Core/Models/ProvinceDisplayNames.cs`
- `src/HainanSettlementTool.Core/Models/ProvinceStage1LedgerUpdateIssueKinds.cs`

Excel:

- `src/HainanSettlementTool.Excel/RawDetailRowReader.cs`
- `src/HainanSettlementTool.Excel/RawDetailReader.cs`
- `src/HainanSettlementTool.Excel/CustomerCodeReader.cs`
- `src/HainanSettlementTool.Excel/LedgerStage1Updater.cs`
- `src/HainanSettlementTool.Excel/Stage2SettlementGenerator.cs`
- `src/HainanSettlementTool.Excel/EmployeeRewardGenerator.cs`
- `src/HainanSettlementTool.Excel/ClosedXmlSettlementExcelGateway.cs`
- `src/HainanSettlementTool.Excel/IProvinceStage1Adapter.cs`
- `src/HainanSettlementTool.Excel/ChongqingProvinceStage1Adapter.cs`

UI:

- `src/HainanSettlementTool.WinForms/MainForm.cs`
- `src/HainanSettlementTool.Wpf/MainWindow.xaml`
- `src/HainanSettlementTool.Wpf/MainWindow.xaml.cs`
- `src/HainanSettlementTool.Wpf/MainWindowDialogController.cs`
- `src/HainanSettlementTool.Wpf/MainWindowPathPickerController.cs`
- `src/HainanSettlementTool.Wpf/MainWindowProgressController.cs`
- `src/HainanSettlementTool.Wpf/MainWindowResultController.cs`
- `src/HainanSettlementTool.Wpf/MainWindowLogController.cs`
- `src/HainanSettlementTool.Wpf/ProvinceUiProfile.cs`
- `src/HainanSettlementTool.Wpf/Stage2PreflightWindow.xaml`
- `src/HainanSettlementTool.Wpf/ProvinceStage1LedgerPreflightWindow.xaml`

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

1. Have the user test the local `main` acceptance package `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-140505.zip` against a Chongqing work copy.
2. If the user accepts this package, decide whether to push local `main` to `origin/main`, cut a formal release, or keep it as a local acceptance build.
3. If the user authorizes a real-data smoke, run it read-only against specifically authorized Chongqing input files and write outputs only to an explicitly selected test/output folder.
4. Review and, if accepted, merge `codex/wpf-log-controller` back to `main`; it is a low-risk internal WPF slice and does not require a new acceptance package unless the user wants one.
5. Decide later whether repeated manual matches should remain one-time only or support a user-maintained alias table.
6. Consider adding sanitized employee reward, Stage 2, Chongqing, and `.xls` fixture workbooks later; current regressions use dynamically generated synthetic workbooks and local authorized smoke only.
