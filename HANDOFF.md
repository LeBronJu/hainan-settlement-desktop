# Handoff

Last updated: 2026-07-10

## Project

Standalone C# desktop rewrite of the Hainan retail electricity settlement automation tool.

- Local path: `D:\Document\文件处理\hainan-settlement-desktop`
- GitHub: `https://github.com/LeBronJu/hainan-settlement-desktop`
- Original Python baseline: `D:\Document\文件处理\hainan-settlement-tool`
- Original Python GitHub: `https://github.com/LeBronJu/hainan-settlement-tool`
- Stable local reference folder: `D:\Document\文件处理\稳定参考版海南结算`

The Python project remains the historical full-function reference. The C# desktop project is versioned and released independently.

## Start Here

Every new work session must start with:

```powershell
git status --short --branch
```

Then read:

1. `AGENTS.md`
2. `HANDOFF.md`
3. `docs/README.md`
4. The owning document for the task area, as listed in `docs/README.md`

Do not treat this file as a business specification. It is only the current handoff: branch state, safety boundaries, latest validation, packages, and next steps.

## Data Safety

- Never commit real ledgers, settlement outputs, customer data, screenshots, payment data, or generated files from real workbooks.
- Never overwrite user workbooks in place. Generated outputs must go to an output folder or a clearly named copy.
- Real production environment root: `C:\Users\juqx2\Desktop\2026海南`.
- Chongqing local working root: `C:\Users\juqx2\Desktop\2026年-重庆`.
- Stable local reference folder: `D:\Document\文件处理\稳定参考版海南结算`.
- Files under these roots are real business data. Do not read them unless the user explicitly authorizes the file, folder, or read-only analysis scope for the current task.
- Prior read-only authorizations do not grant write permission.
- Known user test/output area: `C:\Users\juqx2\Desktop\2026海南\test`. Files there may still contain real data and must not be committed.
- Repository-safe real smoke script: `scripts/run_real_smoke.ps1`. It is currently documented as a historical tool note because service/gateway names changed; refresh it before treating it as runnable.

Historical real-data authorizations recorded for context only. They are not standing permission for new reads:

- 2026-07-02: read-only inspection of employee reward reference and then-current Hainan production ledger for employee reward design.
- 2026-07-06: read-only inspection of selected Chongqing Stage 1 source/comparison/ledger workbooks for Stage 1 design.
- 2026-07-07: read-only inspection of `C:\Users\juqx2\Desktop\2026年-重庆` for Chongqing Stage 2 analysis.
- 2026-07-07: read-only inspection of `C:\Users\juqx2\Desktop\2026海南` for Hainan Stage 2 behavior comparison.
- 2026-07-10: read-only inspection and temporary-output smoke of `C:\Users\juqx2\Desktop\广东\交易结算\2026广东代理` for Guangdong month-sheet preparation. Temporary real-data copies were deleted after verification; source workbooks were not modified.

## Current Git State

- Current branch: `codex/guangdong-stage2-month-preparation`
- Upstream: to be pushed after focused validation.
- Branch purpose: Guangdong low-risk proxy/intermediary/refund workbook month-sheet preparation plus third-province WPF capability cleanup.
- The branch is based on local `main` after the completed Chongqing Stage 2 branch was fast-forwarded. Chongqing closeout commit `6be47aa` is already pushed on `origin/codex/chongqing-stage2-analysis`; pushing updated `main` is pending a stable GitHub connection.
- Do not merge to `main`, tag, or publish a release without explicit user authorization.

- No formal release tag has been cut after `v1.0.1`.
- Completed milestone history is summarized in `docs/CHANGELOG.md`.

## Current Packages

Latest local WPF test package with Guangdong month-sheet preparation:

- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260710-123600.zip`
- Unpacked directory: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260710-123600`

Latest local WPF test package built after the Chongqing Stage 2 practical UI/summary-format fixes:

- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260709-150129.zip`
- Unpacked directory: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260709-150129`

Previous package after the Chongqing Stage 2 backtest fixes:

- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260709-113859.zip`

Previous Chongqing 3-5月 combined-backtest package before the latest fixes:

- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260709-104346.zip`

Previous Chongqing Stage 2 first workbook-generation package:

- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260708-101959.zip`

Latest local acceptance packages before the log-controller merge:

- Win10/11 WPF: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-140505.zip`
- Win7/8 WinForms maintenance: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win7-8-Release-20260707-140515.zip`

User practical testing on 2026-07-07 reported no blocking issues so far with:

- `HainanSettlementTool-Win10-11-Release-20260707-140505.zip`

User practical testing on 2026-07-09 reported no blocking issues so far with:

- `HainanSettlementTool-Win10-11-Release-20260708-155953.zip`

Current formal release:

- Tag: `v1.0.1`
- Release page: `https://github.com/LeBronJu/hainan-settlement-desktop/releases/tag/v1.0.1`
- Assets:
  - `HainanSettlementTool-Win7-8-v1.0.1.zip`
  - `HainanSettlementTool-Win10-11-v1.0.1.zip`

## Functional State

Hainan:

- Hainan Stage 1 and Stage 2 are the mature production path.
- Hainan Stage 2 current behavior is now documented in `docs/hainan-stage2-current-behavior.md`.
- Hainan Stage 2 now requires an explicit `清能`/`清辉` preflight choice for new summary subjects whose payment party cannot be inherited. Existing summary-subject inheritance remains unchanged.
- Do not infer Chongqing behavior from Hainan defaults unless the Chongqing note explicitly says so.

Employee reward:

- Merged to `main`.
- Exposed in Win10/11 WPF.
- User tested the practical flow on 2026-07-02 and reported no blocking issues.
- Current Core/Excel/WPF workflow implementation is explicitly named `HainanEmployeePowerReward...`; future Chongqing power reward work should not reuse it without a Chongqing-specific rule review.

Chongqing Stage 1:

- Implemented in Win10/11 WPF.
- Supports power cleaning and ledger update.
- Unmatched power customers must be explicitly handled in preflight: create, skip, or match one existing ledger customer.
- Create/skip decisions can repeat; a concrete existing ledger customer target can be used only once per preflight.
- Newly created Chongqing ledger rows write only customer name and target-month power in the output copy. They do not auto-fill B-column account number, `电力用户编码`,负责人,代理/居间 fields, or other manual fields.

Chongqing Stage 2:

- First workbook-generation slice is implemented for Core/Excel and Win10/11 WPF.
- Analysis is documented in `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`.
- First Core contract slice is implemented: `ChongqingStage2Options`, preflight/report/issue models, `ChongqingStage2Service`, `IChongqingStage2ExcelGateway`, and workflow plan/complete/run entry points.
- Excel slice reads the Chongqing ledger target-month 30-column block, identifies proxy/intermediary/refund groups, requires payment-party selection for new summary subjects, writes proxy/intermediary/refund split workbooks, writes the summary workbook copy, and writes JSON/validation reports.
- WPF entry now runs preflight/payment-party confirmation and then calls `Generate`.
- Real 5月 local smoke succeeded on 2026-07-08 against authorized Chongqing inputs, outputting only to `%TEMP%`; it reported proxy 19 rows/2 groups, refund 4 rows/3 groups, intermediary 0 rows/0 groups, with 0 warnings and 0 audit issues.
- User实机测试 on 2026-07-09 reported重庆阶段二入口能完整跑完；代理/退补/居间分表基本符合人工结果，个别特殊客户仍需要手动小改；汇总表副本、清能/清辉月度 sheet、JSON 报告和校验报告正常；暂未发现金额、公式、隐藏列、少回收电能量电费或支付方选择的静默错误。
- 2026-07-10 用户确认重庆数据验证基本没有问题。重庆阶段二当前作为首个可用生产基线进入长期观察；个别客户临时人工调整和已记录退补复核项不自动硬编码为通用规则。
- 2026-07-09 authorized 3/4/5月 combined backtest completed for Stage 1 + Stage 2. Stage 1 target-month power columns matched the authorized latest ledger historical month blocks with 0 power diffs and 0 missing rows. Stage 2 fixed target-month sheet preservation, payment-party sheet hidden-column preservation, and refund extra-block C-G power sync.
- Chongqing Stage 2 now pops a concise WPF completion warning when generated reports contain warning/audit items. It shows counts/categories and the validation report path rather than long per-file warning lines. Current refund extra-block warnings explicitly say C-G was synced and H onward plus summary monthly deduction/actual payment still require manual review.
- Chongqing summary output now moves target-month `清能X月` / `清辉X月` sheets directly after `汇总表` and merges their first-row title through the current `备注` column.
- WPF startup no longer restores the last selected province; saved paths still restore, and province/month must be selected explicitly. Chongqing shows the main settlement tab as `代理费结算` and displays a disabled `员工电量奖励` tab with `正在开发中`.
- Final backtest still returns `ReviewRequired` for expected manual/history differences: date display text, formula text style, refund C-column formula style, and 3月 historical manual deduction/actual-payment values. Do not hardcode these without user confirmation.
- Keep this work in Win10/11 WPF plus Core/Excel shared layers; do not add WinForms parity unless the user explicitly reopens Win7/8 support.

Guangdong:

- Current first slice is only proxy/intermediary/refund workbook month-sheet preparation; it does not read a ledger, calculate settlement amounts, or generate a summary workbook.
- Standard month sheet names are exact numeric text. Non-standard sheets are preserved but ignored for source selection.
- Existing target-month sheets are preserved and normalized; missing target months are copied only from the exact previous numeric month sheet.
- Detail C-F power contents are cleared, total formulas remain, deterministic period/date updates are applied, and uncertain structures are skipped with a report.
- WPF shows Guangdong as a third province with only the three optional split-workbook directories and a month-preparation action. Switching province clears the previous province's input paths.
- Current module rules and validation live in `docs/dev-notes/guangdong-stage2-month-preparation-2026-07-10.md`.

WinForms:

- Win7/8 WinForms is frozen as a historical compatibility entry.
- Do not spend feature, UX, parity, or package work on WinForms unless the user explicitly reopens Win7/8 support.

## Chongqing Stage 2 Reminders

Before implementing, reread:

- `CONTEXT.md`
- `docs/architecture.md`
- `docs/hainan-stage2-current-behavior.md`
- `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`
- `docs/dev-notes/multi-province-readiness-2026-07-07.md`

Key implementation constraints:

- Chongqing ledger month blocks are 30 columns and include hidden columns; read/write by headers/month-block offsets and preserve hidden state.
- Chongqing uses `兆瓦时`; do not reuse Hainan `万千瓦时` assumptions.
- Chongqing summary long-term anchor is `当年费用总计`, not Hainan `累计代理费总计`.
- Chongqing Stage 2 must handle代理费、居间费、退补电费. Current observed data has no effective居间 rows, but the user confirmed future居间 should follow代理 structure.
- Proxy split sheets must carry `少回收电能量电费`; adjusted income is original income minus that value.
- Refund split sheets use segmented尖峰/峰/平/谷 prices and cannot reuse proxy/intermediary single-price formulas.
- Authorized Chongqing proxy/退补 template folders were rechecked on 2026-07-07; current real `.xlsx` templates are standard xlsx ZIP containers. Ignore `._*` / `~$*` noise. Template reads/copies should use shared read because users may have source workbooks open.
- Existing summary rows should inherit template fixed fields and long-term fields.
- New Chongqing summary subjects should not silently default to Hainan `清辉`. Recommended first implementation: WPF preflight requires payment-party selection (`清能`/`清辉`) for new summary subjects, and generated reports flag defaulted fields for manual review.
- 重庆 3-5 月阶段一+阶段二真实回测已完成；`scripts/run_chongqing_backtest.ps1` 保留为后续授权月份回测入口。脚本必须通过参数或 case JSON 接收路径，并只写 `dist/` 或其它明确测试输出目录。

## Documentation Pointers

- Use `docs/README.md` as the canonical document map and task reading router.
- Use `docs/CHANGELOG.md` for completed high-signal milestone history.
- Use `scripts/check_docs_guardrails.ps1` before finishing documentation-affecting changes.
- Keep this file concise. It should not accumulate long historical build logs, detailed investigations, or duplicate code indexes.

## Latest Validation

Most recent non-documentation validation already recorded:

- WPF log-controller package script passed on 2026-07-07 and created `HainanSettlementTool-Win10-11-Release-20260707-150019.zip`.
- Full Debug and Release builds passed during that WPF log-controller validation.
- User tested `HainanSettlementTool-Win10-11-Release-20260707-140505.zip` and reported no blocking issues so far.

Most recent documentation validation:

- 2026-07-07 documentation cleanup created `docs/README.md` and `docs/hainan-stage2-current-behavior.md`.
- Dev notes were marked as current task/current policy/current module/current process/historical where appropriate.
- Stale wording scan and `git diff --check` passed for that cleanup before commit `48c4921`.
- Current documentation quality pass compressed `HANDOFF.md`, added `docs/CHANGELOG.md`, unified reading gates, expanded task-based document routing, and added `scripts/check_docs_guardrails.ps1`. Validation scope is the guardrail script, `git diff --check`, and stale wording/status scans.
- Pre-Chongqing Stage 2 closeout on 2026-07-07: `scripts/check_docs_guardrails.ps1` passed, `git diff --check` passed, stale wording scan had no matches, Debug tests passed (Core 18, Excel 20), and sequential Debug solution build passed. An earlier concurrent Debug build attempt hit a transient file lock while tests were still running; the sequential rerun passed.
- Hainan Stage 2 new-summary payment-party guard on 2026-07-07: Core tests passed (26), Excel tests passed (23), Debug solution build passed, and Release solution build passed.
- Chongqing Stage 2 Core contract and Excel Analyze-only slices on 2026-07-07: Core tests passed (26), Excel tests passed (27), and Debug solution build passed.
- WPF package script passed on 2026-07-07 after those slices and created `HainanSettlementTool-Win10-11-Release-20260707-173137.zip`.
- Chongqing Stage 2 WPF preflight entry / Win7-8 freeze update on 2026-07-07: WPF Debug build passed, Debug tests passed (Core 26, Excel 27), WPF Release build passed, docs guardrails passed, `git diff --check` passed, and WPF package script created `HainanSettlementTool-Win10-11-Release-20260707-175657.zip`. Real Chongqing 5月 Analyze-only preflight was run read-only against latest summary and 20260512 historical summary; both returned 0 payment-party issues.
- Chongqing Stage 2 first workbook-generation slice on 2026-07-08: focused Excel tests passed, WPF Debug build passed, real 5月 local smoke generated outputs under `%TEMP%` with 0 warnings/audit issues, Core tests passed (26), Excel tests passed (27), WPF Release build passed, and WPF package script created `HainanSettlementTool-Win10-11-Release-20260708-101959.zip`.
- Multi-province code-quality first slice on 2026-07-08: Hainan Stage 2 Excel generator was renamed/split into `HainanStage2...` components with external Analyze/Generate behavior unchanged; Core tests passed (26), Excel tests passed (27), and WPF Debug/Release builds passed.
- Multi-province code-quality second slice on 2026-07-08: WPF `MainWindowInputController` now owns saved input paths, saved-province restore, options construction, month/province selection reads, and clear-input actions; Core tests passed (26), Excel tests passed (27), and WPF Debug/Release builds passed.
- Multi-province code-quality third slice on 2026-07-08: WPF `MainWindowProvinceUiController` now owns settlement-month enablement, province tab/panel visibility, province UI copy, and province button enablement; Core tests passed (26), Excel tests passed (27), and WPF Debug/Release builds passed.
- Multi-province code-quality fourth slice on 2026-07-08: WPF `MainWindowStage2WorkflowController` now owns Hainan/Chongqing Stage 2 plan-confirm-complete orchestration and `SettlementWorkflowFactory` owns workflow construction; Core tests passed (26), Excel tests passed (27), and WPF Debug/Release builds passed.
- Multi-province code-quality fifth slice on 2026-07-08: WPF `MainWindowHainanStage1WorkflowController` and `MainWindowHainanEmployeePowerRewardWorkflowController` now own Hainan Stage 1 / clean-power and employee reward orchestration; Core tests passed (26), Excel tests passed (27), and WPF Debug/Release builds passed.
- Multi-province code-quality sixth slice on 2026-07-08: employee reward Core/Excel/WPF workflow names now explicitly use `HainanEmployeePowerReward...`; model fields use `ResponsiblePerson`, `ProjectDeveloper`, and `MonthlyPowers`; Core tests passed (26), Excel tests passed (27), and WPF Debug/Release builds passed.
- Multi-province code-quality seventh slice on 2026-07-08: Hainan Stage 2 Core/WPF contracts now explicitly use `HainanStage2...`; workflow entry points use `Analyze/Plan/Complete/RunHainanStage2`; Hainan audit issue construction moved to `HainanStage2AuditIssueFactory` while `Stage2SettlementCalculator` remains shared amount calculation/formatting; Core tests passed (26), Excel tests passed (27), and WPF Debug/Release builds passed.
- Multi-province code-quality eighth slice on 2026-07-08: project-wide naming cleanup reached the current low-risk endpoint; Hainan Stage 1 Core/Excel contracts, Hainan ledger layout, Hainan raw-detail readers, Hainan Stage 2 detail rows, and Chongqing Stage 1 WPF private workflow methods now use explicit province names; Core tests passed (26), Excel tests passed (27), and WPF Debug/Release builds passed.
- Documentation status/package refresh on 2026-07-08: `AGENTS.md` and `README.md` now describe Chongqing Stage 2 as a preflight-confirm-generate first implementation instead of Analyze-only; WPF package README no longer advertises Win7/8 packaging as a normal path; docs guardrails passed, stale wording scan had no matches, Core tests passed (26), Excel tests passed (27), WPF Debug build passed, WPF Release build passed, and WPF package script created `HainanSettlementTool-Win10-11-Release-20260708-155953.zip`.
- Chongqing Stage 2 practical test feedback on 2026-07-09: user reported the 20260708-155953 package completed重庆阶段二 and produced normal summary/report outputs with no blocking issues; some special customers still need manual adjustments.
- Chongqing Stage 2 backtest runner on 2026-07-09: `scripts/run_chongqing_stage2_backtest.ps1` was added and pushed in `f16872a`; parser check passed, Release build passed, fake-path no-real-data run wrote a failed summary successfully, docs guardrails passed, and `git diff --check` passed.
- Multi-province code-quality ninth slice on 2026-07-09: WPF `MainWindowChongqingStage1WorkflowController` now owns Chongqing Stage 1 clean-power and ledger-update orchestration; `MainWindow.xaml.cs` no longer contains those workflow methods and is down to 637 lines; Core tests passed (26), Excel tests passed (27), WPF Debug build passed, and WPF Release build passed.
- Chongqing Stage 2 backtest-fix pass on 2026-07-09: fixed target-month split/payment-party sheet preservation, preserved existing payment-party hidden columns, synced C-G power for safe refund extra-block rows, and added WPF post-run review popup for warnings/audit issues. Final real 3/4/5月回测 completed with 0 failures and 0 decision blockers; Stage 1 power diffs remained 0; Stage 2 hidden-column diffs cleared; refund extra-block C-G numeric diffs cleared. Excel tests passed (30), Core tests passed (26), WPF Debug build passed, WPF Release build passed, and WPF package script created `HainanSettlementTool-Win10-11-Release-20260709-113859.zip`. Remaining `ReviewRequired` items are non-blocking date/formula text differences and 3月 historical manual deduction/actual-payment differences.
- Chongqing Stage 2 practical UI/summary-format pass on 2026-07-09: optimized the WPF review popup, moved target-month `清能/清辉` sheets after `汇总表`, merged target sheet first-row titles, made startup province blank, changed Chongqing main tab to `代理费结算`, reserved disabled重庆 `员工电量奖励`, and compacted the main function area. Excel tests passed (30), Core tests passed (26), WPF Debug build passed, WPF Release/package build passed, and WPF package script created `HainanSettlementTool-Win10-11-Release-20260709-150129.zip`.
- Chongqing Stage 2 acceptance closeout on 2026-07-10: user reported current data validation had no significant issue; Core tests passed (26), Excel tests passed (30), build portability passed, and full Debug/Release solution builds passed before fast-forwarding local `main`.
- Guangdong month-sheet preparation focused validation on 2026-07-10: Core tests passed (29), Excel tests passed (35), and WPF Debug build passed. Authorized proxy-folder Analyze found 41 workbooks: 27 create, 14 already prepared, 0 skipped. Temporary-output smoke generated 41/41 successfully; repeat Analyze classified all 41 as already prepared, then the temporary real-data copies were deleted.
- Guangdong final validation on 2026-07-10: build portability passed; full Debug and Release solution builds passed; WPF package script passed and created `HainanSettlementTool-Win10-11-Release-20260710-123600.zip`. Package inspection found the executable, config, 17 DLLs and 20 total entries.

For new code changes, rerun focused tests and builds. For pure documentation changes, run at least `git diff --check` plus targeted stale-wording/link scans.

## Next Steps

1. Ask the user to test `dist\HainanSettlementTool-Win10-11-Release-20260710-123600.zip`.
2. Guangdong test focus: province-specific input visibility, preflight counts, output folder tree, strict standard sheet selection, C-F clearing, date update and report readability.
3. Keep Guangdong amount calculation, ledger integration and summary workbook generation out of this first slice.
4. Retry pushing updated `main` and the Guangdong branch when the GitHub connection is stable. Do not publish a formal release unless the user asks.
