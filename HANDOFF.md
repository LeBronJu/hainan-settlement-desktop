# Handoff

Last updated: 2026-07-15

## Project

Standalone C# Win10/11 desktop tool for multi-province retail electricity settlement automation.

- Repository: `D:\Document\文件处理\hainan-settlement-desktop`
- GitHub: `https://github.com/LeBronJu/retail-power-settlement-desktop`
- Historical Python reference: `D:\Document\文件处理\hainan-settlement-tool`
- Stable Hainan reference folder: `D:\Document\文件处理\稳定参考版海南结算`

## Start Here

Every resumed session must run:

```powershell
git status --short --branch
```

Then read, in order:

1. `AGENTS.md`
2. `HANDOFF.md`
3. `docs/README.md`
4. The owning document selected by the task router in `docs/README.md`

For the next performance/Open XML work, also read:

- `CONTEXT.md`
- `docs/architecture.md`
- `docs/dev-notes/excel-performance-openxml-research-2026-07-10.md`
- the current behavior/module note for the Hainan workflow being changed

## Data Safety

- Never commit real ledgers, customer data, screenshots, payment data, settlement outputs, or generated files from real workbooks.
- Never overwrite a user workbook. Write only to an explicit output folder or clearly named copy.
- Real Hainan root: `C:\Users\juqx2\Desktop\2026海南`.
- Real Chongqing root: `C:\Users\juqx2\Desktop\2026年-重庆`.
- Guangdong analysis roots are also real business data.
- Historical authorizations are **not standing permission for new reads**. Obtain current explicit authorization for the concrete file/folder and read-only scope before any real-data benchmark, smoke, or comparison.
- Prefer synthetic workbooks and temporary output directories. Delete temporary real-data copies after an authorized run.

## Git State

- `main` contains the released Guangdong month-sheet preparation and WPF third-province baseline.
- The active safety-fix branch is `codex/guangdong-skipped-workbook-safety`, created from `origin/main` in an isolated worktree.
- The historical implementation branch is `codex/guangdong-stage2-month-preparation`; new development should not continue on it.
- Do not tag or publish a formal release without explicit user authorization.
- Formal release is `v1.1.0`.

## Latest Formal Release

- Tag: `v1.1.0`.
- Release: `https://github.com/LeBronJu/retail-power-settlement-desktop/releases/tag/v1.1.0`.
- Asset: `RetailPowerSettlementTool-Win10-11-v1.1.0.zip`.
- Local ZIP: `D:\Document\文件处理\hainan-settlement-desktop\dist\RetailPowerSettlementTool-Win10-11-v1.1.0.zip`.
- SHA256: `24B03BE8A708FF0C91166170DC87EF52C5EA7FBA4F7AF7D4F926140D193B20A0`.
- This release contains only the maintained Win10/11 WPF entry. Win7/8 WinForms remains frozen and is not packaged.

## Latest Test Package

- ZIP: `D:\Document\文件处理\hainan-settlement-desktop-gd-safety\dist\RetailPowerSettlementTool-Win10-11-Release-20260715-171227.zip`
- SHA256: `414817A401D10A3A36806D7D79DA055EE3D3649C3033B3593CC66A555A2FD66A`.
- Source: `codex/guangdong-skipped-workbook-safety` at code commit `23e79ff`.
- This is a local Win10/11 user-test package for the Guangdong skipped-workbook safety fix, not a formal release asset.
- Package inspection passed: one WPF executable, one config, 17 DLLs, one README, 20 ZIP entries, no workbook/report data, and the packaged executable passed a short startup smoke.

## Current Product State

Hainan:

- Stage 1, Stage 2, and employee power reward are the mature user-owned workflows.
- Stage 2 current behavior is defined in `docs/hainan-stage2-current-behavior.md`.
- The next technical mainline starts with Hainan performance measurement because the user owns and understands these rules.

Chongqing:

- Stage 1 and Stage 2 are available in WPF.
- Stage 2 passed practical testing and authorized March-May combined backtesting and is the first usable production baseline under long-term observation.
- Special manual refund/deduction cases remain explicit review items; do not hardcode one-month exceptions.
- Current details live in `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`.

Guangdong:

- Only proxy/intermediary/refund workbook month-sheet preparation is implemented.
- It strictly selects exact numeric month sheets, preserves non-standard sheets, normalizes an existing target month, or copies only the exact previous numeric month.
- Conflicting or unreliable dates remain conservative skip conditions and are not automatically corrected.
- Skipped or generation-failed input workbooks are preserved unchanged under the corresponding `【未处理-需人工复核】` directory; HTML/JSON/TXT reports and the WPF result state expose partial completion and input classification totals.
- It does not read a ledger, write settlement power, calculate amounts, or generate a summary workbook.
- The current feature and performance satisfy the immediate user request. Full Guangdong settlement and further performance work are paused while Hainan is studied first.
- Current rules live in `docs/dev-notes/guangdong-stage2-month-preparation-2026-07-10.md`.

WinForms:

- Win7/8 WinForms is frozen. Do not add feature, UI, parity, packaging, or refactor work unless the user explicitly reopens support.

## Performance And Open XML Direction

The durable research record is `docs/dev-notes/excel-performance-openxml-research-2026-07-10.md`.

Current conclusions:

- Do not perform a big-bang ClosedXML replacement.
- The project already receives `DocumentFormat.OpenXml 3.1.1` transitively from ClosedXML.
- First establish Hainan timing and workbook-lifecycle baselines without changing output behavior.
- Remove repeated reads/preflight and unnecessary saves before attributing cost to the library.
- First Open XML experiment should be a read-only shadow Adapter that compares normalized DTOs with ClosedXML.
- Keep `.xls/.csv` on existing ExcelDataReader/CSV adapters.
- Keep structural writers on ClosedXML until formula caches, hidden structures, sheet relationships, Excel/WPS compatibility, and semantic equivalence are proven.
- Hainan is a correctness pilot, not a Guangdong throughput proxy; use synthetic workbook batches for scale tests.

High-risk writers that must not be migrated first:

- `HainanStage1LedgerUpdater`
- `HainanStage2SplitWorkbookWriter`
- `HainanStage2SummaryWorkbookWriter`
- the write side of `HainanEmployeePowerRewardGenerator`

## Latest Validation

Most recent completed validation for the released baseline:

- Core tests: 29 passed.
- Excel tests: 35 passed.
- Build portability check passed.
- Full Debug and Release solution builds passed.
- The authorized Guangdong closeout gate reran all four checks after refreshing `origin`; all passed before the fast-forward to `main`.
- The `v1.1.0` release gate passed full Debug solution tests, Release build, build portability, documentation guardrails, and formal Win10/11 packaging.
- The formal package contains one neutral WPF executable, one config, 17 DLLs, and 20 ZIP entries; the executable file version is `1.1.0.0`.
- Authorized Guangdong sample smoke generated 41/41 outputs; repeat analysis classified all 41 as already prepared; temporary copies were deleted.
- Package inspection found the executable, config, 17 DLLs, and 20 ZIP entries.
- User then completed the larger 600+ workbook practical run without reporting a blocking correctness issue.

Current Guangdong safety-fix branch validation:

- Core tests: 29/29 passed.
- Excel tests: 41/41 passed, including 11 Guangdong month-preparation regressions using only synthetic workbooks.
- Full Debug and Release solution builds passed; the WPF partial-completion UI compiled successfully.
- Build portability, documentation guardrails, `git diff --check`, and debug-marker scan passed.
- The exact user-authorized incident workbook smoke produced 0 normal / 1 preserved skip / 0 failed: the review copy was SHA-256 identical, retained sheet `4`, did not add sheet `5`, and the HTML displayed partial completion. The temporary copy and output were deleted.
- No other real business workbook, ledger, customer data, or settlement output was read during this validation.

## Next Session

1. Review the committed `codex/guangdong-skipped-workbook-safety` fix and integrate it only after explicit user authorization.
2. Do not merge, tag, publish, create a test package, or create a formal release without explicit user authorization.
3. After this safety fix is closed, resume the Hainan performance baseline: measure Stage 1, Stage 2, and employee reward separately before any Open XML writer change.
4. Keep Guangdong full settlement, formal release/tagging, WinForms parity, and structural Open XML writers out of the first performance slice.

## Documentation Maintenance

Use `docs/README.md` as the document router and `docs/CHANGELOG.md` for completed history. Keep this file current and concise; detailed investigations belong in dated dev notes or git history.
