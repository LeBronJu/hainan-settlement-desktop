# Handoff

Last updated: 2026-07-10

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

- `main` contains Guangdong month-sheet preparation, WPF third-province support, and the current documentation/performance research closeout.
- The historical implementation branch is `codex/guangdong-stage2-month-preparation`; new development should not continue on it.
- The next development branch must be created from the updated `origin/main` with the `codex/` prefix.
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

- ZIP: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260710-123600.zip`
- SHA256: `70939036D93BC85DAF6069D1C536531BCB3B04F7D73330027CBB3D6BA18B3B03`
- This is a user-test package, not a formal release asset.
- User simple acceptance passed. A real Guangdong batch of more than 600 workbooks completed in a little over four minutes.

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

Most recent code/package validation for this branch:

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

The documentation closeout after that package changes no business code. Its required validation is the docs guardrail script, stale-reference scans, and `git diff --check`.

## Next Session

1. Create `codex/hainan-excel-performance-baseline` from the updated `origin/main` before code changes; do not continue development directly on `main` or the historical Guangdong branch.
2. Treat Hainan performance baseline/instrumentation as the next mainline. Do not begin with Open XML writer replacement.
3. Measure Stage 1, Stage 2, and employee reward separately: scan/read, business calculation, workbook mutation, formula evaluation, save, total time, and peak memory where practical.
4. Use synthetic workbooks unless the user provides a new explicit real-data authorization.
5. After measurement, implement the smallest behavior-preserving removal of repeated reads or saves and rerun existing tests/builds.
6. Only then implement one Open XML read-only shadow Adapter and compare DTO output against ClosedXML.
7. Keep Guangdong full settlement, formal release/tagging, WinForms parity, and structural Open XML writers out of the first performance slice.

## Documentation Maintenance

Use `docs/README.md` as the document router and `docs/CHANGELOG.md` for completed history. Keep this file current and concise; detailed investigations belong in dated dev notes or git history.
