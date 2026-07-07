# Handoff

Last updated: 2026-07-07

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

## Current Git State

- Current branch: `codex/chongqing-stage2-analysis`
- Upstream: `origin/codex/chongqing-stage2-analysis`
- Branch purpose: documentation quality cleanup plus Chongqing Stage 2 design/implementation preparation.
- The branch is based on `main` after the WPF log-controller state.
- This branch continues the documentation quality cleanup that started at `48c4921 Organize documentation map and current references`.
- Do not merge to `main`, tag, or publish a release without explicit user authorization.

- No formal release tag has been cut after `v1.0.1`.
- Completed milestone history is summarized in `docs/CHANGELOG.md`.

## Current Packages

Latest local WPF test package built after the WPF log-controller merge:

- `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-150019.zip`
- Unpacked directory: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-150019`

Latest local acceptance packages before the log-controller merge:

- Win10/11 WPF: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win10-11-Release-20260707-140505.zip`
- Win7/8 WinForms maintenance: `D:\Document\文件处理\hainan-settlement-desktop\dist\HainanSettlementTool-Win7-8-Release-20260707-140515.zip`

User practical testing on 2026-07-07 reported no blocking issues so far with:

- `HainanSettlementTool-Win10-11-Release-20260707-140505.zip`

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

Chongqing Stage 1:

- Implemented in Win10/11 WPF.
- Supports power cleaning and ledger update.
- Unmatched power customers must be explicitly handled in preflight: create, skip, or match one existing ledger customer.
- Create/skip decisions can repeat; a concrete existing ledger customer target can be used only once per preflight.
- Newly created Chongqing ledger rows write only customer name and target-month power in the output copy. They do not auto-fill B-column account number, `电力用户编码`,负责人,代理/居间 fields, or other manual fields.

Chongqing Stage 2:

- Partially implemented for Core contract and Excel preflight analysis only; not user-runnable yet.
- Analysis is documented in `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`.
- First Core contract slice is implemented: `ChongqingStage2Options`, preflight/report/issue models, `ChongqingStage2Service`, `IChongqingStage2ExcelGateway`, and workflow plan/complete/run entry points.
- Excel Analyze-only slice is implemented: `ChongqingStage2SettlementGenerator` reads the Chongqing ledger target-month 30-column block, identifies proxy/intermediary/refund groups, and reports new summary subjects requiring payment-party selection. Workbook generation still explicitly throws as not implemented, and WPF entry is still not implemented, so the feature is not user-runnable yet.
- Keep this work in Win10/11 WPF plus Core/Excel shared layers; do not add WinForms parity unless explicitly requested.

WinForms:

- Win7/8 WinForms is maintenance mode only.
- Keep it buildable/packageable and fix blocking bugs, but route new province UX work to WPF.

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
- Some `.xlsx` files in authorized Chongqing template folders were unreadable non-xlsx containers; preflight/report must not silently ignore them.
- Existing summary rows should inherit template fixed fields and long-term fields.
- New Chongqing summary subjects should not silently default to Hainan `清辉`. Recommended first implementation: WPF preflight requires payment-party selection (`清能`/`清辉`) for new summary subjects, and generated reports flag defaulted fields for manual review.
- First validation target should be a read-only 5月 replay against the manual 5月 result, with outputs only in an explicitly authorized test directory.

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

For new code changes, rerun focused tests and builds. For pure documentation changes, run at least `git diff --check` plus targeted stale-wording/link scans.

## Next Steps

1. Implement `ChongqingStage2SettlementGenerator` behind `IChongqingStage2ExcelGateway`, starting with synthetic workbook tests for ledger month-block parsing and report/preflight output.
2. Add WPF entry only after Excel preflight/generation can run on synthetic fixtures.
3. Use the first validation target from the Chongqing Stage 2 note: read-only 5月 replay and compare against the manual 5月 result, writing outputs only to an explicitly authorized test directory.
4. Defer persistent customer alias tables, WinForms parity, and cross-province generic Stage 2 abstraction unless the user explicitly reprioritizes them.
