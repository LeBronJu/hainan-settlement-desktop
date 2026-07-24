# Handoff

Last updated: 2026-07-24

## Project

Standalone C# Win10/11 desktop tool for multi-province retail electricity settlement automation.

- Repository: `D:\Document\文件处理\hainan-settlement-desktop`
- GitHub: `https://github.com/LeBronJu/retail-power-settlement-desktop`
- Historical Python reference: `D:\Document\文件处理\hainan-settlement-tool`
- Stable Hainan reference folder: `D:\Document\文件处理\稳定参考版海南结算`

## Start Here

Every resumed session must first run:

```powershell
git status --short --branch
```

Then read completely, in order:

1. `AGENTS.md`
2. `HANDOFF.md`
3. `docs/README.md`
4. The owning documents selected by the router

The next planned work is Guangdong Stage 1. Before discussing or implementing it, also read:

- `CONTEXT.md`
- `docs/architecture.md`
- `docs/dev-notes/guangdong-stage1-research-2026-07-24.md`
- `docs/dev-notes/multi-province-readiness-2026-07-07.md`
- `docs/dev-notes/settlement-year-2027-readiness-2026-07-23.md` only if the user reopens 2027

## Data Safety

- Never commit real ledgers, customer data, screenshots, payment data, settlement outputs, or generated files from real workbooks.
- Never overwrite a user workbook. Write only to an explicit output folder or clearly named copy.
- Real Hainan root: `C:\Users\juqx2\Desktop\2026海南`.
- Real Chongqing root: `C:\Users\juqx2\Desktop\2026年-重庆`.
- Guangdong analysis roots are also real business data.
- `outputs/` contains boss-report material owned by another conversation. It is locally Git-excluded; do not read, modify, stage, upload, or publish it.
- Historical authorizations are **not standing permission for new reads**. Obtain current explicit authorization for the exact file/folder and read-only/output scope before any real-data analysis, benchmark, smoke, or comparison.
- On 2026-07-24 the user explicitly authorized read-only research of exactly two Guangdong monthly retail-detail workbooks for May/June 2026 and three Guangdong ledger snapshots dated 2026-07-07, 2026-07-13, and 2026-07-16. They were not modified or uploaded, no surrounding directory was scanned, and local analysis artifacts remain Git-ignored. This permission expires with that research task.

## Git State

- Latest formal release source is merge commit `7f29bd216674f20d0f893d2c3d8db8f646bbab97`.
- Tag `v1.3.0` points exactly to that release source.
- `codex/stage2-template-browser` is pushed and merged. Its main commits are `933de14` (open generated splits on the target month) and `f1ec9e3` (scalable Stage 2 template selection).
- Guangdong Stage 1 research documentation lives in post-release repository history and was originally prepared on `codex/guangdong-stage1-research`; it does not change the `v1.3.0` tag or package.
- At handoff, `main` should be clean and aligned with `origin/main`. Verify instead of assuming.
- Any future code change must start from current local `main` on a fresh `codex/` branch.
- Do not push, merge, tag, package, or publish without the user's current authorization. The 2026-07-24 release authorization has been consumed.

## Release State

- Latest formal release: `v1.3.0`.
- Release: `https://github.com/LeBronJu/retail-power-settlement-desktop/releases/tag/v1.3.0`.
- Formal asset: `RetailPowerSettlementTool-Win10-11-v1.3.0.zip`.
- Local formal package: `dist/releases/RetailPowerSettlementTool-Win10-11-v1.3.0.zip`.
- SHA-256: `895CC8D79EC0A37B36A7E11E51E38DAAB3D8506F37EDFE863C2C35EBBEFB27DF`.
- Only the maintained Win10/11 WPF entry is published. Win7/8 WinForms remains frozen.
- `dist/` separates `test-packages/`, `releases/`, and `legacy-win7-8/`; local research/backtests belong under Git-ignored `local-validation/`.

## Current Product State

Hainan / Chongqing Stage 2:

- The final Stage 2 rules live in `docs/dev-notes/stage2-preflight-integrity-2026-07-22.md`.
- Shared WPF preflight uses Blocker / RequiredDecision / Review / Information and groups one `(settlement kind, subject)` into one readable card.
- Existing payment party inherits; new or unresolved payment party requires explicit Qingneng/Qinghui choice. There is no Qinghui default.
- Payee cell text is opaque and may contain several people; only conflicting complete non-empty texts block.
- Hainan and Chongqing formal Stage 2 remain all-or-nothing. They do not create routing JSON or partial formal summaries.
- Both provinces publish self-contained HTML plus JSON/TXT reports inside the same verified batch as workbooks.
- New-subject template choice now scales to large catalogs: random batches of at most five without repeats within a round, reshuffle after exhaustion, and full-catalog whitespace-term AND search across subject, owner, and file name.
- Hainan/Chongqing generated splits and normal Guangdong month-preparation outputs open on the target month as the only active/selected sheet. Summary workbooks are unchanged.
- Hainan borrowed templates for new subjects keep only the target month and remove source comments; exact historical templates preserve their own comments.
- Hainan loan remaining is clamped to zero and the concrete settlement year/month is written when the current run first clears a positive balance.
- Chongqing self-operated rows keep internal owner/developer attribution but do not create proxy/intermediary settlement.

Guangdong:

- The released function is still only proxy/intermediary/refund workbook month-sheet preparation with independent-workbook partial-completion semantics and HTML/JSON/TXT reporting.
- Guangdong Stage 1 has been researched but is **not implemented or visible in WPF**.
- The original downloaded monthly workbook contains only the three official sheets `零售结算明细`, `市场联动价格`, and `零售合同模式`. The researched working copies had two later settlement-staff sheets appended; those manual sheets must be ignored by the program.
- Guangdong Stage 1 must match by customer code, aggregate official detail by code, and use a dedicated 32-column ledger-month implementation. It must not reuse Chongqing name matching or either province's updater.
- Detailed evidence, safe first scope, preflight categories, architecture seams, and open decisions are in `docs/dev-notes/guangdong-stage1-research-2026-07-24.md`.

Other boundaries:

- Hainan/Chongqing performance optimization is paused because monthly volume is small.
- 2027 support is researched but deferred; do not implement it without a new decision.
- The Chongqing backtest scripts still need adaptation to the required Stage 2 preflight signature/input fingerprint.

## Latest Validation

Release code:

- Core tests: 83/83.
- Excel tests: 129/129.
- WPF tests: 11/11 in Debug and Release after final review fixes.
- Full Release solution build: 0 warnings, 0 errors.
- Build portability, documentation guardrails, and `git diff --check` passed.
- Formal package directory and ZIP each contain the same 20 files and no workbook/CSV/JSON/log/PPTX/DOCX/PDF business material.
- Packaged executable reports FileVersion `1.3.0.0`, ProductVersion `1.3.0+7f29bd...`, and passed a five-second hidden startup smoke.
- GitHub release asset digest matches the local SHA-256 above.

Guangdong Stage 1 research:

- Only the five explicitly authorized real workbooks were inspected read-only.
- Official detail reconciled to ledger total/peak/flat/valley in the researched months.
- Existing ledger-only customers correctly carried zero monthly power; therefore target power fields must be cleared before source aggregates are written.
- The later ledger snapshot adds a new 32-column month block and the same new customer codes seen in the contract source, while many business fields remain intentionally incomplete for operator follow-up.
- Automated Guangdong Stage 1 tests do not exist yet because no implementation has started. Future tests must use repository-owned synthetic workbooks.

## Guangdong Stage 1 Decisions Still Needed

Before implementing final ledger writes, confirm with the user:

1. New-customer row placement: append, or insert according to the existing business grouping/order?
2. Which row is the safe style/formula template, and which inherited fields must always be cleared?
3. Should a new customer's start month be written automatically as the target `YYYYMM`?
4. Should contract-mode fields be initialized from `零售合同模式`, or remain blank with a reminder as in the current manual process?
5. Should contract-only customers with no current-month power be added as zero-power rows?
6. Should detail-only customers missing from the contract sheet block, or generate with review?
7. Is `市场联动价格` completely ignored/report-only in the first version?
8. Does the user want a standalone six-column clean-power workbook as well as the updated ledger?
9. What decimal/rounding tolerance should govern total versus peak/flat/valley reconciliation?

Do not silently decide these by copying the observed manual result. Reader, aggregation, synthetic preflight, and output-safety work can be designed before the answers; new-row and contract-field writes cannot.

## Suggested Next Session

1. Verify clean `main`, then report current release and documentation state.
2. Discuss the eight Guangdong decisions above; several may collapse into one agreed “safe first version”.
3. After explicit authorization, create a fresh `codex/guangdong-stage1-*` branch.
4. Implement in slices: official-source reader and aggregation; synthetic preflight; clean-only output; Core/WPF neutralization; ledger month-block update; confirmed new-customer behavior.
5. Keep Guangdong workbook rules in its own Adapter and retain existing Hainan/Chongqing behavior.
6. Do not read the previously authorized five real workbooks again unless the user grants fresh authorization.

## Window Separation

- The other conversation remains dedicated to boss-report material.
- This conversation owns C# program development.
- Do not synchronize or edit report materials unless the user explicitly asks this development conversation to do so.

## Documentation Maintenance

Use `docs/README.md` as the router and `docs/CHANGELOG.md` for completed history. Keep this file concise; detailed rules belong in current behavior/dev-note documents.
