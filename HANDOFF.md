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

The latest completed release work is Guangdong Stage 1. Before reviewing or changing it, also read:

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
- Historical real-data authorizations are not standing permission for new reads unless this file records a later explicit extension. Obtain current explicit authorization for every other exact file/folder and read-only/output scope before real-data analysis, benchmark, smoke, or comparison.
- On 2026-07-24 the user explicitly extended permanent read-only authorization for exactly two Guangdong monthly retail-detail workbooks for May/June 2026 and three Guangdong ledger snapshots dated 2026-07-07, 2026-07-13, and 2026-07-16. This does not authorize surrounding-directory scans, writes, generated real-data copies, or any other workbook. They were not reread during the implementation pass; all automated tests use synthetic workbooks.

## Git State

- The `v1.4.0` tag points to exact release-source merge commit `02760d18b90235410908d4eb87b92f538eb8b8af`.
- `codex/stage2-template-browser` is pushed and merged. Its main commits are `933de14` (open generated splits on the target month) and `f1ec9e3` (scalable Stage 2 template selection).
- Guangdong Stage 1 research documentation was merged to `main` at `8fa2534`; the implementation and cross-province Stage 1 clearing fixes are released in `v1.4.0`.
- Development branch `codex/guangdong-stage1` was created from `8fa2534`, committed, pushed, and merged for `v1.4.0`.
- Any future code change must start from current local `main` on a fresh `codex/` branch.
- The user explicitly authorized the `v1.4.0` commit, push, merge, tag, formal package, and GitHub Release in the current turn. That authorization is consumed after this release; future external mutations require new authorization.

## Release State

- Latest formal release: `v1.4.0`.
- Release: `https://github.com/LeBronJu/retail-power-settlement-desktop/releases/tag/v1.4.0`.
- Formal asset: `RetailPowerSettlementTool-Win10-11-v1.4.0.zip`.
- Local formal package: `dist/releases/RetailPowerSettlementTool-Win10-11-v1.4.0.zip`.
- SHA-256: `66DA8810A4AD620323BA3AD30EAA2220DA911ADD9C27B64D77A27B3706E37C99`.
- Only the maintained Win10/11 WPF entry is published. Win7/8 WinForms remains frozen.
- `dist/` separates `test-packages/`, `releases/`, and `legacy-win7-8/`; local research/backtests belong under Git-ignored `local-validation/`.
- Latest Guangdong Stage 1 development test package: `dist/test-packages/RetailPowerSettlementTool-Win10-11-Release-20260724-163923-947.zip`.
- Development test package SHA-256: `A16435F140D64EF39813EA668308E9CAA7D14E3BC0C3D156C9D9F8EEC8353F84`. It contains the first-test-feedback fixes and the approved preflight/coefficient wording, and is not a formal release asset.
- The earlier `154432-647` and `162850-177` packages are obsolete and must not be used for latest-behavior validation.

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

- `v1.4.0` formally includes Guangdong Stage 1 in WPF as “整理本月电量” and “生成本月台账”, alongside the previously released proxy/intermediary/refund workbook month-sheet preparation.
- The first local test package exposed usability/output issues; the released implementation contains the numeric-zero, combined clean-workbook output, HTML report, concise preflight, and approved wording fixes.
- The original downloaded monthly workbook contains only the three official sheets `零售结算明细`, `市场联动价格`, and `零售合同模式`. The researched working copies had two later settlement-staff sheets appended; those manual sheets must be ignored by the program.
- Stage 1 reads only exact sheet `零售结算明细`; the other official sheets and arbitrary manual sheets are ignored.
- It matches by customer code, aggregates official detail by code in MWh, validates decimal conservation at `0.00001 MWh`, uses the first valid complete coefficient pair in source order, and maintains a dedicated Guangdong 2026 ledger-month implementation. It does not reuse Chongqing name matching or either province's updater.
- Existing target months clear all customers' four power cells plus coefficients before write; ledger-only customers receive numeric zero in all four power cells and keep their prior coefficient pair. New codes append at the bottom, inherit style/formulas only, and receive sequence/code/name/`2026MM`/power/coefficient safe fields.
- The combined action outputs the eight-column clean workbook as well as the ledger. Both Stage 1 actions now create self-contained HTML reports; the WPF confirmation groups issues into concise focus cards, shows automatic handling separately, and keeps full per-customer detail collapsed.
- Detailed evidence, confirmed rules, preflight categories, architecture seams, and tests are in `docs/dev-notes/guangdong-stage1-research-2026-07-24.md`.

Other boundaries:

- Hainan and Chongqing existing target months now also clear every ledger customer's target-month power before writing. Hainan new rows retain formulas/styles but clear non-formula constants.
- Customer name keys now use NFKC and remove whitespace/zero-width format characters while retaining punctuation.
- Hainan/Chongqing performance optimization is paused because monthly volume is small.
- 2027 support is researched but deferred; do not implement it without a new decision.
- The Chongqing backtest scripts still need adaptation to the required Stage 2 preflight signature/input fingerprint.

## Latest Validation

`v1.4.0` release validation:

- Release Core tests: 87/87.
- Release Excel tests: 136/136.
- Release WPF tests: 14/14.
- Full Debug and Release solution builds: 0 warnings, 0 errors.
- Documentation guardrails, build portability, and `git diff --check` passed.
- Tests and builds used synthetic/local code inputs only. The five authorized real Guangdong workbooks and `outputs/` were not read during implementation or validation.
- The formal package directory and ZIP contain the same 20 flat files and no workbook/CSV/JSON/log/HTML/PPTX/DOCX/PDF/image/PDB business material.
- Packaged executable reports FileVersion `1.4.0.0`, ProductVersion `1.4.0+02760d18b90235410908d4eb87b92f538eb8b8af`, and passes a five-second hidden startup smoke.
- GitHub Release `v1.4.0` is published as Latest, is neither draft nor prerelease, and contains exactly the Win10/11 WPF asset. GitHub reports digest `sha256:66da8810a4ad620323ba3ad30eaa2220da911add9c27b64d77a27b3706e37c99`, matching the local package.

Guangdong Stage 1 implementation:

- All implementation tests use repository-owned synthetic workbooks; the five authorized real workbooks were not reread and no real-data output was created.
- Clean-only output is an eight-column workbook plus JSON and self-contained HTML. The combined action outputs the same clean workbook together with the updated ledger, JSON, and HTML.
- Synthetic coverage includes three-official-sheet downloads with no staff sheets, arbitrary extra sheets, header-order variation, same-code aggregation, Unicode names, coefficient conflict order, conservation tolerance/blocking, same-code name conflict, existing/missing 32-column month blocks, full clearing, ledger-only numeric-zero power with coefficient preservation, safe new rows, source immutability, repeat runs, combined clean-workbook output, HTML guidance, and grouped/collapsed WPF confirmation presentation.
- Targeted cross-province tests also cover Hainan/Chongqing full clearing and Hainan formula-only template inheritance.

## Guangdong Stage 1 Confirmed Decisions

1. Append new customers at the bottom of the customer area.
2. Inherit the last valid customer row's style/formulas; clear all non-formula constants.
3. Write new-customer `履约开始月份` as target `2026MM`.
4. Ignore `零售合同模式` and `市场联动价格` in Stage 1; manual business fields stay blank.
5. Add only new codes present in current-month detail; do not add contract-only zero-power rows.
6. Match exclusively by code. Same code/name difference writes by code without renaming; same name/different code never auto-merges.
7. Normalize names with NFKC and invisible whitespace removal, without punctuation stripping or fuzzy matching.
8. Produce the standalone eight-column clean workbook.
9. Use decimal with `0.00001 MWh` reconciliation tolerance and no intermediate rounding.
10. Coefficients do not affect proxy-fee calculation; take the first valid complete pair in source order.
11. Existing target month: clear all customers' power and coefficients, rewrite source customers, and restore coefficients for ledger-only/no-valid-source-coefficient customers.

## Suggested Next Session

1. Resume from current `main`; run `git status --short --branch` before any work and create a fresh `codex/` branch for code changes.
2. Treat `v1.4.0` and `RetailPowerSettlementTool-Win10-11-v1.4.0.zip` as the production baseline.
3. The timestamped Guangdong Stage 1 packages are superseded by the formal package. Do not push, merge, tag, create another package, or publish unless the user explicitly asks.
4. Do not write outputs from the five real Guangdong workbooks. Their standing authorization is read-only only; any real-data smoke output needs separate explicit scope.

## Window Separation

- The other conversation remains dedicated to boss-report material.
- This conversation owns C# program development.
- Do not synchronize or edit report materials unless the user explicitly asks this development conversation to do so.

## Documentation Maintenance

Use `docs/README.md` as the router and `docs/CHANGELOG.md` for completed history. Keep this file concise; detailed rules belong in current behavior/dev-note documents.
