# Handoff

Last updated: 2026-07-22

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

Then read, in order:

1. `AGENTS.md`
2. `HANDOFF.md`
3. `docs/README.md`
4. The owning document selected by the router

For the current Stage 2 integrity work, also read:

- `CONTEXT.md`
- `docs/architecture.md`
- `docs/dev-notes/stage2-preflight-integrity-2026-07-22.md`
- `docs/hainan-stage2-current-behavior.md`
- `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`

## Data Safety

- Never commit real ledgers, customer data, screenshots, payment data, settlement outputs, or generated files from real workbooks.
- Never overwrite a user workbook. Write only to an explicit output folder or clearly named copy.
- Real Hainan root: `C:\Users\juqx2\Desktop\2026海南`.
- Real Chongqing root: `C:\Users\juqx2\Desktop\2026年-重庆`.
- Guangdong analysis roots are also real business data.
- Historical authorizations are **not standing permission for new reads**. Obtain current explicit authorization for the concrete file/folder and read-only scope before any real-data benchmark, smoke, or comparison.
- Automated validation on 2026-07-22 used only repository-owned synthetic workbooks. For the later loan-balance diagnosis, the user explicitly authorized one exact generated Hainan summary workbook; it was inspected read-only, and no other real workbook or settlement directory was read.

## Git State

- Current development branch: `codex/stage2-preflight-integrity`.
- It was created from local `main` at `4baaf88` (`Integrate Guangdong safety fix into local main`).
- Local `main` is three commits ahead of `origin/main`: `23e79ff`, `0089f58`, and `4baaf88`; none has been pushed.
- The Stage 2 integrity implementation and documentation are committed on the `codex/` branch as `ed9452a` (`Add Stage 2 preflight and batch integrity safeguards`).
- Artifact organization is committed as `8b2de9b` (`Organize package and local validation artifacts`). The Hainan new-subject preflight UX/template-choice follow-up and its selected-template display fix are committed on the current branch.
- The borrowed-template comment cleanup, shared readable HTML renderer, Hainan/Chongqing HTML batch reports, WPF report button, Chongqing self-operated regression fix, tests, and documentation are committed as `f495d5d` (`Add readable Stage 2 reports and settlement fixes`).
- The current branch tip prepares the WPF product/file version and release documentation for `v1.2.0`.
- On 2026-07-22 the user confirmed practical testing across all provinces found no further issue and explicitly authorized commit, merge to `main`, push, tag, and formal release.
- Preserve the Window Separation / `outputs` exclusions in this handoff and do not revert them.
- Future push, merge, tag, package, or publication again requires a new explicit authorization after this release closes.

## Release State

- Latest formal release: `v1.2.0`.
- Release: `https://github.com/LeBronJu/retail-power-settlement-desktop/releases/tag/v1.2.0`.
- Formal asset: `RetailPowerSettlementTool-Win10-11-v1.2.0.zip`; only the maintained Win10/11 WPF entry is published.
- The preceding local Win10/11 test package is `dist/test-packages/RetailPowerSettlementTool-Win10-11-Release-20260722-172921-704.zip`; SHA-256 is `273477D21A1F97AE5CB1AC88BEF83429B0AD60A076A7042B36D97954934E6798`.
- `dist/` now contains distributable artifacts grouped under `test-packages/`, `releases/`, and `legacy-win7-8/`; local backtests and smoke outputs belong under the Git-ignored `local-validation/` tree.

## Current Product State

Hainan / Chongqing Stage 2:

- The 2026-07-22 branch implements the final rules in `docs/dev-notes/stage2-preflight-integrity-2026-07-22.md`.
- Summary identity is normalized `(subject, settlement kind)`; multi-owner data aggregates once and uses the first physical relationship row's owner.
- Proxy/intermediary relationship fields are validated before amount filtering; one identity must have one withholding tax rate.
- Full payee cell text is opaque and is never split into people. Conflicting non-empty complete texts block; blank plus one unique value continues with review.
- Existing payment party inherits from reliable fields/sheet membership. New or unresolved payment party requires an explicit Qingneng/Qinghui choice; there is no Qinghui default. Hainan Jingyan's confirmed Qingneng override remains.
- New subject defaults are no delegation, payee equals subject, platform invoice, ledger withholding rate, 13% total rate, `H=J-I`, first owner, and target month, with review reminder.
- Hainan and Chongqing use a shared four-level WPF preflight presentation, while province-specific workbook rules remain in separate Excel adapters. The window groups all issues for one `(settlement kind, subject)` into one card and folds full paths into technical details.
- Hainan new subjects with several non-exact same-kind split templates now require an explicit template choice in the same card as the Qingneng/Qinghui choice; zero candidates and duplicate exact historical templates still block. The chosen template only supplies layout, and the new output keeps only the target-month sheet. Chongqing has not opted into this province rule and keeps its prior blocker behavior.
- The shared WPF ComboBox template now forwards `ItemTemplateSelector`; selected template options therefore keep their friendly subject/owner/file label instead of falling back to the ViewModel type name.
- Hainan borrowed-template output for a new subject keeps only the target-month sheet and removes legacy cell comments copied from the source subject. Exact matched templates for existing subjects keep their own comments.
- Hainan loan remaining is written as `MAX(0, loan - deducted)`; when the program-owned current-month deduction first clears a positive balance and the completion month is blank, the concrete target settlement month is written as a date value and displayed as `yyyy年m月` (for example, a June 2026 run writes `2026年6月`). Existing completion months are preserved and historical missing months are not guessed.
- Analyze is read-only. A plan freezes preflight signature and file/output-state fingerprint; changes after confirmation force a new preflight.
- All workbooks and reports are generated in one staging workspace, strongly verified, then published as one batch. Failed staging is retained under `【未完成-禁止付款】`.
- There is no payment-routing JSON and no partial formal summary generation for Hainan/Chongqing.
- Hainan and Chongqing now publish a self-contained readable HTML report inside the same all-or-nothing batch as workbooks and JSON/TXT reports. The WPF completion card opens it directly. Guangdong uses the same escaped responsive renderer, while each province keeps its own report facts, status mapping, and publish semantics.
- Chongqing self-operated rows keep project developer/owner as internal attribution but produce no proxy/intermediary settlement. The integrity regression that treated project developer as a residual proxy subject is fixed; only real monthly proxy parameters or intermediary fields block. Refund settlement remains independent.

Guangdong:

- Proxy/intermediary/refund month-sheet preparation and its safety hardening passed user practical testing.
- Full Guangdong settlement remains paused.
- Guangdong's independent-workbook partial-completion behavior must not be generalized to Hainan/Chongqing formal settlement.
- Guangdong's existing readable HTML now uses the shared escaped responsive renderer without changing its independent-workbook partial-completion semantics.

Performance / UI support:

- Hainan and Chongqing performance optimization is paused because monthly volume is small. Do not resume it unless the user explicitly reopens it.
- Win7/8 WinForms remains frozen. Win10/11 WPF is the maintained entry.

## Latest Validation

Current branch, synthetic inputs only:

- Core tests: 80/80 passed.
- Excel tests: 126/126 passed.
- WPF presentation/control tests: 4/4 passed.
- Focused Hainan Stage 2 tests: 57/57 passed.
- Focused Chongqing Stage 2 tests: 40/40 passed.
- Full Debug and Release solution builds passed.
- WPF Debug build passed with 0 warnings / 0 errors.
- Build portability check passed.
- Documentation guardrails and `git diff --check` passed.
- Independent read-only diff review found no remaining P0/P1 in the current scope; its two P2 findings (card-count wording and whitespace-normalized subject keys) were addressed.
- The latest `172921-704` Win10/11 test package build passed; its directory and ZIP each contain the same 20 files, no workbook/CSV/JSON/log data, and the packaged WPF executable remained running through a five-second hidden startup smoke.
- One user-authorized generated Hainan summary workbook was inspected read-only for the reported loan-balance row. The later exact user-authorized new-agent split workbook was inspected read-only and contained one legacy comment at `6月!P6`; no surrounding business directory was scanned and the source workbook was not modified. Automated tests remained synthetic. No `outputs` material was read or produced.
- For the later Chongqing self-operated regression, the user explicitly authorized one exact historical ledger. It was inspected and replayed read-only: all seven reported rows had project developer/owner attribution but no monthly proxy/intermediary parameters; the fixed reader reports zero self-operated blockers. No surrounding directory was scanned and the workbook was not modified.

## Known Boundary

- Chongqing refund templates that contain both a standard month sheet such as `2` and a copy-style sheet such as `2 (2)` remain a documented pre-existing technical risk. The current task does not change that rule; obtain business confirmation before choosing a new policy.
- Hainan/Chongqing intentionally remain all-or-nothing for formal Stage 2 output. A future non-payable partial draft would require a separate design and explicit authorization.
- Chongqing loan deduction remains template/manual-driven, so the Hainan automatic zero-clamp/completion-month behavior was intentionally not generalized there; doing so could hide a real manual over-deduction.
- The program does not infer payment routing from a multi-person payee cell; cashiers continue to handle distribution agreements outside this program.
- The Chongqing backtest scripts still need adaptation to the new required preflight signature/input fingerprint before they can be treated as runnable regression tools; their output paths are already isolated under `local-validation/backtests/`.

## Next Session

1. Run the required reading gate from clean `main` and confirm it matches `origin/main` after the `v1.2.0` closeout.
2. Preserve the `outputs/` exclusions and obtain new explicit authorization before any future merge, push, tag, release, or real-workbook read.
3. Treat `v1.2.0` as the current production baseline; future settlement changes should start from a new `codex/` branch and synthetic regression tests.
4. Do not resume Hainan/Chongqing performance work, add routing JSON, change the ledger format, or add partial formal summaries without a new explicit decision.

## Window Separation

- The other conversation remains dedicated to boss-report material under `outputs/boss-report-2026-07-14`.
- This conversation owns C# program development. Do not modify report materials unless the user explicitly asks this development window to synchronize a fact.
- `outputs/` is locally Git-excluded. Report PPT/Markdown/Word/PDF files and screenshots must not be staged, committed, or pushed.

## Documentation Maintenance

Use `docs/README.md` as the document router and `docs/CHANGELOG.md` for completed history. Keep this file current and concise; detailed rules belong in the current behavior/dev-note documents and detailed history belongs in git.
