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
- `docs/dev-notes/stage2-template-candidate-browser-2026-07-23.md`

## Data Safety

- Never commit real ledgers, customer data, screenshots, payment data, settlement outputs, or generated files from real workbooks.
- Never overwrite a user workbook. Write only to an explicit output folder or clearly named copy.
- Real Hainan root: `C:\Users\juqx2\Desktop\2026海南`.
- Real Chongqing root: `C:\Users\juqx2\Desktop\2026年-重庆`.
- Guangdong analysis roots are also real business data.
- Historical authorizations are **not standing permission for new reads**. Obtain current explicit authorization for the concrete file/folder and read-only scope before any real-data benchmark, smoke, or comparison.
- Automated validation on 2026-07-22 used only repository-owned synthetic workbooks. For the later loan-balance diagnosis, the user explicitly authorized one exact generated Hainan summary workbook; it was inspected read-only, and no other real workbook or settlement directory was read.

## Git State

- Current development branch: `codex/stage2-template-browser`, created from clean local `main` at `db696e2`.
- Commit `933de14` on this branch brings in the already verified target-month active-sheet change from the previous branch. The current uncommitted work adds the shared Stage 2 template browser and enables Chongqing multi-candidate template decisions.
- Current production baseline remains `main` / `v1.2.0`; this branch has not been committed for the new browser scope, pushed, merged, tagged, or released. A local timestamped Win10/11 test package has been built for user verification.
- The previous `codex/stage2-active-month-sheet` commit was pushed as `488e6db` but remains unmerged; do not cherry-pick both `488e6db` and `933de14`, because they represent the same change.
- `codex/stage2-preflight-integrity` was merged into `main` as `3e4c6f9` (`Merge Stage 2 integrity and v1.2.0 release`).
- The main implementation commits are `ed9452a` (preflight/integrity), `8029a61` (new-subject preflight UX), `54549ab` (template display and Hainan loan status), `f495d5d` (readable reports and settlement fixes), and `ca0f26b` (release preparation).
- The previously local-only Guangdong safety commits and the v1.2.0 Stage 2 integrity branch are now pushed through `origin/main`.
- Tag `v1.2.0` points to the exact merge commit used to build the formal package. The later main commit only records post-release handoff state.
- Preserve the Window Separation / `outputs` exclusions in this handoff and do not revert them.
- Future push, merge, tag, package, publication, or real-workbook read again requires new explicit user authorization.

## Release State

- Latest formal release: `v1.2.0`.
- Release: `https://github.com/LeBronJu/retail-power-settlement-desktop/releases/tag/v1.2.0`.
- Formal asset: `RetailPowerSettlementTool-Win10-11-v1.2.0.zip`; only the maintained Win10/11 WPF entry is published.
- Local formal package: `dist/releases/RetailPowerSettlementTool-Win10-11-v1.2.0.zip`.
- Formal package SHA-256: `2C914708542D8D6C56336AA3C3A852421617B4379F4FC21AA15345C63E85F3A3`.
- The latest local Win10/11 test package is `dist/test-packages/RetailPowerSettlementTool-Win10-11-Release-20260724-104552-717.zip`; SHA-256 is `0CFB4B0B64471236856309F583E1BF54B4C7BF4D5A71453C23E6FFF2FF55E05C`.
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
- Both provinces now turn several non-exact same-kind split templates into an explicit required choice; zero candidates, duplicate exact historical templates, and unreadable templates still block. The complete candidate set remains signed and fingerprinted, and the selected path is revalidated and actually used by the province writer.
- The shared WPF browser shuffles the complete candidate set once per round, shows at most five per batch without repeats, supports previous/next and end-of-round reshuffle, and keeps an existing selection while browsing. Its full-catalog search uses case-insensitive whitespace-term AND matching across subject, owner, and file name, with Enter/double-click/Esc keyboard behavior.
- Template display and search use province-neutral file stem, parent folder, and file-name metadata; WPF does not read workbook contents or decide province template eligibility.
- Hainan borrowed-template output for a new subject keeps only the target-month sheet and removes legacy cell comments copied from the source subject. Exact matched templates for existing subjects keep their own comments.
- The current development branch ensures the target-month sheet is visible and is the only active and selected tab in generated Hainan/Chongqing Stage 2 splits and normal Guangdong month-preparation outputs, so opening a new split goes directly to the settlement month instead of inheriting a hidden/template-month view. Summary workbooks remain unchanged.
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

- 2026-07-24 full Debug solution test run passed Core 83/83, Excel 129/129, and WPF 10/10; after independent review fixes, WPF Debug and Release both passed 11/11. New tests cover Chongqing required/invalid/conflicting template decisions, actual second-candidate copying, stable output planning, `3 / 7 / 13` candidate batches, round uniqueness/reshuffle, selection retention, cross-field AND search, shared Chongqing presentation, and Enter-key focus routing.
- The inherited 2026-07-23 target-month active-sheet regressions remain included: Hainan/Chongqing generated splits and normal Guangdong month-preparation outputs reopen on the visible target month as the only active/selected sheet.
- Full Release solution build passed with 0 warnings / 0 errors.
- Build portability check passed.
- Documentation guardrails and `git diff --check` passed.
- The latest `20260724-104552-717` Win10/11 test package includes the target-month active-sheet and Hainan/Chongqing template-browser work. Its directory and ZIP each contain the same 20 files, no workbook/CSV/JSON/log/PPTX/DOCX/PDF data, and the packaged WPF executable remained running through a five-second hidden startup smoke. The executable file version remains the current production baseline `1.2.0.0`; this timestamped package is not a formal release.
- The formal `v1.2.0` package repeats those checks: directory and ZIP each contain the same 20 files, no business-data file types, executable file version `1.2.0.0`, and a successful five-second hidden startup smoke.
- One user-authorized generated Hainan summary workbook was inspected read-only for the reported loan-balance row. The later exact user-authorized new-agent split workbook was inspected read-only and contained one legacy comment at `6月!P6`; no surrounding business directory was scanned and the source workbook was not modified. Automated tests remained synthetic. No `outputs` material was read or produced.
- For the later Chongqing self-operated regression, the user explicitly authorized one exact historical ledger. It was inspected and replayed read-only: all seven reported rows had project developer/owner attribution but no monthly proxy/intermediary parameters; the fixed reader reports zero self-operated blockers. No surrounding directory was scanned and the workbook was not modified.

## Known Boundary

- Chongqing refund templates that contain both a standard month sheet such as `2` and a copy-style sheet such as `2 (2)` remain a documented pre-existing technical risk. The current task does not change that rule; obtain business confirmation before choosing a new policy.
- When several Chongqing refund templates are candidates, template-specific extra-deduction blocks cannot be compared before the operator chooses one; generation still uses the selected workbook and keeps the existing report-based manual review boundary for those blocks.
- Hainan/Chongqing intentionally remain all-or-nothing for formal Stage 2 output. A future non-payable partial draft would require a separate design and explicit authorization.
- Chongqing loan deduction remains template/manual-driven, so the Hainan automatic zero-clamp/completion-month behavior was intentionally not generalized there; doing so could hide a real manual over-deduction.
- 2027 support is analyzed but intentionally deferred in `docs/dev-notes/settlement-year-2027-readiness-2026-07-23.md`. Year parameterization, January cross-year inheritance, and the planned 2027 Hainan ledger-format change require later confirmation before implementation.
- The program does not infer payment routing from a multi-person payee cell; cashiers continue to handle distribution agreements outside this program.
- The Chongqing backtest scripts still need adaptation to the new required preflight signature/input fingerprint before they can be treated as runnable regression tools; their output paths are already isolated under `local-validation/backtests/`.

## Next Session

1. Resume on `codex/stage2-template-browser`; inspect the uncommitted browser/backend/doc changes and the final review result before committing.
2. Preserve the `outputs/` exclusions and obtain new explicit authorization before any push, merge, tag, package/release, or real-workbook read.
3. Treat `v1.2.0` as the current production baseline until this branch is separately tested, committed, merged, and released.
4. Do not resume Hainan/Chongqing performance work, implement 2027, add routing JSON, change the ledger format, or add partial formal summaries without a new explicit decision.

## Window Separation

- The other conversation remains dedicated to boss-report material under `outputs/boss-report-2026-07-14`.
- This conversation owns C# program development. Do not modify report materials unless the user explicitly asks this development window to synchronize a fact.
- `outputs/` is locally Git-excluded. Report PPT/Markdown/Word/PDF files and screenshots must not be staged, committed, or pushed.

## Documentation Maintenance

Use `docs/README.md` as the document router and `docs/CHANGELOG.md` for completed history. Keep this file current and concise; detailed rules belong in the current behavior/dev-note documents and detailed history belongs in git.
