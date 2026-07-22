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
- Preserve the Window Separation / `outputs` exclusions in this handoff and do not revert them.
- Do not push, merge, tag, package a formal version, or publish without explicit user authorization.

## Release State

- Latest formal release remains `v1.1.0`.
- Release: `https://github.com/LeBronJu/retail-power-settlement-desktop/releases/tag/v1.1.0`.
- The Guangdong safety hardening package passed user practical testing, but the user explicitly decided not to publish a new version yet.
- The latest local Win10/11 test package for user retesting is `dist/test-packages/RetailPowerSettlementTool-Win10-11-Release-20260722-161252-785.zip`; SHA-256 is `CAAED0557E8A34A566F279B56D54F2759D017020C00A4DD0B7B101066C09134F`. It supersedes the earlier `155344-666` package and includes both the template-choice display fix and the Hainan loan fixes.
- This is not a formal release: no release tag, GitHub Release, push, merge, or publication was performed.
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
- Hainan loan remaining is written as `MAX(0, loan - deducted)`; when the program-owned current-month deduction first clears a positive balance and the completion month is blank, the concrete target settlement month is written as a date value and displayed as `yyyy年m月` (for example, a June 2026 run writes `2026年6月`). Existing completion months are preserved and historical missing months are not guessed.
- Analyze is read-only. A plan freezes preflight signature and file/output-state fingerprint; changes after confirmation force a new preflight.
- All workbooks and reports are generated in one staging workspace, strongly verified, then published as one batch. Failed staging is retained under `【未完成-禁止付款】`.
- There is no payment-routing JSON and no partial formal summary generation for Hainan/Chongqing.

Guangdong:

- Proxy/intermediary/refund month-sheet preparation and its safety hardening passed user practical testing.
- Full Guangdong settlement remains paused.
- Guangdong's independent-workbook partial-completion behavior must not be generalized to Hainan/Chongqing formal settlement.
- A read-only design review recommends extending Guangdong's self-contained readable HTML through a shared escaped HTML presentation renderer while keeping each workflow's status/business mapping and publish semantics province-specific. This has not been implemented yet.

Performance / UI support:

- Hainan and Chongqing performance optimization is paused because monthly volume is small. Do not resume it unless the user explicitly reopens it.
- Win7/8 WinForms remains frozen. Win10/11 WPF is the maintained entry.

## Latest Validation

Current branch, synthetic inputs only:

- Core tests: 80/80 passed.
- Excel tests: 125/125 passed.
- WPF presentation/control tests: 3/3 passed.
- Focused Hainan Stage 2 tests: 57/57 passed.
- Focused Chongqing Stage 2 tests: 39/39 passed.
- Full Debug and Release solution builds passed.
- WPF Debug build passed with 0 warnings / 0 errors.
- Build portability check passed.
- Documentation guardrails and `git diff --check` passed.
- Independent read-only diff review found no remaining P0/P1 in the current scope; its two P2 findings (card-count wording and whitespace-normalized subject keys) were addressed.
- The latest Win10/11 test package build passed; its directory and readable ZIP each contain the same 20 files, no workbook/CSV/JSON/log data, and the packaged WPF executable remained running through a five-second hidden startup smoke.
- One user-authorized generated Hainan summary workbook was inspected read-only for the reported loan-balance row; automated tests remained synthetic. No other real workbook/business directory or `outputs` material was read or produced, and no formal release was created.

## Known Boundary

- Chongqing refund templates that contain both a standard month sheet such as `2` and a copy-style sheet such as `2 (2)` remain a documented pre-existing technical risk. The current task does not change that rule; obtain business confirmation before choosing a new policy.
- Hainan/Chongqing intentionally remain all-or-nothing for formal Stage 2 output. A future non-payable partial draft would require a separate design and explicit authorization.
- Chongqing loan deduction remains template/manual-driven, so the Hainan automatic zero-clamp/completion-month behavior was intentionally not generalized there; doing so could hide a real manual over-deduction.
- The program does not infer payment routing from a multi-person payee cell; cashiers continue to handle distribution agreements outside this program.
- The Chongqing backtest scripts still need adaptation to the new required preflight signature/input fingerprint before they can be treated as runnable regression tools; their output paths are already isolated under `local-validation/backtests/`.

## Next Session

1. Run the required reading gate and continue on `codex/stage2-preflight-integrity`; do not restart the implementation from `main`.
2. Inspect the working tree and the current task note before changing anything. Preserve unrelated user edits and `outputs` exclusions.
3. If the user wants practical workbook validation, obtain explicit current authorization for the exact read-only source and write only to a separate temporary output. Otherwise keep using synthetic fixtures.
4. Use the latest `161252-785` test package for the next practical check. Confirm that the selected split-template option keeps its friendly label, then replay the June Hainan output and verify a just-cleared loan shows remaining 0 and fills `2026年6月` as the completion month. Do not create a formal package, commit to `main`, push, merge, tag, or release unless the user explicitly asks.
5. Do not resume Hainan/Chongqing performance work, add routing JSON, change the ledger format, or add partial formal summaries without a new explicit decision.

## Window Separation

- The other conversation remains dedicated to boss-report material under `outputs/boss-report-2026-07-14`.
- This conversation owns C# program development. Do not modify report materials unless the user explicitly asks this development window to synchronize a fact.
- `outputs/` is locally Git-excluded. Report PPT/Markdown/Word/PDF files and screenshots must not be staged, committed, or pushed.

## Documentation Maintenance

Use `docs/README.md` as the document router and `docs/CHANGELOG.md` for completed history. Keep this file current and concise; detailed rules belong in the current behavior/dev-note documents and detailed history belongs in git.
