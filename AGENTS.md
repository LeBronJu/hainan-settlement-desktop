# Agent Instructions

This repository is the standalone C# desktop project for multi-province retail electricity settlement automation.
Hainan is the mature first province module. Chongqing Stage 1 power cleaning and ledger update are implemented; Chongqing Stage 2 is under analysis/implementation.

## Safety Rules

- Never commit real ledgers, settlement outputs, customer names, screenshots with sensitive data, or finance/payment data.
- Do not overwrite user workbooks in place. Generated files must be written to an output folder or a clearly named copy.
- Preserve user edits. If local files are dirty, inspect them and work with them.
- Prefer explicit validation reports and user-visible logs over silent guesses.
- If a business rule is unclear and cannot be inferred from existing documentation, ask the user before encoding it.

## Project Scope

This repository contains only the C# rewrite. The earlier Python project remains in the original repository and remains the full historical behavior reference:

- `https://github.com/LeBronJu/hainan-settlement-tool`

The C# version is being built as a maintainable Windows desktop app. It should eventually replace or reduce reliance on the Python app, but only after staged validation.

## Repository Layout

- `HainanSettlementTool.sln`: solution file.
- `src/HainanSettlementTool.WinForms/`: Win7/8 maintenance UI for the existing Hainan workflow. Keep it buildable and fix blocking bugs, but do not add new province features or UX improvements unless the user explicitly asks.
- `src/HainanSettlementTool.Wpf/`: Win10/11 WPF main UI. New province entries, new UI features, and UX improvements belong here by default.
- `src/HainanSettlementTool.Core/`: business models, services, and interfaces.
- `src/HainanSettlementTool.Excel/`: ClosedXML workbook reading/writing.
- `docs/README.md`: documentation map and current source-of-truth index.
- `docs/architecture.md`: layering and migration boundary.
- `docs/hainan-stage2-current-behavior.md`: current Hainan Stage 2 behavior and validation summary.
- `CONTEXT.md`: domain vocabulary and settlement rules.
- `docs/dev-notes/`: architecture reviews, robustness priorities, and one-off technical notes.
- `HANDOFF.md`: current state for future sessions.

## Agent skills

### Issue tracker

Issues and PRDs are tracked in GitHub Issues for `LeBronJu/hainan-settlement-desktop`. See `docs/agents/issue-tracker.md`.

### Triage labels

Use the default five-role triage label vocabulary. See `docs/agents/triage-labels.md`.

### Domain docs

This is a single-context repo. See `docs/agents/domain.md`.

## Engineering Rules

- Start every project work session by checking the branch and reading current project instructions before editing. Minimum gate: run `git status --short --branch`, then read `AGENTS.md` and `HANDOFF.md`. Do not rely only on chat history, memory, or a prior agent summary.
- Before changing an area, read that area's owning document: business settlement rules require `CONTEXT.md`; module boundaries or workflow seams require `docs/architecture.md`; release/packaging requires `docs/RELEASE_CHECKLIST.md`; user-visible setup or package status requires `README.md`; current branch/task state requires `HANDOFF.md`. If ownership is unclear, read `docs/README.md` first.
- Before new-province onboarding, WPF province UI, Core multi-province workflow, or Excel multi-province adapter work, also read `docs/dev-notes/multi-province-readiness-2026-07-07.md` and use its P0/P1/P2 readiness order.
- If context was compacted, the thread was resumed after a pause, or the task direction changed, repeat the relevant reading gate before making further edits.
- Do not make development changes directly on `main` or `master`. Create a development branch first, using the `codex/` prefix unless the user requests another branch name.
- This is a local single-developer project. Pull requests are optional; it is acceptable to commit and push a `codex/` branch, then merge locally to `main` when the user authorizes it. Still run the documented validation and documentation-impact checks before merging or releasing.
- This project allows Codex subagents, spawning, parallel exploration, and other efficiency tools when they materially speed up safe work. Do not repeatedly ask for permission before routine low-risk use. Pause and warn or ask only before high-risk operations such as reading real business files outside an authorized scope, modifying production/user workbooks, destructive git commands, merging to `main`, tagging/releasing, deleting/moving large file trees, or any action that could affect settlement correctness or sensitive data.
- If an issue is uncertain, ambiguous, or risky, especially when it may affect settlement correctness, workbook safety, or user-visible business rules, stop and analyze it explicitly for the user. Do not encode a guess; ask the user to decide.
- UI must not contain Excel parsing, matching, amount calculation, or workbook template rules.
- Core must not reference ClosedXML, WinForms, WPF, or file-format implementation details.
- Excel layer owns workbook reading/writing and template copying.
- Keep stage boundaries explicit.
- Win7/8 WinForms is in maintenance mode. Do not spend quality/refactor work on WinForms parity unless needed for compilation, packaging, blocking bugfixes, shared Core/Excel behavior changes, or explicit user authorization.
- WPF confirmation, warning, and error dialogs must use project-native modern WPF windows/styles. Do not add system `MessageBox` prompts to new WPF flows; keep OS-native dialogs only for file/folder pickers where appropriate.
- Keep documentation current without creating noise. Each code, config, script, packaging, release, workflow, architecture, business-rule, UI-behavior, test-process, or task-state change must end with a documentation impact judgment.
- Final responses for development work must explicitly include documentation impact judgment, validation performed, and work intentionally not done when applicable. This is required even when no documentation was updated. Missing the documentation impact judgment means the task is not complete.
- Update only documents whose responsibility is affected. User-visible behavior usually affects `README.md` and `HANDOFF.md`; business rules affect `CONTEXT.md`; module boundaries affect `AGENTS.md` plus an ADR or dated dev-note; release and packaging changes affect `README.md`, `HANDOFF.md`, and `docs/RELEASE_CHECKLIST.md`; branch state, validation results, or next steps affect `HANDOFF.md`.
- Temporary local setup or new-machine exploration that does not enter the project mainline can state that no project docs were needed.
- For temporary investigations or one-off architecture notes, add or update a dated file under `docs/dev-notes/`.

## Current Functional Boundary

The app is evolving from a Hainan-only desktop tool into a multi-province settlement automation tool. Keep province-specific business rules isolated behind province/module naming. Do not add broad `if Hainan / if Chongqing` branches in shared logic when a province-specific service or Excel generator is the cleaner boundary.

Current business scope and stage rules live in `CONTEXT.md`. Current architecture and module boundaries live in `docs/architecture.md`. Hainan Stage 2 implementation details live in `docs/hainan-stage2-current-behavior.md`. Chongqing Stage 2 analysis lives in `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`.

## Business Rules To Preserve

- Hainan ledger power unit is `万千瓦时`; older Hainan wording saying `兆瓦时` was wrong. Chongqing Stage 1 cleaning and ledger update use `兆瓦时`.
- Do not rename ledger customer names only to match payment-account names in summary workbooks.
- `项目开发人` is an agent/intermediary relationship under a负责人, not the salesperson themselves.
- Historical January/February 2026 data may be irregular. Do not generalize those quirks.
- New customers can be left with blank负责人/项目开发人 for manual review in stage 1.
- Hainan Stage 2 keeps the existing new-subject invoice/payment defaults described in `docs/hainan-stage2-current-behavior.md`. Do not apply those defaults to Chongqing without an explicit Chongqing rule; Chongqing new summary subjects should use a preflight/default model designed for Chongqing.

## Build Command

```powershell
dotnet msbuild ".\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

## Build Portability Check

```powershell
.\scripts\check_build_portability.ps1
```

## UI Support Policy

- Win10/11 WPF is the primary user experience and default target for new UI features.
- Win7/8 WinForms remains available as a maintenance compatibility entry.
- Shared settlement correctness, workbook safety, and report generation fixes still belong in Core/Excel and benefit both entries.
- WinForms-only changes should be limited to blocking bugfixes, build/package compatibility, or explicitly requested support.

## Compatibility Target

- Target framework: `.NET Framework 4.7.2`
- Intended runtime: Windows 7 SP1 and later, with .NET Framework 4.7.2 or newer installed.
- Development requires .NET SDK 8 or newer, or equivalent MSBuild that can resolve SDK-style .NET Framework projects.
- Packaging scripts should prefer `dotnet msbuild`; if `dotnet` is unavailable, use `vswhere` to discover MSBuild.exe instead of hard-coding a Visual Studio version path.
