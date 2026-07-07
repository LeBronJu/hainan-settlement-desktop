# Win7/8 Maintenance Mode

Date: 2026-06-29

Status: current policy note. Use this file to decide whether WinForms should receive a change; current document ownership is listed in `docs/README.md`.

## Decision

Win10/11 WPF is now the primary UI entry for new features and UX improvements.

Win7/8 WinForms remains available as a maintenance compatibility entry:

- keep it buildable;
- keep it packageable while Win7/8 packages are still released;
- fix blocking bugs;
- adapt it only when shared Core/Excel changes require compatibility work;
- do not add new WinForms-only features or UX improvements unless the user explicitly asks.

## Why

Win7/8 users are now rare, but the WinForms entry still provides compatibility for a small user group and is already part of the public release package set.

Maintaining feature parity across WinForms and WPF is now more expensive than its value. The project should reduce duplicate UI workflow work while preserving workbook correctness and safety in shared Core/Excel modules.

## Practical Rules

- New UI behavior defaults to WPF.
- Settlement correctness, workbook safety, and report behavior still belong in Core/Excel whenever possible.
- WinForms changes are appropriate for build failures, packaging failures, blocking user bugs, or explicit Win7/8 support requests.
- Release packages may continue to include Win7/8 as a maintenance package; release notes should avoid implying it is the primary experience.

## Not Decided

- Win7/8 package removal is not scheduled.
- The WinForms project is not deleted.
- Existing WinForms workflows are not rewritten only for parity with WPF.
