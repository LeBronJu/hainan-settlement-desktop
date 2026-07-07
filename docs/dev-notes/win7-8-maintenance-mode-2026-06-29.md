# Win7/8 Frozen Compatibility Policy

Date: 2026-06-29

Status: current policy note. Use this file to decide whether WinForms should receive work; current document ownership is listed in `docs/README.md`.

## Decision

As of 2026-07-07, Win7/8 WinForms support is frozen and no longer maintained as a release target.

Win10/11 WPF is the primary and only actively maintained desktop UI entry for new features, UX work, packaging, and user testing.

Win7/8 WinForms remains in the repository only as a historical compatibility entry:

- do not add new WinForms features;
- do not implement WinForms parity for WPF work;
- do not spend refactor or UX effort on WinForms;
- do not create new Win7/8 release packages by default;
- shared Core/Excel settlement fixes may still benefit WinForms incidentally;
- reopen Win7/8 work only if the user explicitly reverses this policy.

## Why

The user decided on 2026-07-07 to end Win7/8 maintenance instead of continuing to carry duplicate UI and packaging work.

Keeping the project focused on Win10/11 WPF reduces UI drift and lets workbook correctness work stay in the shared Core/Excel layers.

## Practical Rules

- New UI behavior goes to WPF only.
- Formal and test packages default to Win10/11 WPF only.
- WinForms build failures are not automatically release blockers unless they break shared solution validation selected for the task.
- If shared Core/Excel changes affect WinForms incidentally, do not add WinForms-specific workaround code unless the user explicitly asks.

## Not Deleted

- The WinForms project is not deleted in this policy change.
- Existing historical releases and historical Win7/8 packages remain historical artifacts.
- `scripts/package_release.ps1` may remain in the repository as a historical compatibility script, but it is not part of the default future release path.
