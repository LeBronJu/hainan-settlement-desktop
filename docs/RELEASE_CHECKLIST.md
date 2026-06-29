# Release Checklist

Use this checklist for formal tags and GitHub Releases. It is not required for local smoke builds or temporary machine setup.

## Before Release

- Confirm `git status --short --branch`.
- Release from `main` after approved changes have been merged.
- Do not include real ledgers, settlement outputs, customer data, screenshots, or finance/payment data in commits or assets.
- Confirm the release version and tag name, for example `v1.0.1`.

## Validation

Run these checks unless the user explicitly narrows the release scope:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m
.\scripts\check_build_portability.ps1
git diff --check
.\scripts\package_release.ps1
.\scripts\package_wpf_release.ps1
```

## Documentation Impact Gate

Make a documentation impact judgment before tagging:

- Release version, assets, package names, or release status changed: update `README.md` and `HANDOFF.md`.
- Packaging or release procedure changed: update this checklist.
- Architecture boundaries changed: update `docs/architecture.md` and, if needed, a dated dev-note.
- Business rules changed: review and update `CONTEXT.md`.
- No docs affected: state the reason in the final response.

Do not update unrelated documents just to satisfy the gate.

## Assets And Publish

- GitHub Release assets use stable ASCII filenames:
  - `HainanSettlementTool-Win7-8-<tag>.zip`
  - `HainanSettlementTool-Win10-11-<tag>.zip`
- Push `main` and the tag.
- Create or update the GitHub Release with the two zip assets.
- Final response must include:
  - Documentation impact judgment.
  - Validation commands run.
  - Work intentionally not done.
