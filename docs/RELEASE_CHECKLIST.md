# Release Checklist

Use this checklist for formal tags and GitHub Releases. It is not required for local smoke builds or temporary machine setup.

## Before Release

- Confirm `git status --short --branch`.
- Release from `main` after approved changes have been merged.
- Do not include real ledgers, settlement outputs, customer data, screenshots, or finance/payment data in commits or assets.
- Confirm the release version and semantic-version tag name, for example `v1.1.0`.

## Validation

Run these checks unless the user explicitly narrows the release scope:

```powershell
dotnet test .\HainanSettlementTool.sln /p:Configuration=Debug
dotnet msbuild .\HainanSettlementTool.sln /restore /p:Configuration=Release /m
.\scripts\check_build_portability.ps1
.\scripts\check_docs_guardrails.ps1
git diff --check
.\scripts\package_wpf_release.ps1 -ReleaseTag <tag>
```

Win7/8 WinForms is frozen as a historical compatibility entry. Do not run `.\scripts\package_release.ps1` for default releases unless the user explicitly reopens Win7/8 support.

`package_wpf_release.ps1` writes timestamped local test packages to `dist/test-packages/` when `-ReleaseTag` is omitted, and stable versioned formal packages to `dist/releases/` when `-ReleaseTag` is supplied. An explicit `-OutputRoot` is used exactly as supplied. The script refuses to overwrite an existing package directory or zip, verifies the executable, config, Core/Excel assemblies and complete DLL set in both the directory and readable zip, and prints the zip SHA-256 hash. The frozen Win7/8 helper defaults to `dist/legacy-win7-8/`.

## Documentation Impact Gate

Make a documentation impact judgment before merging release-bound changes or tagging:

- Release version, assets, package names, or release status changed: update `README.md` and `HANDOFF.md`.
- Packaging or release procedure changed: update this checklist.
- Architecture boundaries changed: update `docs/architecture.md` and, if needed, a dated dev-note.
- Business rules changed: review and update `CONTEXT.md`.
- Documentation ownership, canonical links, or new current-behavior docs changed: update `docs/README.md`.
- Completed milestone history changed in a way future agents should discover: update `docs/CHANGELOG.md`.
- Branch state, validation results, or next steps changed: update `HANDOFF.md`.
- No docs affected: state the reason in the final response.

Do not update unrelated documents just to satisfy the gate.

## Validation Gate

- Run relevant tests, builds, and packaging checks, or explain why they are not needed.
- Confirm no real ledgers, customer data, settlement outputs, screenshots, or sensitive finance data are included.
- Confirm generated packages contain all required `.dll`, `.config`, and executable files.
- Record the SHA-256 printed by the packaging script and confirm the formal package was created under `dist/releases/` (or the explicitly supplied output root) without replacing an existing artifact.

## Assets And Publish

- GitHub Release assets use stable ASCII filenames:
  - `RetailPowerSettlementTool-Win10-11-<tag>.zip`
- Confirm the packaged executable file/product version matches the release version and the package `README.txt` names the release tag.
- Push `main` and the tag.
- Create or update the GitHub Release with the Win10/11 WPF zip asset.
- Final response must include:
  - Documentation impact judgment.
  - Validation commands run.
  - Work intentionally not done.
