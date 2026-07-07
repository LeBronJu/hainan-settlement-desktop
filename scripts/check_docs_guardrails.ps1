param(
    [string]$Root = "",
    [int]$MaxHandoffLines = 250
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}

function Add-Violation {
    param(
        [System.Collections.Generic.List[object]]$Violations,
        [string]$File,
        [string]$Rule,
        [string]$Detail
    )

    $Violations.Add([pscustomobject]@{
        File = $File
        Rule = $Rule
        Detail = $Detail
    }) | Out-Null
}

function Get-RelativePath {
    param([string]$Path)

    $rootWithSlash = $Root.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    if ($Path.StartsWith($rootWithSlash, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $Path.Substring($rootWithSlash.Length)
    }

    return $Path
}

$previousLocation = Get-Location
Set-Location -LiteralPath $Root

try {
    $violations = New-Object System.Collections.Generic.List[object]

    $handoffPath = Join-Path $Root "HANDOFF.md"
    if (-not (Test-Path -LiteralPath $handoffPath -PathType Leaf)) {
        Add-Violation $violations "HANDOFF.md" "Required file" "HANDOFF.md is missing."
    }
    else {
        $handoffLines = @(Get-Content -LiteralPath $handoffPath -Encoding UTF8)
        if ($handoffLines.Count -gt $MaxHandoffLines) {
            Add-Violation $violations "HANDOFF.md" "Size limit" "HANDOFF.md has $($handoffLines.Count) lines; max is $MaxHandoffLines. Move history to docs/CHANGELOG.md or docs/dev-notes/."
        }

        $forbiddenHandoffHeadings = @(
            "## Documentation Map",
            "## Documentation Rule",
            "## Useful Files",
            "## Release 1.0",
            "## Release 1.0.1"
        )

        foreach ($heading in $forbiddenHandoffHeadings) {
            if ($handoffLines -contains $heading) {
                Add-Violation $violations "HANDOFF.md" "No duplicate long sections" "Remove '$heading'. Keep HANDOFF.md as current handoff and point to docs/README.md or docs/CHANGELOG.md."
            }
        }

        if (-not (($handoffLines -join "`n") -match "not standing permission for new reads")) {
            Add-Violation $violations "HANDOFF.md" "Real-data authorization clarity" "Historical real-data authorizations must be marked as not standing permission for new reads."
        }
    }

    $docsReadmePath = Join-Path $Root "docs/README.md"
    if (-not (Test-Path -LiteralPath $docsReadmePath -PathType Leaf)) {
        Add-Violation $violations "docs/README.md" "Required file" "docs/README.md is missing."
    }
    else {
        $docsReadmeText = (Get-Content -LiteralPath $docsReadmePath -Encoding UTF8) -join "`n"
        $requiredPatterns = @(
            "## 当前准绳",
            "## 常见任务阅读顺序",
            "## 维护规则",
            "docs/CHANGELOG.md",
            "HANDOFF.md",
            "docs/dev-notes/documentation-sync-gate-2026-06-25.md"
        )

        foreach ($pattern in $requiredPatterns) {
            if ($docsReadmeText -notmatch [regex]::Escape($pattern)) {
                Add-Violation $violations "docs/README.md" "Document routing" "Missing required routing marker: $pattern"
            }
        }
    }

    $changelogPath = Join-Path $Root "docs/CHANGELOG.md"
    if (-not (Test-Path -LiteralPath $changelogPath -PathType Leaf)) {
        Add-Violation $violations "docs/CHANGELOG.md" "Required file" "docs/CHANGELOG.md is missing. Completed high-signal milestones need a place outside HANDOFF.md."
    }

    $currentBehaviorFiles = @(Get-ChildItem -LiteralPath (Join-Path $Root "docs") -Filter "*-current-behavior.md" -File -ErrorAction SilentlyContinue)
    foreach ($file in $currentBehaviorFiles) {
        $firstLines = @(Get-Content -LiteralPath $file.FullName -Encoding UTF8 -TotalCount 8)
        if (-not (($firstLines -join "`n") -match "状态：当前行为文档")) {
            Add-Violation $violations (Get-RelativePath $file.FullName) "Current behavior status" "current-behavior docs must declare '状态：当前行为文档' near the top."
        }
    }

    $devNotesRoot = Join-Path $Root "docs/dev-notes"
    if (Test-Path -LiteralPath $devNotesRoot -PathType Container) {
        $devNoteFiles = @(Get-ChildItem -LiteralPath $devNotesRoot -Filter "*.md" -File)
        foreach ($file in $devNoteFiles) {
            $firstLines = @(Get-Content -LiteralPath $file.FullName -Encoding UTF8 -TotalCount 8)
            if (-not (($firstLines -join "`n") -match "(?m)^(状态：|Status:)")) {
                Add-Violation $violations (Get-RelativePath $file.FullName) "Dev note status" "dev notes must declare a current/historical status near the top."
            }
        }
    }

    if ($violations.Count -gt 0) {
        Write-Host "Documentation guardrails failed." -ForegroundColor Red
        $violations | Format-Table File, Rule, Detail -AutoSize
        exit 1
    }

    Write-Host "Documentation guardrails passed."
}
finally {
    Set-Location -LiteralPath $previousLocation
}
