param(
    [string]$Root = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}

$patterns = @(
    [pscustomobject]@{
        Name = "Hard-coded Visual Studio MSBuild.exe path"
        Regex = "(?i)(C:|%ProgramFiles(?:\(x86\))?%|\`$env:ProgramFiles(?:\(x86\))?).*Microsoft Visual Studio.*MSBuild.*MSBuild\.exe"
    },
    [pscustomobject]@{
        Name = "Hard-coded Visual Studio version directory"
        Regex = "(?i)Microsoft Visual Studio\\(?:20\d{2}|1\d)\\"
    },
    [pscustomobject]@{
        Name = "Hard-coded MSBuild Current Bin path"
        Regex = "(?i)MSBuild\\Current\\Bin\\(?:amd64\\)?MSBuild\.exe"
    },
    [pscustomobject]@{
        Name = "Hard-coded BuildTools install path"
        Regex = "(?i)Program Files.*Microsoft Visual Studio.*BuildTools"
    }
)

$previousLocation = Get-Location
Set-Location -LiteralPath $Root

try {
    $files = & git ls-files
    if ($LASTEXITCODE -ne 0) {
        throw "git ls-files failed. Exit code: $LASTEXITCODE"
    }

    $scriptRelativePath = "scripts/check_build_portability.ps1"
    $violations = New-Object System.Collections.Generic.List[object]

    foreach ($file in $files) {
        if ($file -eq $scriptRelativePath) {
            continue
        }

        $path = Join-Path $Root $file
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            continue
        }

        $lineNumber = 0
        foreach ($line in Get-Content -LiteralPath $path -Encoding UTF8) {
            $lineNumber++
            foreach ($pattern in $patterns) {
                if ($line -match $pattern.Regex) {
                    $violations.Add([pscustomobject]@{
                        File = $file
                        Line = $lineNumber
                        Rule = $pattern.Name
                        Text = $line.Trim()
                    }) | Out-Null
                }
            }
        }
    }

    if ($violations.Count -gt 0) {
        Write-Host "Build portability check failed. Replace hard-coded VS/MSBuild paths with dotnet msbuild or vswhere discovery." -ForegroundColor Red
        $violations | Format-Table File, Line, Rule, Text -AutoSize
        exit 1
    }

    Write-Host "Build portability check passed."
}
finally {
    Set-Location -LiteralPath $previousLocation
}
