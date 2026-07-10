param(
    [string]$Configuration = "Release",
    [string]$OutputRoot = "",
    [string]$ReleaseTag = ""
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Solution = Join-Path $ProjectRoot "HainanSettlementTool.sln"

function Get-DotNetPath {
    $command = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($command -and $command.Source) {
        return $command.Source
    }

    $candidates = @()
    if ($env:ProgramFiles) {
        $candidates += Join-Path $env:ProgramFiles "dotnet\dotnet.exe"
    }
    if (${env:ProgramFiles(x86)}) {
        $candidates += Join-Path ${env:ProgramFiles(x86)} "dotnet\dotnet.exe"
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Get-MSBuildPath {
    $vswhereCandidates = @()
    if (${env:ProgramFiles(x86)}) {
        $vswhereCandidates += Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    }
    if ($env:ProgramFiles) {
        $vswhereCandidates += Join-Path $env:ProgramFiles "Microsoft Visual Studio\Installer\vswhere.exe"
    }

    foreach ($vswhere in $vswhereCandidates) {
        if (Test-Path -LiteralPath $vswhere) {
            $found = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
            if (-not [string]::IsNullOrWhiteSpace($found)) {
                return $found
            }
        }
    }

    $command = Get-Command msbuild -ErrorAction SilentlyContinue
    if ($command -and $command.Source) {
        return $command.Source
    }

    throw "Neither dotnet nor MSBuild could be found."
}

function Invoke-SolutionBuild {
    $dotnet = Get-DotNetPath
    if ($dotnet) {
        Write-Host "Building with dotnet msbuild: $dotnet"
        & $dotnet msbuild $Solution /restore /p:Configuration=$Configuration /m
        if ($LASTEXITCODE -ne 0) {
            throw "dotnet msbuild failed. Exit code: $LASTEXITCODE"
        }
        return
    }

    $msbuild = Get-MSBuildPath
    Write-Host "Building with MSBuild: $msbuild"
    & $msbuild $Solution /restore /p:Configuration=$Configuration /m
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed. Exit code: $LASTEXITCODE"
    }
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $ProjectRoot "dist"
}

if ([string]::IsNullOrWhiteSpace($ReleaseTag)) {
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $packageName = "RetailPowerSettlementTool-Win10-11-$Configuration-$stamp"
    $packageLabel = "test build"
}
else {
    if ($ReleaseTag -notmatch '^v\d+\.\d+\.\d+$') {
        throw "ReleaseTag must use semantic version tag format such as v1.1.0."
    }

    $packageName = "RetailPowerSettlementTool-Win10-11-$ReleaseTag"
    $packageLabel = $ReleaseTag
}

$packageDir = Join-Path $OutputRoot $packageName
$zipPath = Join-Path $OutputRoot "$packageName.zip"
$publishSource = Join-Path $ProjectRoot "src\HainanSettlementTool.Wpf\bin\$Configuration\net472"
$executableName = "清能电力-结算自动化工具Win10-11版.exe"

Invoke-SolutionBuild

if (Test-Path -LiteralPath $packageDir) {
    Remove-Item -LiteralPath $packageDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $packageDir | Out-Null

$requiredFiles = @($executableName, "$executableName.config")
foreach ($fileName in $requiredFiles) {
    $sourcePath = Join-Path $publishSource $fileName
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required WPF package file is missing: $sourcePath"
    }

    Copy-Item -LiteralPath $sourcePath -Destination $packageDir -Force
}

Get-ChildItem -LiteralPath $publishSource -Filter "*.dll" -File |
    Copy-Item -Destination $packageDir -Force

$readme = @(
    "Retail Power Settlement Tool - Win10/11 $packageLabel",
    "",
    "How to run:",
    "1. Double-click the Win10/11 exe file.",
    "2. Keep all dll and config files in this folder. Do not copy the exe alone.",
    "3. Do not edit source workbooks while generation is running; opened workbooks are read with shared access where supported.",
    "",
    "Notes:",
    "- This build uses the Win10/11 WPF shell and the verified Core/Excel settlement logic.",
    "- The Win7/8 build is frozen as a historical compatibility entry and is no longer packaged by default."
)
$readme | Set-Content -LiteralPath (Join-Path $packageDir "README.txt") -Encoding UTF8

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $packageDir "*") -DestinationPath $zipPath -Force

Write-Host "Package directory: $packageDir"
Write-Host "Zip package: $zipPath"
