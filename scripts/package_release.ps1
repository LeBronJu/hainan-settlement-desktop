param(
    [string]$Configuration = "Release",
    [string]$OutputRoot = ""
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Solution = Join-Path $ProjectRoot "HainanSettlementTool.sln"
$MSBuild = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"

if (-not (Test-Path -LiteralPath $MSBuild)) {
    throw "MSBuild not found: $MSBuild"
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $ProjectRoot "dist"
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$packageName = "HainanSettlementTool-Win7-8-$Configuration-$stamp"
$packageDir = Join-Path $OutputRoot $packageName
$zipPath = Join-Path $OutputRoot "$packageName.zip"
$publishSource = Join-Path $ProjectRoot "src\HainanSettlementTool.WinForms\bin\$Configuration\net472"

& $MSBuild $Solution /restore /p:Configuration=$Configuration /m
if ($LASTEXITCODE -ne 0) {
    throw "MSBuild failed. Exit code: $LASTEXITCODE"
}

if (Test-Path -LiteralPath $packageDir) {
    Remove-Item -LiteralPath $packageDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $packageDir | Out-Null

$include = @("*.exe", "*.config", "*.dll")
foreach ($pattern in $include) {
    Get-ChildItem -LiteralPath $publishSource -Filter $pattern -File |
        Copy-Item -Destination $packageDir -Force
}

$readme = @(
    "Hainan Settlement Tool - Win7/8 test build",
    "",
    "How to run:",
    "1. Make sure .NET Framework 4.7.2 or newer is installed.",
    "2. Double-click the exe file.",
    "3. Keep all dll and config files in this folder. Do not copy the exe alone.",
    "",
    "Safety:",
    "- Do not commit real ledgers, customer data, settlement outputs, screenshots, or finance data.",
    "- Close Excel workbooks before running generation."
)
$readme | Set-Content -LiteralPath (Join-Path $packageDir "README.txt") -Encoding UTF8

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $packageDir "*") -DestinationPath $zipPath -Force

Write-Host "Package directory: $packageDir"
Write-Host "Zip package: $zipPath"
