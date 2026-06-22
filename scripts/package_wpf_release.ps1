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
$packageName = "HainanSettlementTool-Wpf-$Configuration-$stamp"
$packageDir = Join-Path $OutputRoot $packageName
$zipPath = Join-Path $OutputRoot "$packageName.zip"
$publishSource = Join-Path $ProjectRoot "src\HainanSettlementTool.Wpf\bin\$Configuration\net472"

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
    "海南售电结算自动化工具 - 现代界面测试版",
    "",
    "运行方式：",
    "1. 双击 `"海南售电结算自动化工具现代版.exe`"。",
    "2. 请保持本文件夹内的 dll 和 config 文件完整，不要只复制 exe。",
    "3. 运行前请关闭正在打开的 Excel 台账或结算结果文件。",
    "",
    "说明：",
    "- 本版本使用 WPF 现代界面壳，结算逻辑复用已验收的 Core/Excel 项目。",
    "- 兜底版 WinForms 程序仍由 scripts/package_release.ps1 打包。"
)
$readme | Set-Content -LiteralPath (Join-Path $packageDir "README.txt") -Encoding UTF8

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $packageDir "*") -DestinationPath $zipPath -Force

Write-Host "Package directory: $packageDir"
Write-Host "Zip package: $zipPath"
