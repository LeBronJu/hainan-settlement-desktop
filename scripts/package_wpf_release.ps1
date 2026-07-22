param(
    [ValidateSet("Debug", "Release")]
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

if ([string]::IsNullOrWhiteSpace($ReleaseTag)) {
    if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
        $OutputRoot = Join-Path $ProjectRoot "dist\test-packages"
    }

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss-fff"
    $packageName = "RetailPowerSettlementTool-Win10-11-$Configuration-$stamp"
    $packageLabel = "test build"
}
else {
    if ($ReleaseTag -notmatch '^v\d+\.\d+\.\d+$') {
        throw "ReleaseTag must use semantic version tag format such as v1.1.0."
    }

    if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
        $OutputRoot = Join-Path $ProjectRoot "dist\releases"
    }

    $packageName = "RetailPowerSettlementTool-Win10-11-$ReleaseTag"
    $packageLabel = $ReleaseTag
}

$packageDir = Join-Path $OutputRoot $packageName
$zipPath = Join-Path $OutputRoot "$packageName.zip"
$publishSource = Join-Path $ProjectRoot "src\HainanSettlementTool.Wpf\bin\$Configuration\net472"
$executableName = "清能电力-结算自动化工具Win10-11版.exe"

if ((Test-Path -LiteralPath $OutputRoot) -and -not (Test-Path -LiteralPath $OutputRoot -PathType Container)) {
    throw "Package output root is not a directory: $OutputRoot"
}

New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null

if (Test-Path -LiteralPath $packageDir) {
    throw "Package directory already exists; refusing to overwrite it: $packageDir"
}
if (Test-Path -LiteralPath $zipPath) {
    throw "Zip package already exists; refusing to overwrite it: $zipPath"
}

Invoke-SolutionBuild

if (Test-Path -LiteralPath $packageDir) {
    throw "Package directory appeared during the build; refusing to overwrite it: $packageDir"
}
if (Test-Path -LiteralPath $zipPath) {
    throw "Zip package appeared during the build; refusing to overwrite it: $zipPath"
}

$requiredFiles = @(
    $executableName,
    "$executableName.config",
    "HainanSettlementTool.Core.dll",
    "HainanSettlementTool.Excel.dll"
)
foreach ($fileName in $requiredFiles) {
    $sourcePath = Join-Path $publishSource $fileName
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required WPF package file is missing: $sourcePath"
    }
}

$sourceDlls = @(Get-ChildItem -LiteralPath $publishSource -Filter "*.dll" -File | Sort-Object Name)
if ($sourceDlls.Count -eq 0) {
    throw "No WPF package DLL files were found: $publishSource"
}

New-Item -ItemType Directory -Path $packageDir | Out-Null

Copy-Item -LiteralPath (Join-Path $publishSource $executableName) -Destination $packageDir
Copy-Item -LiteralPath (Join-Path $publishSource "$executableName.config") -Destination $packageDir
$sourceDlls | Copy-Item -Destination $packageDir -Force

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

$expectedFiles = @(
    @($executableName, "$executableName.config", "README.txt")
    @($sourceDlls | ForEach-Object { $_.Name })
) | Sort-Object -Unique
$actualFiles = @(Get-ChildItem -LiteralPath $packageDir -File | ForEach-Object { $_.Name } | Sort-Object -Unique)
$missingPackageFiles = @($expectedFiles | Where-Object { $_ -notin $actualFiles })
$unexpectedPackageFiles = @($actualFiles | Where-Object { $_ -notin $expectedFiles })
if ($missingPackageFiles.Count -gt 0 -or $unexpectedPackageFiles.Count -gt 0) {
    throw "Package directory content validation failed. Missing: $($missingPackageFiles -join ', '); unexpected: $($unexpectedPackageFiles -join ', ')."
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory(
    $packageDir,
    $zipPath,
    [System.IO.Compression.CompressionLevel]::Optimal,
    $false)

$packageFileHashes = @{}
foreach ($file in @(Get-ChildItem -LiteralPath $packageDir -File)) {
    $packageFileHashes[$file.Name] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
}

$archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try {
    $zipFiles = @()
    foreach ($entry in @($archive.Entries | Where-Object { -not [string]::IsNullOrEmpty($_.Name) })) {
        $entryName = $entry.FullName.Replace("/", "\")
        $zipFiles += $entryName

        $entryStream = $entry.Open()
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        try {
            $entryHashBytes = $sha256.ComputeHash($entryStream)
            $entryHash = [System.BitConverter]::ToString($entryHashBytes).Replace("-", "")
        }
        finally {
            $sha256.Dispose()
            $entryStream.Dispose()
        }

        if (-not $packageFileHashes.ContainsKey($entryName) -or $packageFileHashes[$entryName] -ne $entryHash) {
            throw "Zip package entry content does not match the package directory: $entryName"
        }
    }
}
finally {
    $archive.Dispose()
}

$expectedZipFiles = @($expectedFiles | Sort-Object -Unique)
$uniqueZipFiles = @($zipFiles | Sort-Object -Unique)
$missingZipFiles = @($expectedZipFiles | Where-Object { $_ -notin $uniqueZipFiles })
$unexpectedZipFiles = @($uniqueZipFiles | Where-Object { $_ -notin $expectedZipFiles })
if ($missingZipFiles.Count -gt 0 -or $unexpectedZipFiles.Count -gt 0 -or $zipFiles.Count -ne $uniqueZipFiles.Count) {
    throw "Zip package content validation failed. Missing: $($missingZipFiles -join ', '); unexpected: $($unexpectedZipFiles -join ', ')."
}

$zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash

Write-Host "Package directory: $packageDir"
Write-Host "Zip package: $zipPath"
Write-Host "Zip SHA256: $zipHash"
