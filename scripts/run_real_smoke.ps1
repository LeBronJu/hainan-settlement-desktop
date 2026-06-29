param(
    [Parameter(Mandatory = $true)]
    [int]$Month,

    [string]$RawDetailPath = "",
    [string]$ExistingPowerPath = "",
    [string]$BaseLedgerPath = "",
    [string]$ReferenceLedgerPath = "",
    [string]$ReviewedLedgerPath = "",
    [string]$ProxyTemplateDirectory = "",
    [string]$IntermediaryTemplateDirectory = "",
    [string]$SummaryTemplatePath = "",
    [string]$OutputRoot = "",
    [string]$OutputLedgerName = "",
    [string]$OutputSummaryName = "",
    [string]$Configuration = "Release",

    [switch]$CopyReferenceExisting,
    [switch]$AllowMissingOwner,
    [switch]$SkipBuild,
    [switch]$SkipStage1Ledger,
    [switch]$SkipStage2
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Solution = Join-Path $ProjectRoot "HainanSettlementTool.sln"

function Assert-File {
    param(
        [string]$Path,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Label path is required."
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label file not found: $Path"
    }
}

function Assert-Directory {
    param(
        [string]$Path,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Label path is required."
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "$Label directory not found: $Path"
    }
}

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

function Add-AssemblyResolution {
    param(
        [string[]]$Directories
    )

    $handler = [System.ResolveEventHandler]{
        param($sender, $args)
        $name = (New-Object System.Reflection.AssemblyName($args.Name)).Name + ".dll"
        foreach ($directory in $Directories) {
            $candidate = Join-Path $directory $name
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return [System.Reflection.Assembly]::LoadFrom($candidate)
            }
        }

        return $null
    }

    [System.AppDomain]::CurrentDomain.add_AssemblyResolve($handler)
}

function New-Map {
    param($Rows)

    $map = @{}
    foreach ($row in $Rows) {
        $map[$row.Key] = $row
    }

    return $map
}

function Compare-PowerRows {
    param(
        $GeneratedRows,
        $ExistingRows
    )

    $generatedMap = New-Map $GeneratedRows
    $existingMap = New-Map $ExistingRows
    $missing = @($existingMap.Keys | Where-Object { -not $generatedMap.ContainsKey($_) }).Count
    $extra = @($generatedMap.Keys | Where-Object { -not $existingMap.ContainsKey($_) }).Count
    $valueDiffs = 0

    foreach ($key in $generatedMap.Keys) {
        if (-not $existingMap.ContainsKey($key)) {
            continue
        }

        $generated = $generatedMap[$key]
        $existing = $existingMap[$key]
        if ([Math]::Round($generated.Total - $existing.Total, 4) -ne 0 -or
            [Math]::Round($generated.Sharp - $existing.Sharp, 4) -ne 0 -or
            [Math]::Round($generated.Peak - $existing.Peak, 4) -ne 0 -or
            [Math]::Round($generated.Flat - $existing.Flat, 4) -ne 0 -or
            [Math]::Round($generated.Valley - $existing.Valley, 4) -ne 0) {
            $valueDiffs++
        }
    }

    return [ordered]@{
        GeneratedRows = $GeneratedRows.Count
        ExistingRows = $ExistingRows.Count
        MissingInGenerated = $missing
        ExtraInGenerated = $extra
        ValueDiffRows = $valueDiffs
    }
}

function Measure-FormulaErrorText {
    param([string]$Root)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $patterns = @("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A")
    $hits = 0
    $files = @(Get-ChildItem -LiteralPath $Root -Recurse -Filter "*.xlsx" | Where-Object { $_.Name -notlike "~$*" })

    foreach ($file in $files) {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
        try {
            foreach ($entry in $zip.Entries) {
                if ($entry.FullName -like "xl/worksheets/*" -or $entry.FullName -like "xl/sharedStrings.xml") {
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    try {
                        $text = $reader.ReadToEnd()
                    }
                    finally {
                        $reader.Dispose()
                    }

                    foreach ($pattern in $patterns) {
                        if ($text.Contains($pattern)) {
                            $hits++
                        }
                    }
                }
            }
        }
        finally {
            $zip.Dispose()
        }
    }

    return [ordered]@{
        FilesScanned = $files.Count
        ErrorTextHits = $hits
    }
}

if ($Month -le 1) {
    throw "Month must be greater than 1."
}

if ([string]::IsNullOrWhiteSpace($RawDetailPath) -and [string]::IsNullOrWhiteSpace($ExistingPowerPath)) {
    throw "Provide RawDetailPath or ExistingPowerPath."
}

if (-not $SkipStage1Ledger) {
    Assert-File $BaseLedgerPath "Base ledger"
}

if (-not $SkipStage2) {
    Assert-File $ReviewedLedgerPath "Reviewed ledger"
    Assert-Directory $ProxyTemplateDirectory "Proxy template"
    Assert-Directory $IntermediaryTemplateDirectory "Intermediary template"
    Assert-File $SummaryTemplatePath "Summary template"
}

if (-not [string]::IsNullOrWhiteSpace($RawDetailPath)) {
    Assert-File $RawDetailPath "Raw detail"
}
if (-not [string]::IsNullOrWhiteSpace($ExistingPowerPath)) {
    Assert-File $ExistingPowerPath "Existing power workbook"
}
if (-not [string]::IsNullOrWhiteSpace($ReferenceLedgerPath)) {
    Assert-File $ReferenceLedgerPath "Reference ledger"
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $ProjectRoot "dist"
}

if (-not $SkipBuild) {
    Invoke-SolutionBuild
}

$excelBin = Join-Path $ProjectRoot "src\HainanSettlementTool.Excel\bin\$Configuration\net472"
$coreBin = Join-Path $ProjectRoot "src\HainanSettlementTool.Core\bin\$Configuration\net472"
$coreDll = Join-Path $coreBin "HainanSettlementTool.Core.dll"
$excelDll = Join-Path $excelBin "HainanSettlementTool.Excel.dll"
Assert-File $coreDll "Core assembly"
Assert-File $excelDll "Excel assembly"

Add-AssemblyResolution @($excelBin, $coreBin)
[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$smokeRoot = Join-Path $OutputRoot "real-smoke-$stamp"
$stage1CleanDir = Join-Path $smokeRoot "stage1-clean"
$stage1LedgerDir = Join-Path $smokeRoot "stage1-ledger"
$stage2Dir = Join-Path $smokeRoot "stage2"
New-Item -ItemType Directory -Force -Path $stage1CleanDir | Out-Null
if (-not $SkipStage1Ledger) {
    New-Item -ItemType Directory -Force -Path $stage1LedgerDir | Out-Null
}
if (-not $SkipStage2) {
    New-Item -ItemType Directory -Force -Path $stage2Dir | Out-Null
}

$gateway = [HainanSettlementTool.Excel.ClosedXmlStage1ExcelGateway]::new()
$stage1 = [HainanSettlementTool.Core.Services.Stage1Service]::new($gateway)
$stage2 = [HainanSettlementTool.Core.Services.Stage2Service]::new($gateway)

$powerPath = $ExistingPowerPath
$cleanSummary = $null

if (-not [string]::IsNullOrWhiteSpace($RawDetailPath)) {
    $powerPath = Join-Path $stage1CleanDir "零售侧用户电量数据处理表.xlsx"
    $cleanReport = $stage1.CleanPowerData($RawDetailPath, $powerPath, { param($message) })
    $generatedRows = $gateway.ReadPowerRows($powerPath)
    $cleanSummary = [ordered]@{
        RawRows = $cleanReport.RawRows
        PowerRows = $cleanReport.PowerRows
        MonthTotal = $cleanReport.MonthTotal
        Output = $powerPath
    }

    if (-not [string]::IsNullOrWhiteSpace($ExistingPowerPath)) {
        $existingRows = $gateway.ReadPowerRows($ExistingPowerPath)
        $cleanSummary.CompareWithExisting = Compare-PowerRows $generatedRows $existingRows
    }
}

$stage1Summary = $null
if (-not $SkipStage1Ledger) {
    $stage1Options = [HainanSettlementTool.Core.Models.Stage1Options]::new()
    $stage1Options.Month = $Month
    $stage1Options.BaseLedgerPath = $BaseLedgerPath
    $stage1Options.PowerPath = $powerPath
    $stage1Options.RawDetailPath = $RawDetailPath
    $stage1Options.ReferenceLedgerPath = $ReferenceLedgerPath
    $stage1Options.OutputDirectory = $stage1LedgerDir
    $stage1Options.OutputLedgerName = $OutputLedgerName
    $stage1Options.CopyReferenceExisting = [bool]$CopyReferenceExisting

    $stage1Report = $stage1.Run($stage1Options, { param($message) })
    $stage1Summary = [ordered]@{
        Output = $stage1Report.Output
        ReportPath = $stage1Report.ReportPath
        PowerRows = $stage1Report.PowerRows
        MatchedRows = $stage1Report.MatchedRows
        NewRows = $stage1Report.NewRows
        MissingManualInfo = $stage1Report.MissingManualInfo.Count
        MissingCodes = $stage1Report.MissingCodes.Count
        DuplicateNamesInPowerFile = $stage1Report.DuplicateNamesInPowerFile.Count
    }
}

$stage2Summary = $null
if (-not $SkipStage2) {
    $stage2Options = [HainanSettlementTool.Core.Models.Stage2Options]::new()
    $stage2Options.Month = $Month
    $stage2Options.LedgerPath = $ReviewedLedgerPath
    $stage2Options.ProxyTemplateDirectory = $ProxyTemplateDirectory
    $stage2Options.IntermediaryTemplateDirectory = $IntermediaryTemplateDirectory
    $stage2Options.SummaryTemplatePath = $SummaryTemplatePath
    $stage2Options.OutputDirectory = $stage2Dir
    $stage2Options.OutputSummaryName = $OutputSummaryName
    $stage2Options.AllowMissingOwner = [bool]$AllowMissingOwner

    $preflight = $stage2.Analyze($stage2Options)
    $stage2Report = $stage2.Run($stage2Options, { param($message) })
    $proxyXlsx = @(Get-ChildItem -LiteralPath (Join-Path $stage2Dir "2026年代理 - 海南") -Recurse -Filter "*.xlsx" -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "~$*" })
    $intermediaryXlsx = @(Get-ChildItem -LiteralPath (Join-Path $stage2Dir "2026年居间 - 海南") -Recurse -Filter "*.xlsx" -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "~$*" })
    $stage2Xlsx = @(Get-ChildItem -LiteralPath $stage2Dir -Recurse -Filter "*.xlsx" | Where-Object { $_.Name -notlike "~$*" })
    $formulaScan = Measure-FormulaErrorText $stage2Dir

    $stage2Summary = [ordered]@{
        OutputDirectory = $stage2Dir
        Summary = $stage2Report.Summary
        ReportPath = $stage2Report.ReportPath
        ValidationReport = Join-Path $stage2Dir "阶段二校验报告.txt"
        PreflightIssueCount = $preflight.Issues.Count
        ProxyRows = $stage2Report.ProxyRows
        IntermediaryRows = $stage2Report.IntermediaryRows
        ProxyGroups = $stage2Report.ProxyGroups
        IntermediaryGroups = $stage2Report.IntermediaryGroups
        ProxyWorkbookCount = $proxyXlsx.Count
        IntermediaryWorkbookCount = $intermediaryXlsx.Count
        XlsxCount = $stage2Xlsx.Count
        MissingOwners = $stage2Report.MissingOwners.Count
        Warnings = $stage2Report.Warnings.Count
        AuditIssueCount = $stage2Report.AuditIssues.Count
        FormulaFilesScanned = $formulaScan.FilesScanned
        FormulaErrorTextHits = $formulaScan.ErrorTextHits
    }
}

$summary = [ordered]@{
    Month = $Month
    SmokeRoot = $smokeRoot
    Configuration = $Configuration
    Clean = $cleanSummary
    Stage1 = $stage1Summary
    Stage2 = $stage2Summary
}

$summaryPath = Join-Path $smokeRoot "smoke-summary.json"
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding UTF8

Write-Host "Smoke output: $smokeRoot"
Write-Host "Smoke summary: $summaryPath"
$summary | ConvertTo-Json -Depth 8
