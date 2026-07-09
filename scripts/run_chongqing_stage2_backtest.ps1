param(
    [string]$CaseFile = "",
    [int[]]$Months = @(),
    [string]$LedgerPath = "",
    [string]$ProxyTemplateDirectory = "",
    [string]$IntermediaryTemplateDirectory = "",
    [string]$RefundTemplateDirectory = "",
    [string]$SummaryTemplatePath = "",
    [string]$OutputRoot = "",
    [string]$Configuration = "Release",
    [switch]$SkipBuild
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

function Assert-OptionalDirectory {
    param(
        [string]$Path,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
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

function Invoke-SolutionBuild {
    $dotnet = Get-DotNetPath
    if (-not $dotnet) {
        throw "dotnet could not be found."
    }

    Write-Host "Building with dotnet msbuild: $dotnet"
    & $dotnet msbuild $Solution /restore /p:Configuration=$Configuration /m
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet msbuild failed. Exit code: $LASTEXITCODE"
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

function Get-JsonProperty {
    param(
        $Object,
        [string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) {
        return $Default
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $Default
    }

    return $property.Value
}

function Measure-FormulaErrorText {
    param([string]$Root)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $patterns = @("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A")
    $hits = 0
    $files = @(Get-ChildItem -LiteralPath $Root -Recurse -Filter "*.xlsx" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "._*" })

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

function Measure-XlsxOpenability {
    param([string]$Root)

    $files = @(Get-ChildItem -LiteralPath $Root -Recurse -Filter "*.xlsx" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "._*" })
    $opened = 0
    $failures = @()

    foreach ($file in $files) {
        try {
            $workbook = [ClosedXML.Excel.XLWorkbook]::new($file.FullName)
            try {
                $null = $workbook.Worksheets.Count
            }
            finally {
                $workbook.Dispose()
            }

            $opened++
        }
        catch {
            $failures += [ordered]@{
                Path = $file.FullName
                Error = $_.Exception.Message
            }
        }
    }

    return [ordered]@{
        FilesScanned = $files.Count
        Opened = $opened
        Failed = $failures.Count
        Failures = @($failures | Select-Object -First 20)
    }
}

function Convert-Issues {
    param($Issues)

    return @($Issues | ForEach-Object {
        [ordered]@{
            Severity = $_.Severity
            Kind = $_.Kind
            SettlementKind = $_.SettlementKind
            Entity = $_.Entity
            Owner = $_.Owner
            RequiresPaymentPartySelection = $_.RequiresPaymentPartySelection
            Message = $_.Message
            Suggestion = $_.Suggestion
        }
    })
}

function Add-SummarySubjectDecisions {
    param(
        $Options,
        $DecisionSpecs
    )

    foreach ($spec in @($DecisionSpecs)) {
        if ($null -eq $spec) {
            continue
        }

        $decision = [HainanSettlementTool.Core.Models.ChongqingStage2SummarySubjectDecision]::new()
        $decision.SettlementKind = [string](Get-JsonProperty $spec "settlementKind" "")
        $decision.Entity = [string](Get-JsonProperty $spec "entity" "")
        $decision.PaymentParty = [string](Get-JsonProperty $spec "paymentParty" "")
        $Options.SummarySubjectDecisions.Add($decision)
    }
}

function Invoke-BacktestCase {
    param(
        $Service,
        $Case,
        $Manifest,
        [string]$BacktestRoot
    )

    $month = [int](Get-JsonProperty $Case "month" 0)
    $caseName = [string](Get-JsonProperty $Case "name" ("month-" + $month.ToString("00")))
    $caseOutput = [string](Get-JsonProperty $Case "outputDirectory" "")
    if ([string]::IsNullOrWhiteSpace($caseOutput)) {
        $caseOutput = Join-Path $BacktestRoot $caseName
    }

    $options = [HainanSettlementTool.Core.Models.ChongqingStage2Options]::new()
    $options.Month = $month
    $options.LedgerPath = [string](Get-JsonProperty $Case "ledgerPath" (Get-JsonProperty $Manifest "ledgerPath" ""))
    $options.ProxyTemplateDirectory = [string](Get-JsonProperty $Case "proxyTemplateDirectory" (Get-JsonProperty $Manifest "proxyTemplateDirectory" ""))
    $options.IntermediaryTemplateDirectory = [string](Get-JsonProperty $Case "intermediaryTemplateDirectory" (Get-JsonProperty $Manifest "intermediaryTemplateDirectory" ""))
    $options.RefundTemplateDirectory = [string](Get-JsonProperty $Case "refundTemplateDirectory" (Get-JsonProperty $Manifest "refundTemplateDirectory" ""))
    $options.SummaryTemplatePath = [string](Get-JsonProperty $Case "summaryTemplatePath" (Get-JsonProperty $Manifest "summaryTemplatePath" ""))
    $options.OutputDirectory = $caseOutput
    $options.OutputSummaryName = [string](Get-JsonProperty $Case "outputSummaryName" "")

    Add-SummarySubjectDecisions $options (Get-JsonProperty $Case "summarySubjectDecisions" @())

    $result = [ordered]@{
        Name = $caseName
        Month = $month
        OutputDirectory = $caseOutput
        LedgerPath = $options.LedgerPath
        SummaryTemplatePath = $options.SummaryTemplatePath
        Status = "Started"
    }

    try {
        Assert-File $options.LedgerPath "Chongqing ledger"
        Assert-Directory $options.ProxyTemplateDirectory "Chongqing proxy template"
        Assert-OptionalDirectory $options.IntermediaryTemplateDirectory "Chongqing intermediary template"
        Assert-Directory $options.RefundTemplateDirectory "Chongqing refund template"
        Assert-File $options.SummaryTemplatePath "Chongqing summary template"

        $preflight = $Service.Analyze($options)
        $paymentIssues = @($preflight.Issues | Where-Object { $_.RequiresPaymentPartySelection })
        $result.PreflightIssueCount = $preflight.Issues.Count
        $result.PaymentDecisionIssueCount = $paymentIssues.Count
        if ($paymentIssues.Count -gt 0) {
            $result.Status = "NeedsPaymentPartyDecisions"
            $result.PaymentDecisionIssues = Convert-Issues $paymentIssues
            return $result
        }

        $messages = New-Object System.Collections.Generic.List[string]
        $report = $Service.Run($options, { param($message) $messages.Add($message) })
        $xlsxFiles = @(Get-ChildItem -LiteralPath $caseOutput -Recurse -Filter "*.xlsx" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "._*" })
        $formulaScan = Measure-FormulaErrorText $caseOutput
        $openScan = Measure-XlsxOpenability $caseOutput

        $result.Status = "Passed"
        $result.SummaryWorkbook = $report.Summary
        $result.ReportPath = $report.ReportPath
        $result.ValidationReportPath = $report.ValidationReportPath
        $result.ProxyRows = $report.ProxyRows
        $result.IntermediaryRows = $report.IntermediaryRows
        $result.RefundRows = $report.RefundRows
        $result.ProxyGroups = $report.ProxyGroups
        $result.IntermediaryGroups = $report.IntermediaryGroups
        $result.RefundGroups = $report.RefundGroups
        $result.ProxyTotal = $report.ProxyTotal
        $result.IntermediaryTotal = $report.IntermediaryTotal
        $result.RefundTotal = $report.RefundTotal
        $result.WarningCount = $report.Warnings.Count
        $result.AuditIssueCount = $report.AuditIssues.Count
        $result.GeneratedWorkbookCount = $xlsxFiles.Count
        $result.FormulaFilesScanned = $formulaScan.FilesScanned
        $result.FormulaErrorTextHits = $formulaScan.ErrorTextHits
        $result.OpenedWorkbookCount = $openScan.Opened
        $result.OpenFailedWorkbookCount = $openScan.Failed
        $result.OpenFailures = $openScan.Failures
        $result.LogMessages = @($messages)

        if ($report.Warnings.Count -gt 0 -or $report.AuditIssues.Count -gt 0 -or $formulaScan.ErrorTextHits -gt 0 -or $openScan.Failed -gt 0) {
            $result.Status = "ReviewRequired"
        }

        return $result
    }
    catch {
        $result.Status = "Failed"
        $result.ErrorType = $_.Exception.GetType().FullName
        $result.Error = $_.Exception.Message
        return $result
    }
}

if (-not [string]::IsNullOrWhiteSpace($CaseFile)) {
    Assert-File $CaseFile "Case file"
    $manifest = Get-Content -LiteralPath $CaseFile -Raw -Encoding UTF8 | ConvertFrom-Json
}
else {
    if ($Months.Count -eq 0) {
        throw "Provide -CaseFile or at least one -Months value."
    }

    $caseRows = @($Months | ForEach-Object {
        [pscustomobject]@{
            name = "month-" + ([int]$_).ToString("00")
            month = [int]$_
            summaryTemplatePath = $SummaryTemplatePath
        }
    })

    $manifest = [pscustomobject]@{
        ledgerPath = $LedgerPath
        proxyTemplateDirectory = $ProxyTemplateDirectory
        intermediaryTemplateDirectory = $IntermediaryTemplateDirectory
        refundTemplateDirectory = $RefundTemplateDirectory
        summaryTemplatePath = $SummaryTemplatePath
        cases = $caseRows
    }
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = [string](Get-JsonProperty $manifest "outputRoot" "")
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
$backtestRoot = Join-Path $OutputRoot "chongqing-stage2-backtest-$stamp"
New-Item -ItemType Directory -Force -Path $backtestRoot | Out-Null

$gateway = [HainanSettlementTool.Excel.ClosedXmlSettlementExcelGateway]::new()
$service = [HainanSettlementTool.Core.Services.ChongqingStage2Service]::new($gateway)
$results = @()

foreach ($case in @($manifest.cases)) {
    $results += Invoke-BacktestCase $service $case $manifest $backtestRoot
}

$summary = [ordered]@{
    StartedAt = $stamp
    BacktestRoot = $backtestRoot
    Configuration = $Configuration
    CaseCount = $results.Count
    Passed = @($results | Where-Object { $_.Status -eq "Passed" }).Count
    ReviewRequired = @($results | Where-Object { $_.Status -eq "ReviewRequired" }).Count
    NeedsPaymentPartyDecisions = @($results | Where-Object { $_.Status -eq "NeedsPaymentPartyDecisions" }).Count
    Failed = @($results | Where-Object { $_.Status -eq "Failed" }).Count
    Cases = $results
}

$summaryPath = Join-Path $backtestRoot "chongqing-stage2-backtest-summary.json"
$summary | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $summaryPath -Encoding UTF8

Write-Host "Chongqing Stage 2 backtest output: $backtestRoot"
Write-Host "Chongqing Stage 2 backtest summary: $summaryPath"
$summary | ConvertTo-Json -Depth 10
