param(
    [string]$Root = "",
    [string]$CaseFile = "",
    [int[]]$Months = @(3, 4, 5),
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
$Tolerance = 0.0001

function Assert-File {
    param([string]$Path, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "$Label path is required." }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label file not found: $Path" }
}

function Assert-Directory {
    param([string]$Path, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "$Label path is required." }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw "$Label directory not found: $Path" }
}

function Assert-OptionalDirectory {
    param([string]$Path, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { throw "$Label directory not found: $Path" }
}

function Get-DotNetPath {
    $command = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($command -and $command.Source) { return $command.Source }
    if ($env:ProgramFiles) {
        $candidate = Join-Path $env:ProgramFiles "dotnet\dotnet.exe"
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    throw "dotnet could not be found."
}

function Invoke-SolutionBuild {
    $dotnet = Get-DotNetPath
    Write-Host "Building with dotnet msbuild: $dotnet"
    & $dotnet msbuild $Solution /restore /p:Configuration=$Configuration /m
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet msbuild failed. Exit code: $LASTEXITCODE"
    }
}

function Add-AssemblyResolution {
    param([string[]]$Directories)
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
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-SafeName {
    param([string]$Value)
    $text = if ([string]::IsNullOrWhiteSpace($Value)) { "case" } else { $Value }
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $text = $text.Replace($char, "_")
    }
    return $text
}

function Get-CustomerKey {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    return ([regex]::Replace($Value.Trim(), "\s+", "")).ToUpperInvariant()
}

function Get-CellText {
    param($Cell)
    if ($null -eq $Cell) { return "" }
    return ([string]$Cell.GetFormattedString()).Trim()
}

function Get-CellComparable {
    param($Cell)
    if ($null -eq $Cell -or $Cell.IsEmpty()) { return "" }
    $formula = [string]$Cell.FormulaA1
    if (-not [string]::IsNullOrWhiteSpace($formula)) {
        return "FORMULA:" + $formula.Trim()
    }
    return Get-CellText $Cell
}

function Get-CellNumberOrNull {
    param($Cell)
    if ($null -eq $Cell -or $Cell.IsEmpty()) { return $null }
    try {
        if ($Cell.DataType -eq [ClosedXML.Excel.XLDataType]::Number) {
            return [double]$Cell.GetDouble()
        }
    }
    catch {
    }
    $text = (Get-CellText $Cell).Replace(",", "")
    $value = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
        return $value
    }
    if ([double]::TryParse($text, [ref]$value)) {
        return $value
    }
    return $null
}

function Same-NumberOrBlank {
    param($Actual, $Expected)
    if ($null -eq $Actual -and $null -eq $Expected) { return $true }
    if ($null -eq $Actual -or $null -eq $Expected) { return $false }
    return [Math]::Abs([double]$Actual - [double]$Expected) -le $Tolerance
}

function Is-BlankCell {
    param($Cell)
    return $null -eq $Cell -or $Cell.IsEmpty() -or [string]::IsNullOrWhiteSpace((Get-CellText $Cell))
}

function Are-CellValuesEquivalent {
    param($ActualCell, $ExpectedCell)
    $actualNumber = Get-CellNumberOrNull $ActualCell
    $expectedNumber = Get-CellNumberOrNull $ExpectedCell
    if ($null -ne $actualNumber -and $null -ne $expectedNumber) {
        return [Math]::Abs([double]$actualNumber - [double]$expectedNumber) -le $Tolerance
    }
    if ((Is-BlankCell $ActualCell) -and $null -ne $expectedNumber) {
        return [Math]::Abs([double]$expectedNumber) -le $Tolerance
    }
    if ((Is-BlankCell $ExpectedCell) -and $null -ne $actualNumber) {
        return [Math]::Abs([double]$actualNumber) -le $Tolerance
    }
    return $false
}

function Open-WorkbookShared {
    param([string]$Path)
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        return [ClosedXML.Excel.XLWorkbook]::new($stream)
    }
    finally {
        $stream.Dispose()
    }
}

function Find-WorksheetWithHeader {
    param($Workbook, [string]$Header)
    foreach ($worksheet in $Workbook.Worksheets) {
        $lastColumn = if ($worksheet.LastColumnUsed()) { $worksheet.LastColumnUsed().ColumnNumber() } else { 1 }
        $lastRow = if ($worksheet.LastRowUsed()) { [Math]::Min(10, $worksheet.LastRowUsed().RowNumber()) } else { 1 }
        for ($row = 1; $row -le $lastRow; $row++) {
            for ($column = 1; $column -le $lastColumn; $column++) {
                if ((Get-CellText $worksheet.Cell($row, $column)) -eq $Header) {
                    return $worksheet
                }
            }
        }
    }
    throw "No worksheet with header '$Header' was found."
}

function Find-HeaderColumn {
    param($Worksheet, [string]$Header)
    $lastColumn = if ($Worksheet.LastColumnUsed()) { $Worksheet.LastColumnUsed().ColumnNumber() } else { 1 }
    $lastRow = if ($Worksheet.LastRowUsed()) { [Math]::Min(10, $Worksheet.LastRowUsed().RowNumber()) } else { 1 }
    for ($row = 1; $row -le $lastRow; $row++) {
        for ($column = 1; $column -le $lastColumn; $column++) {
            if ((Get-CellText $Worksheet.Cell($row, $column)) -eq $Header) {
                return $column
            }
        }
    }
    throw "Header '$Header' was not found."
}

function Find-LedgerMonthStart {
    param($Worksheet, [int]$Month)
    $label = $Month.ToString() + "月"
    $lastColumn = if ($Worksheet.LastColumnUsed()) { $Worksheet.LastColumnUsed().ColumnNumber() } else { 1 }
    for ($column = 1; $column -le $lastColumn; $column++) {
        if ((Get-CellText $Worksheet.Cell(1, $column)) -eq $label) {
            return $column
        }
    }
    throw "Ledger month block '$label' was not found."
}

function Read-LedgerRows {
    param($Worksheet, [int]$CustomerColumn)
    $rows = @{}
    $duplicates = 0
    $lastRow = if ($Worksheet.LastRowUsed()) { $Worksheet.LastRowUsed().RowNumber() } else { 4 }
    for ($row = 4; $row -le $lastRow; $row++) {
        $name = Get-CellText $Worksheet.Cell($row, $CustomerColumn)
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $key = Get-CustomerKey $name
        if ($rows.ContainsKey($key)) {
            $duplicates++
            continue
        }
        $rows[$key] = [pscustomobject]@{ Row = $row; Name = $name }
    }
    return [pscustomobject]@{ Rows = $rows; DuplicateKeys = $duplicates }
}

function Compare-LedgerMonthPower {
    param([string]$ActualPath, [string]$ExpectedPath, [int]$Month)
    $actualWorkbook = Open-WorkbookShared $ActualPath
    $expectedWorkbook = Open-WorkbookShared $ExpectedPath
    try {
        $actualSheet = Find-WorksheetWithHeader $actualWorkbook "电力用户名称"
        $expectedSheet = Find-WorksheetWithHeader $expectedWorkbook "电力用户名称"
        $actualNameColumn = Find-HeaderColumn $actualSheet "电力用户名称"
        $expectedNameColumn = Find-HeaderColumn $expectedSheet "电力用户名称"
        $actualStart = Find-LedgerMonthStart $actualSheet $Month
        $expectedStart = Find-LedgerMonthStart $expectedSheet $Month
        $actualRows = Read-LedgerRows $actualSheet $actualNameColumn
        $expectedRows = Read-LedgerRows $expectedSheet $expectedNameColumn
        $keys = New-Object System.Collections.Generic.HashSet[string]
        foreach ($key in $actualRows.Rows.Keys) { [void]$keys.Add($key) }
        foreach ($key in $expectedRows.Rows.Keys) { [void]$keys.Add($key) }

        $diffs = 0
        $missingActual = 0
        $missingExpected = 0
        $details = @()
        $labels = @("总", "尖", "峰", "平", "谷")
        foreach ($key in $keys) {
            if (-not $actualRows.Rows.ContainsKey($key)) {
                $missingActual++
                if ($details.Count -lt 100) { $details += [ordered]@{ Customer = $expectedRows.Rows[$key].Name; Field = "客户"; Expected = "存在"; Actual = "缺失" } }
                continue
            }
            if (-not $expectedRows.Rows.ContainsKey($key)) {
                $missingExpected++
                if ($details.Count -lt 100) { $details += [ordered]@{ Customer = $actualRows.Rows[$key].Name; Field = "客户"; Expected = "缺失"; Actual = "存在" } }
                continue
            }
            for ($offset = 0; $offset -lt 5; $offset++) {
                $actualValue = Get-CellNumberOrNull $actualSheet.Cell($actualRows.Rows[$key].Row, $actualStart + $offset)
                $expectedValue = Get-CellNumberOrNull $expectedSheet.Cell($expectedRows.Rows[$key].Row, $expectedStart + $offset)
                if (-not (Same-NumberOrBlank $actualValue $expectedValue)) {
                    $diffs++
                    if ($details.Count -lt 100) {
                        $details += [ordered]@{
                            Customer = $expectedRows.Rows[$key].Name
                            Field = $labels[$offset]
                            Expected = $expectedValue
                            Actual = $actualValue
                        }
                    }
                }
            }
        }

        return [ordered]@{
            ComparedCustomers = $keys.Count
            PowerValueDiffs = $diffs
            MissingInActual = $missingActual
            MissingInExpected = $missingExpected
            ActualDuplicateKeys = $actualRows.DuplicateKeys
            ExpectedDuplicateKeys = $expectedRows.DuplicateKeys
            Details = $details
        }
    }
    finally {
        $actualWorkbook.Dispose()
        $expectedWorkbook.Dispose()
    }
}

function Compare-WorksheetCells {
    param($ActualSheet, $ExpectedSheet, [int]$FirstRow = 1, [int]$FirstColumn = 1, [int]$LastRow = 0, [int]$LastColumn = 0)
    if ($LastRow -le 0) {
        $actualLastRow = if ($ActualSheet.LastRowUsed()) { $ActualSheet.LastRowUsed().RowNumber() } else { 1 }
        $expectedLastRow = if ($ExpectedSheet.LastRowUsed()) { $ExpectedSheet.LastRowUsed().RowNumber() } else { 1 }
        $LastRow = [Math]::Max($actualLastRow, $expectedLastRow)
    }
    if ($LastColumn -le 0) {
        $actualLastColumn = if ($ActualSheet.LastColumnUsed()) { $ActualSheet.LastColumnUsed().ColumnNumber() } else { 1 }
        $expectedLastColumn = if ($ExpectedSheet.LastColumnUsed()) { $ExpectedSheet.LastColumnUsed().ColumnNumber() } else { 1 }
        $LastColumn = [Math]::Max($actualLastColumn, $expectedLastColumn)
    }

    $formulaDiffs = 0
    $valueDiffs = 0
    $hiddenColumnDiffs = 0
    $details = @()
    for ($column = $FirstColumn; $column -le $LastColumn; $column++) {
        if ($ActualSheet.Column($column).IsHidden -ne $ExpectedSheet.Column($column).IsHidden) {
            $hiddenColumnDiffs++
            if ($details.Count -lt 100) {
                $details += [ordered]@{ Address = "Column " + $column; Difference = "HiddenColumn"; Expected = $ExpectedSheet.Column($column).IsHidden; Actual = $ActualSheet.Column($column).IsHidden }
            }
        }
    }

    for ($row = $FirstRow; $row -le $LastRow; $row++) {
        for ($column = $FirstColumn; $column -le $LastColumn; $column++) {
            $actual = $ActualSheet.Cell($row, $column)
            $expected = $ExpectedSheet.Cell($row, $column)
            $actualFormula = ([string]$actual.FormulaA1).Trim()
            $expectedFormula = ([string]$expected.FormulaA1).Trim()
            if ($actualFormula -ne $expectedFormula) {
                if (-not ([string]::IsNullOrWhiteSpace($actualFormula) -and [string]::IsNullOrWhiteSpace($expectedFormula))) {
                    $formulaDiffs++
                    if ($details.Count -lt 100) {
                        $details += [ordered]@{ Address = $actual.Address.ToString(); Difference = "Formula"; Expected = $expectedFormula; Actual = $actualFormula }
                    }
                    continue
                }
            }

            $actualValue = Get-CellComparable $actual
            $expectedValue = Get-CellComparable $expected
            if ($actualValue -ne $expectedValue) {
                if (Are-CellValuesEquivalent $actual $expected) {
                    continue
                }

                $valueDiffs++
                if ($details.Count -lt 100) {
                    $details += [ordered]@{ Address = $actual.Address.ToString(); Difference = "Value"; Expected = $expectedValue; Actual = $actualValue }
                }
            }
        }
    }

    return [ordered]@{
        ComparedRange = "R$FirstRow" + "C$FirstColumn" + ":R$LastRow" + "C$LastColumn"
        FormulaDiffs = $formulaDiffs
        ValueDiffs = $valueDiffs
        HiddenColumnDiffs = $hiddenColumnDiffs
        Details = $details
    }
}

function Find-SummaryMonthStart {
    param($Worksheet, [int]$Month)
    $lastColumn = if ($Worksheet.LastColumnUsed()) { $Worksheet.LastColumnUsed().ColumnNumber() } else { 1 }
    for ($column = 1; $column -le $lastColumn; $column++) {
        if ((Get-CellText $Worksheet.Cell(3, $column)) -ne "代理费") { continue }
        $header = Get-CellText $Worksheet.Cell(2, $column)
        try {
            if ($Worksheet.Cell(2, $column).DataType -eq [ClosedXML.Excel.XLDataType]::DateTime) {
                $date = $Worksheet.Cell(2, $column).GetDateTime()
                if ($date.Year -eq 2026 -and $date.Month -eq $Month) { return $column }
            }
        }
        catch {
        }
        if ($header.Contains($Month.ToString() + "月") -or $header.Contains("-" + $Month.ToString("00") + "-")) {
            return $column
        }
    }
    throw "Summary month block for $Month was not found on sheet $($Worksheet.Name)."
}

function Compare-WorkbookSheet {
    param([string]$ActualPath, [string]$ExpectedPath, [string]$SheetName, [int]$FirstRow = 1, [int]$FirstColumn = 1, [int]$LastRow = 0, [int]$LastColumn = 0)
    if (-not (Test-Path -LiteralPath $ActualPath -PathType Leaf)) {
        return [ordered]@{ Status = "MissingActualWorkbook"; ActualPath = $ActualPath; ExpectedPath = $ExpectedPath }
    }
    if (-not (Test-Path -LiteralPath $ExpectedPath -PathType Leaf)) {
        return [ordered]@{ Status = "MissingExpectedWorkbook"; ActualPath = $ActualPath; ExpectedPath = $ExpectedPath }
    }

    $actualWorkbook = Open-WorkbookShared $ActualPath
    $expectedWorkbook = Open-WorkbookShared $ExpectedPath
    try {
        $actualSheet = $actualWorkbook.Worksheets | Where-Object { $_.Name -eq $SheetName } | Select-Object -First 1
        $expectedSheet = $expectedWorkbook.Worksheets | Where-Object { $_.Name -eq $SheetName } | Select-Object -First 1
        if ($null -eq $actualSheet) { return [ordered]@{ Status = "MissingActualSheet"; Sheet = $SheetName; ActualPath = $ActualPath; ExpectedPath = $ExpectedPath } }
        if ($null -eq $expectedSheet) { return [ordered]@{ Status = "MissingExpectedSheet"; Sheet = $SheetName; ActualPath = $ActualPath; ExpectedPath = $ExpectedPath } }
        $comparison = Compare-WorksheetCells $actualSheet $expectedSheet $FirstRow $FirstColumn $LastRow $LastColumn
        $comparison["Status"] = if ($comparison.FormulaDiffs -eq 0 -and $comparison.ValueDiffs -eq 0 -and $comparison.HiddenColumnDiffs -eq 0) { "Matched" } else { "Different" }
        $comparison["Sheet"] = $SheetName
        $comparison["ActualPath"] = $ActualPath
        $comparison["ExpectedPath"] = $ExpectedPath
        return $comparison
    }
    finally {
        $actualWorkbook.Dispose()
        $expectedWorkbook.Dispose()
    }
}

function Compare-SummaryMainMonthBlock {
    param([string]$ActualPath, [string]$ExpectedPath, [int]$Month)
    $actualWorkbook = Open-WorkbookShared $ActualPath
    $expectedWorkbook = Open-WorkbookShared $ExpectedPath
    try {
        $actualSheet = $actualWorkbook.Worksheet("汇总表")
        $expectedSheet = $expectedWorkbook.Worksheet("汇总表")
        $actualStart = Find-SummaryMonthStart $actualSheet $Month
        $expectedStart = Find-SummaryMonthStart $expectedSheet $Month
        $lastRow = [Math]::Max($actualSheet.LastRowUsed().RowNumber(), $expectedSheet.LastRowUsed().RowNumber())
        $comparison = Compare-WorksheetCells $actualSheet $expectedSheet 1 $actualStart $lastRow ($actualStart + 5)
        $comparison["Status"] = if ($comparison.FormulaDiffs -eq 0 -and $comparison.ValueDiffs -eq 0 -and $comparison.HiddenColumnDiffs -eq 0) { "Matched" } else { "Different" }
        $comparison["Sheet"] = "汇总表"
        $comparison["ActualPath"] = $ActualPath
        $comparison["ExpectedPath"] = $ExpectedPath
        return $comparison
    }
    finally {
        $actualWorkbook.Dispose()
        $expectedWorkbook.Dispose()
    }
}

function Get-RelativePath {
    param([string]$Root, [string]$Path)
    $rootUri = [Uri]((Resolve-Path -LiteralPath $Root).Path.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar)
    $pathUri = [Uri]((Resolve-Path -LiteralPath $Path).Path)
    return [Uri]::UnescapeDataString($rootUri.MakeRelativeUri($pathUri).ToString()).Replace('/', [System.IO.Path]::DirectorySeparatorChar)
}

function Add-Stage1Decisions {
    param($Options, $DecisionSpecs)
    foreach ($spec in @($DecisionSpecs)) {
        if ($null -eq $spec) { continue }
        $decision = [HainanSettlementTool.Core.Models.ProvinceStage1CustomerDecision]::new()
        $decision.SourceCustomerName = [string](Get-JsonProperty $spec "sourceCustomerName" "")
        $decision.TargetCustomerName = [string](Get-JsonProperty $spec "targetCustomerName" "")
        $kind = [string](Get-JsonProperty $spec "decisionKind" "")
        $decision.DecisionKind = [Enum]::Parse([HainanSettlementTool.Core.Models.ProvinceStage1CustomerDecisionKind], $kind)
        $Options.CustomerDecisions.Add($decision)
    }
}

function Select-ApplicableStage1DecisionSpecs {
    param($DecisionSpecs, $PowerOnlyCustomers)
    $powerOnlyKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($customer in @($PowerOnlyCustomers)) {
        [void]$powerOnlyKeys.Add((Get-CustomerKey $customer))
    }

    $selected = @()
    foreach ($spec in @($DecisionSpecs)) {
        if ($null -eq $spec) { continue }
        $source = [string](Get-JsonProperty $spec "sourceCustomerName" "")
        if ($powerOnlyKeys.Contains((Get-CustomerKey $source))) {
            $selected += $spec
        }
    }
    return $selected
}

function Add-Stage2Decisions {
    param($Options, $DecisionSpecs)
    foreach ($spec in @($DecisionSpecs)) {
        if ($null -eq $spec) { continue }
        $decision = [HainanSettlementTool.Core.Models.ChongqingStage2SummarySubjectDecision]::new()
        $decision.SettlementKind = [string](Get-JsonProperty $spec "settlementKind" "")
        $decision.Entity = [string](Get-JsonProperty $spec "entity" "")
        $decision.PaymentParty = [string](Get-JsonProperty $spec "paymentParty" "")
        $Options.SummarySubjectDecisions.Add($decision)
    }
}

function Get-PowerFileMonth {
    param([string]$FileName)
    $match = [regex]::Match($FileName, "20\d{2}年0?(?<month>[1-9]|1[0-2])月.*电量确认")
    if (-not $match.Success) {
        $match = [regex]::Match($FileName, "(^|[^\d])0?(?<month>[1-9]|1[0-2])月.*电量确认")
    }
    if ($match.Success) {
        return [int]$match.Groups["month"].Value
    }
    return 0
}

function Resolve-MonthPowerFiles {
    param([string]$Root, [int[]]$Months)
    $files = @(Get-ChildItem -LiteralPath $Root -Filter "*.xlsx" -File |
        Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "._*" -and $_.Name -like "*电量确认*" })
    $result = @{}
    foreach ($month in $Months) {
        $matches = @($files | Where-Object { (Get-PowerFileMonth $_.Name) -eq $month })
        if ($matches.Count -ne 1) {
            throw "Expected exactly one power file for month $month, found $($matches.Count)."
        }
        $result[$month] = $matches[0].FullName
    }
    return $result
}

function New-AutoManifest {
    param([string]$Root, [int[]]$Months)
    Assert-Directory $Root "Backtest root"
    $powerFiles = Resolve-MonthPowerFiles $Root $Months
    $ledger = if ([string]::IsNullOrWhiteSpace($LedgerPath)) {
        @(Get-ChildItem -LiteralPath $Root -Filter "重庆2026年售电结算台账*.xlsx" -File | Where-Object { $_.Name -notlike "~$*" } | Sort-Object Name | Select-Object -Last 1).FullName
    } else { $LedgerPath }
    $summary = if ([string]::IsNullOrWhiteSpace($SummaryTemplatePath)) {
        @(Get-ChildItem -LiteralPath $Root -Filter "重庆2026年代理费汇总表*.xlsx" -File | Where-Object { $_.Name -notlike "~$*" } | Sort-Object Name | Select-Object -Last 1).FullName
    } else { $SummaryTemplatePath }
    $proxy = if ([string]::IsNullOrWhiteSpace($ProxyTemplateDirectory)) { Join-Path $Root "重庆2026代理" } else { $ProxyTemplateDirectory }
    $refund = if ([string]::IsNullOrWhiteSpace($RefundTemplateDirectory)) { Join-Path $Root "重庆 2026 退补" } else { $RefundTemplateDirectory }
    $intermediary = if ([string]::IsNullOrWhiteSpace($IntermediaryTemplateDirectory)) { "" } else { $IntermediaryTemplateDirectory }
    $cases = @($Months | ForEach-Object {
        [pscustomobject]@{
            name = "month-" + ([int]$_).ToString("00")
            month = [int]$_
            rawDetailPath = $powerFiles[[int]$_]
            stage1CustomerDecisions = @()
            stage2SummarySubjectDecisions = @()
        }
    })
    return [pscustomobject]@{
        ledgerPath = $ledger
        proxyTemplateDirectory = $proxy
        intermediaryTemplateDirectory = $intermediary
        refundTemplateDirectory = $refund
        summaryTemplatePath = $summary
        cases = $cases
    }
}

function Invoke-BacktestCase {
    param($Stage1Service, $Stage2Service, $Case, $Manifest, [string]$BacktestRoot)
    $month = [int](Get-JsonProperty $Case "month" 0)
    $name = [string](Get-JsonProperty $Case "name" ("month-" + $month.ToString("00")))
    $caseRoot = Join-Path $BacktestRoot (Get-SafeName $name)
    $stage1CleanDir = Join-Path $caseRoot "stage1-clean"
    $stage1LedgerDir = Join-Path $caseRoot "stage1-ledger"
    $stage2Dir = Join-Path $caseRoot "stage2"
    New-Item -ItemType Directory -Force -Path $stage1CleanDir, $stage1LedgerDir, $stage2Dir | Out-Null

    $ledgerPath = [string](Get-JsonProperty $Case "ledgerPath" (Get-JsonProperty $Manifest "ledgerPath" ""))
    $rawPath = [string](Get-JsonProperty $Case "rawDetailPath" "")
    $summaryPath = [string](Get-JsonProperty $Case "summaryTemplatePath" (Get-JsonProperty $Manifest "summaryTemplatePath" ""))
    $proxyDir = [string](Get-JsonProperty $Case "proxyTemplateDirectory" (Get-JsonProperty $Manifest "proxyTemplateDirectory" ""))
    $intermediaryDir = [string](Get-JsonProperty $Case "intermediaryTemplateDirectory" (Get-JsonProperty $Manifest "intermediaryTemplateDirectory" ""))
    $refundDir = [string](Get-JsonProperty $Case "refundTemplateDirectory" (Get-JsonProperty $Manifest "refundTemplateDirectory" ""))

    $caseResult = [ordered]@{ Name = $name; Month = $month; OutputDirectory = $caseRoot; Status = "Started" }

    try {
        Assert-File $ledgerPath "Chongqing ledger"
        Assert-File $rawPath "Chongqing power file"
        Assert-File $summaryPath "Chongqing summary workbook"
        Assert-Directory $proxyDir "Chongqing proxy directory"
        Assert-Directory $refundDir "Chongqing refund directory"
        Assert-OptionalDirectory $intermediaryDir "Chongqing intermediary directory"

        $cleanOptions = [HainanSettlementTool.Core.Models.ProvinceStage1CleanOptions]::new()
        $cleanOptions.Province = [HainanSettlementTool.Core.Models.ProvinceCode]::Chongqing
        $cleanOptions.Month = $month
        $cleanOptions.RawDetailPath = $rawPath
        $cleanOptions.OutputDirectory = $stage1CleanDir
        $cleanReport = $Stage1Service.CleanPowerData($cleanOptions, { param($message) })

        $ledgerOptions = [HainanSettlementTool.Core.Models.ProvinceStage1LedgerUpdateOptions]::new()
        $ledgerOptions.Province = [HainanSettlementTool.Core.Models.ProvinceCode]::Chongqing
        $ledgerOptions.Month = $month
        $ledgerOptions.LedgerPath = $ledgerPath
        $ledgerOptions.RawDetailPath = $rawPath
        $ledgerOptions.OutputDirectory = $stage1LedgerDir
        $plan = $Stage1Service.PlanLedgerUpdate($ledgerOptions, { param($message) })
        $stage1DecisionSpecs = @(Get-JsonProperty $Case "stage1CustomerDecisions" @())
        $applicableStage1DecisionSpecs = @(Select-ApplicableStage1DecisionSpecs $stage1DecisionSpecs $plan.PowerOnlyCustomers)
        if ($applicableStage1DecisionSpecs.Count -gt 0) {
            Add-Stage1Decisions $ledgerOptions $applicableStage1DecisionSpecs
            $plan = $Stage1Service.PlanLedgerUpdate($ledgerOptions, { param($message) })
        }
        $unresolved = @($plan.PowerOnlyCustomers | Where-Object {
            $sourceKey = Get-CustomerKey $_
            -not @($ledgerOptions.CustomerDecisions | ForEach-Object { Get-CustomerKey $_.SourceCustomerName }).Contains($sourceKey)
        })

        $stage1 = [ordered]@{
            CleanOutput = $cleanReport.OutputWorkbookPath
            CleanReport = $cleanReport.ReportPath
            CleanRows = $cleanReport.CustomerRows
            CleanTotalPower = $cleanReport.TotalPower
            PlanIssueCount = $plan.Issues.Count
            PowerOnlyCustomers = $plan.PowerOnlyCustomers.Count
            ConfiguredCustomerDecisions = $stage1DecisionSpecs.Count
            AppliedCustomerDecisions = $ledgerOptions.CustomerDecisions.Count
            UnresolvedPowerOnlyCustomers = $unresolved
        }

        if ($unresolved.Count -eq 0) {
            $ledgerReport = $Stage1Service.UpdateLedger($ledgerOptions, { param($message) })
            $stage1.UpdateOutput = $ledgerReport.OutputLedgerPath
            $stage1.UpdateReport = $ledgerReport.ReportPath
            $stage1.UpdatedPowerRows = $ledgerReport.UpdatedPowerRows
            $stage1.ManualMatchedRows = $ledgerReport.ManualMatchedRows
            $stage1.CreatedCustomerRows = $ledgerReport.CreatedCustomerRows
            $stage1.SkippedCustomerRows = $ledgerReport.SkippedCustomerRows
            $stage1.LedgerMonthPowerCompare = Compare-LedgerMonthPower $ledgerReport.OutputLedgerPath $ledgerPath $month
        }
        else {
            $stage1.Status = "NeedsStage1CustomerDecisions"
        }

        $stage2Options = [HainanSettlementTool.Core.Models.ChongqingStage2Options]::new()
        $stage2Options.Month = $month
        $stage2Options.LedgerPath = $ledgerPath
        $stage2Options.ProxyTemplateDirectory = $proxyDir
        $stage2Options.IntermediaryTemplateDirectory = $intermediaryDir
        $stage2Options.RefundTemplateDirectory = $refundDir
        $stage2Options.SummaryTemplatePath = $summaryPath
        $stage2Options.OutputDirectory = $stage2Dir
        Add-Stage2Decisions $stage2Options (Get-JsonProperty $Case "stage2SummarySubjectDecisions" @())
        $preflight = $Stage2Service.Analyze($stage2Options)
        $paymentIssues = @($preflight.Issues | Where-Object { $_.RequiresPaymentPartySelection })
        $stage2 = [ordered]@{ PreflightIssueCount = $preflight.Issues.Count; PaymentDecisionIssueCount = $paymentIssues.Count }
        if ($paymentIssues.Count -eq 0) {
            $stage2Report = $Stage2Service.Run($stage2Options, { param($message) })
            $stage2.ReportPath = $stage2Report.ReportPath
            $stage2.ValidationReportPath = $stage2Report.ValidationReportPath
            $stage2.SummaryWorkbook = $stage2Report.Summary
            $stage2.ProxyRows = $stage2Report.ProxyRows
            $stage2.IntermediaryRows = $stage2Report.IntermediaryRows
            $stage2.RefundRows = $stage2Report.RefundRows
            $stage2.ProxyGroups = $stage2Report.ProxyGroups
            $stage2.IntermediaryGroups = $stage2Report.IntermediaryGroups
            $stage2.RefundGroups = $stage2Report.RefundGroups
            $stage2.ProxyTotal = $stage2Report.ProxyTotal
            $stage2.IntermediaryTotal = $stage2Report.IntermediaryTotal
            $stage2.RefundTotal = $stage2Report.RefundTotal
            $stage2.WarningCount = $stage2Report.Warnings.Count
            $stage2.AuditIssueCount = $stage2Report.AuditIssues.Count

            $splitComparisons = @()
            foreach ($pair in @(
                @{ ActualRoot = $stage2Report.ProxyOutputDirectory; ExpectedRoot = $proxyDir; Kind = "代理" },
                @{ ActualRoot = $stage2Report.RefundOutputDirectory; ExpectedRoot = $refundDir; Kind = "退补" },
                @{ ActualRoot = $stage2Report.IntermediaryOutputDirectory; ExpectedRoot = $intermediaryDir; Kind = "居间" }
            )) {
                if ([string]::IsNullOrWhiteSpace($pair.ExpectedRoot) -or -not (Test-Path -LiteralPath $pair.ActualRoot -PathType Container)) { continue }
                $actualFiles = @(Get-ChildItem -LiteralPath $pair.ActualRoot -Recurse -Filter "*.xlsx" -File | Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "._*" })
                foreach ($actualFile in $actualFiles) {
                    $relative = Get-RelativePath $pair.ActualRoot $actualFile.FullName
                    $expectedPath = Join-Path $pair.ExpectedRoot $relative
                    $comparison = Compare-WorkbookSheet $actualFile.FullName $expectedPath ($month.ToString())
                    $comparison["Kind"] = $pair.Kind
                    $comparison["RelativePath"] = $relative
                    $splitComparisons += $comparison
                }
            }

            $summaryComparisons = @()
            $summaryComparisons += Compare-SummaryMainMonthBlock $stage2Report.Summary $summaryPath $month
            foreach ($party in @("清能", "清辉")) {
                $sheetName = $party + $month.ToString() + "月"
                $summaryComparisons += Compare-WorkbookSheet $stage2Report.Summary $summaryPath $sheetName
            }

            $stage2.SplitComparisons = $splitComparisons
            $stage2.SummaryComparisons = $summaryComparisons
            $stage2.SplitDifferent = @($splitComparisons | Where-Object { $_.Status -ne "Matched" }).Count
            $stage2.SummaryDifferent = @($summaryComparisons | Where-Object { $_.Status -ne "Matched" }).Count
        }
        else {
            $stage2.Status = "NeedsStage2PaymentDecisions"
            $stage2.PaymentDecisionIssues = @($paymentIssues | ForEach-Object { [ordered]@{ Kind = $_.SettlementKind; Entity = $_.Entity; Message = $_.Message } })
        }

        $caseResult.Stage1 = $stage1
        $caseResult.Stage2 = $stage2
        $caseResult.Status = "Completed"
        if (($stage1.Status -eq "NeedsStage1CustomerDecisions") -or ($stage2.Status -eq "NeedsStage2PaymentDecisions")) {
            $caseResult.Status = "NeedsDecisions"
        }
        elseif (($null -ne $stage1.LedgerMonthPowerCompare -and (($stage1.LedgerMonthPowerCompare.PowerValueDiffs -gt 0) -or ($stage1.LedgerMonthPowerCompare.MissingInActual -gt 0) -or ($stage1.LedgerMonthPowerCompare.MissingInExpected -gt 0))) -or ($stage2.SplitDifferent -gt 0) -or ($stage2.SummaryDifferent -gt 0) -or ($stage2.WarningCount -gt 0) -or ($stage2.AuditIssueCount -gt 0)) {
            $caseResult.Status = "ReviewRequired"
        }
        else {
            $caseResult.Status = "Passed"
        }
        return $caseResult
    }
    catch {
        $caseResult.Status = "Failed"
        $caseResult.ErrorType = $_.Exception.GetType().FullName
        $caseResult.Error = $_.Exception.Message
        return $caseResult
    }
}

if (-not [string]::IsNullOrWhiteSpace($CaseFile)) {
    Assert-File $CaseFile "Case file"
    $manifest = Get-Content -LiteralPath $CaseFile -Raw -Encoding UTF8 | ConvertFrom-Json
}
else {
    $manifest = New-AutoManifest $Root $Months
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = [string](Get-JsonProperty $manifest "outputRoot" "")
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $ProjectRoot "local-validation\backtests"
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
$backtestRoot = Join-Path $OutputRoot "chongqing-backtest-$stamp"
New-Item -ItemType Directory -Force -Path $backtestRoot | Out-Null

$gateway = [HainanSettlementTool.Excel.ClosedXmlSettlementExcelGateway]::new()
$stage1Service = [HainanSettlementTool.Core.Services.ProvinceStage1Service]::new($gateway)
$stage2Service = [HainanSettlementTool.Core.Services.ChongqingStage2Service]::new($gateway)

$results = @()
foreach ($case in @($manifest.cases)) {
    $results += Invoke-BacktestCase $stage1Service $stage2Service $case $manifest $backtestRoot
}

$summary = [ordered]@{
    StartedAt = $stamp
    BacktestRoot = $backtestRoot
    Configuration = $Configuration
    CaseCount = $results.Count
    Passed = @($results | Where-Object { $_.Status -eq "Passed" }).Count
    ReviewRequired = @($results | Where-Object { $_.Status -eq "ReviewRequired" }).Count
    NeedsDecisions = @($results | Where-Object { $_.Status -eq "NeedsDecisions" }).Count
    Failed = @($results | Where-Object { $_.Status -eq "Failed" }).Count
    Cases = $results
}

$summaryPath = Join-Path $backtestRoot "chongqing-backtest-summary.json"
$summary | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath $summaryPath -Encoding UTF8

Write-Host "Chongqing backtest output: $backtestRoot"
Write-Host "Chongqing backtest summary: $summaryPath"
$summary | ConvertTo-Json -Depth 16
