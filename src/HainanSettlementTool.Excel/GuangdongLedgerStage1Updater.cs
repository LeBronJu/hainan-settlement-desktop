using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class GuangdongLedgerStage1Updater
    {
        private const string LedgerSheetName = "广东2026年结算台账";
        private const int SettlementYear = 2026;
        private const int PowerColumnCount = 4;
        private readonly GuangdongPowerCleanGenerator _powerCleanGenerator;

        public GuangdongLedgerStage1Updater(GuangdongPowerCleanGenerator powerCleanGenerator)
        {
            _powerCleanGenerator = powerCleanGenerator ?? throw new ArgumentNullException(nameof(powerCleanGenerator));
        }

        public ProvinceStage1LedgerUpdatePlan Plan(ProvinceStage1LedgerUpdateOptions options)
        {
            RejectUnsupportedCustomerDecisions(options);
            using (var stream = File.Open(options.LedgerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                return BuildContext(workbook, options, false).Plan;
            }
        }

        public ProvinceStage1LedgerUpdateResult Update(ProvinceStage1LedgerUpdateOptions options)
        {
            RejectUnsupportedCustomerDecisions(options);
            using (var stream = File.Open(options.LedgerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var context = BuildContext(workbook, options, true);
                var outputPath = UniquePath(Path.Combine(
                    options.OutputDirectory,
                    options.Month.ToString(CultureInfo.InvariantCulture) + "月广东售电结算台账-阶段一更新.xlsx"));
                var reportPath = UniquePath(Path.Combine(
                    options.OutputDirectory,
                    options.Month.ToString(CultureInfo.InvariantCulture) + "月广东阶段一台账更新报告.json"));
                var htmlReportPath = UniquePath(Path.Combine(
                    options.OutputDirectory,
                    options.Month.ToString(CultureInfo.InvariantCulture) + "月广东阶段一台账更新报告.html"));
                EnsureOutputDoesNotOverwriteInput(outputPath, options.LedgerPath, options.RawDetailPath);
                FileAccessGuard.RequireWritableWorkbook(outputPath, "广东台账更新输出文件");

                var coefficientSnapshots = SnapshotCoefficients(
                    context.Worksheet,
                    context.MonthMap,
                    context.LedgerRows);
                ClearTargetMonthValues(context.Worksheet, context.MonthMap, context.LedgerRows);

                var updatedRows = 0;
                foreach (var ledgerRow in context.LedgerRows.ToList())
                {
                    GuangdongStage1SourceReader.GuangdongPowerAggregateRow powerRow;
                    if (!context.PowerRowsByCode.TryGetValue(ledgerRow.CodeKey, out powerRow))
                    {
                        RestoreCoefficients(
                            context.Worksheet,
                            context.MonthMap,
                            ledgerRow.RowNumber,
                            coefficientSnapshots[ledgerRow.CodeKey]);
                        continue;
                    }

                    WritePower(context.Worksheet, context.MonthMap, ledgerRow.RowNumber, powerRow);
                    if (powerRow.HasCoefficientPair)
                    {
                        WriteCoefficients(context.Worksheet, context.MonthMap, ledgerRow.RowNumber, powerRow);
                    }
                    else
                    {
                        RestoreCoefficients(
                            context.Worksheet,
                            context.MonthMap,
                            ledgerRow.RowNumber,
                            coefficientSnapshots[ledgerRow.CodeKey]);
                    }

                    updatedRows++;
                }

                foreach (var powerRow in context.NewCustomerRows)
                {
                    InsertNewCustomerRow(
                        context.Worksheet,
                        context.FixedMap,
                        context.MonthMap,
                        context.LedgerRows,
                        powerRow,
                        options.Month);
                    updatedRows++;
                }

                VerifyWrittenValues(context.Worksheet, context.MonthMap, context.LedgerRows, context.PowerRowsByCode);
                workbook.SaveAs(outputPath);
                var cleanWorkbookPath = _powerCleanGenerator.WriteCleanWorkbook(
                    new ProvinceStage1CleanOptions
                    {
                        Province = ProvinceCode.Guangdong,
                        Month = options.Month,
                        RawDetailPath = options.RawDetailPath,
                        OutputDirectory = options.OutputDirectory,
                        OutputWorkbookName = options.Month.ToString(CultureInfo.InvariantCulture)
                            + "月广东零售侧用户电量数据处理表.xlsx"
                    },
                    context.PowerData);

                var result = new ProvinceStage1LedgerUpdateResult
                {
                    Province = ProvinceCode.Guangdong,
                    Month = options.Month,
                    Unit = GuangdongStage1SourceReader.Unit,
                    LedgerPath = options.LedgerPath,
                    RawDetailPath = options.RawDetailPath,
                    OutputPowerWorkbookPath = cleanWorkbookPath,
                    OutputLedgerPath = outputPath,
                    ReportPath = reportPath,
                    HtmlReportPath = htmlReportPath,
                    LedgerCustomerRows = context.Plan.LedgerCustomerRows,
                    PowerCustomerRows = context.Plan.PowerCustomerRows,
                    MatchedRows = context.Plan.MatchedRows,
                    UpdatedPowerRows = updatedRows,
                    ManualMatchedRows = 0,
                    CreatedCustomerRows = context.Plan.CreatedCustomerRows,
                    SkippedCustomerRows = 0,
                    MultiAccountRows = context.PowerData.CustomerRows.Count(row => row.SourceRows > 1),
                    SkippedRows = 0,
                    TotalPower = Convert.ToDouble(
                        context.PowerData.CustomerRows.Sum(row => row.Total),
                        CultureInfo.InvariantCulture),
                    Warnings = context.Plan.Warnings.ToList(),
                    Issues = context.Plan.Issues.ToList()
                };
                new GuangdongStage1ReportWriter().WriteLedgerReport(result);
                return result;
            }
        }

        private LedgerUpdateContext BuildContext(
            XLWorkbook workbook,
            ProvinceStage1LedgerUpdateOptions options,
            bool createTargetMonth)
        {
            var powerData = _powerCleanGenerator.ReadData(new ProvinceStage1CleanOptions
            {
                Province = ProvinceCode.Guangdong,
                Month = options.Month,
                RawDetailPath = options.RawDetailPath,
                OutputDirectory = options.OutputDirectory
            });
            var worksheet = FindLedgerWorksheet(workbook);
            var fixedMap = FindFixedMap(worksheet);
            var ledgerWarnings = new List<string>();
            var monthMap = ResolveMonthMap(
                worksheet,
                options.Month,
                createTargetMonth,
                ledgerWarnings);
            var ledgerRows = ReadLedgerRows(worksheet, fixedMap);
            var ledgerRowsByCode = ledgerRows.ToDictionary(row => row.CodeKey, StringComparer.Ordinal);
            var powerRowsByCode = powerData.CustomerRows.ToDictionary(row => row.Code, StringComparer.Ordinal);
            var matchedCodes = ledgerRowsByCode.Keys.Intersect(powerRowsByCode.Keys, StringComparer.Ordinal).ToList();
            var newCustomerRows = powerData.CustomerRows
                .Where(row => !ledgerRowsByCode.ContainsKey(row.Code))
                .ToList();
            var plan = BuildPlan(
                options,
                powerData,
                worksheet,
                monthMap,
                ledgerRows,
                ledgerRowsByCode,
                powerRowsByCode,
                matchedCodes,
                newCustomerRows,
                ledgerWarnings);

            return new LedgerUpdateContext
            {
                Worksheet = worksheet,
                FixedMap = fixedMap,
                MonthMap = monthMap,
                LedgerRows = ledgerRows,
                PowerData = powerData,
                PowerRowsByCode = powerRowsByCode,
                NewCustomerRows = newCustomerRows,
                Plan = plan
            };
        }

        private static ProvinceStage1LedgerUpdatePlan BuildPlan(
            ProvinceStage1LedgerUpdateOptions options,
            GuangdongStage1SourceReader.GuangdongStage1DataSet powerData,
            IXLWorksheet worksheet,
            MonthMap monthMap,
            IList<LedgerCustomerRow> ledgerRows,
            IDictionary<string, LedgerCustomerRow> ledgerRowsByCode,
            IDictionary<string, GuangdongStage1SourceReader.GuangdongPowerAggregateRow> powerRowsByCode,
            IList<string> matchedCodes,
            IList<GuangdongStage1SourceReader.GuangdongPowerAggregateRow> newCustomerRows,
            IList<string> ledgerWarnings)
        {
            var plan = new ProvinceStage1LedgerUpdatePlan
            {
                Province = ProvinceCode.Guangdong,
                Month = options.Month,
                Unit = GuangdongStage1SourceReader.Unit,
                LedgerCustomerRows = ledgerRows.Count,
                PowerCustomerRows = powerRowsByCode.Count,
                MatchedRows = matchedCodes.Count,
                CreatedCustomerRows = newCustomerRows.Count,
                MissingInLedgerRows = newCustomerRows.Count,
                MissingInPowerRows = ledgerRows.Count - matchedCodes.Count,
                MultiAccountRows = powerData.CustomerRows.Count(row => row.SourceRows > 1),
                Warnings = powerData.Warnings.Concat(ledgerWarnings).ToList()
            };

            if (powerData.Month > 0 && powerData.Month != options.Month)
            {
                AddIssue(
                    plan,
                    ProvinceStage1LedgerUpdateIssueKinds.MonthMismatch,
                    "月份不一致",
                    "警告",
                    null,
                    "零售结算明细文件名识别为"
                    + powerData.Month.ToString(CultureInfo.InvariantCulture)
                    + "月，界面选择为"
                    + options.Month.ToString(CultureInfo.InvariantCulture)
                    + "月。");
            }

            foreach (var code in matchedCodes)
            {
                var ledgerRow = ledgerRowsByCode[code];
                var powerRow = powerRowsByCode[code];
                if (!string.Equals(ledgerRow.CustomerNameKey, powerRow.CustomerNameKey, StringComparison.Ordinal))
                {
                    AddIssue(
                        plan,
                        ProvinceStage1LedgerUpdateIssueKinds.CustomerNameMismatch,
                        "同编码名称不同",
                        "警告",
                        ledgerRow.CustomerName,
                        "用户编号“"
                        + powerRow.Code
                        + "”在电量表中的名称为“"
                        + powerRow.CustomerName
                        + "”；本次仍按编码写入，且不会改名。");
                }

                if (monthMap.TargetMonthAlreadyPresent
                    && !IsBlankPower(worksheet, ledgerRow.RowNumber, monthMap)
                    && !SamePowerVector(worksheet, ledgerRow.RowNumber, monthMap, powerRow))
                {
                    plan.ExistingDifferentPowerRows++;
                    AddIssue(
                        plan,
                        ProvinceStage1LedgerUpdateIssueKinds.ExistingPowerDifference,
                        "已有电量差异",
                        "警告",
                        ledgerRow.CustomerName,
                        "台账目标月份已有电量且与本次零售结算明细不一致；继续后会先清空所有客户当月电量，再按本次明细写入副本。");
                }

                if (powerRow.CoefficientConflictRows > 0)
                {
                    AddIssue(
                        plan,
                        ProvinceStage1LedgerUpdateIssueKinds.CoefficientConflict,
                        "峰平谷系数冲突",
                        "提示",
                        ledgerRow.CustomerName,
                        "同一用户编号的后续计量点存在不同系数；系数不参与代理费计算，本次按原始行顺序采用首个有效完整系数。");
                }
            }

            var ledgerRowsByName = ledgerRows
                .GroupBy(row => row.CustomerNameKey, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
            foreach (var newCustomer in newCustomerRows)
            {
                AddIssue(
                    plan,
                    ProvinceStage1LedgerUpdateIssueKinds.CreatedCustomer,
                    "新增客户到台账",
                    "警告",
                    newCustomer.CustomerName,
                    "用户编号“"
                    + newCustomer.Code
                    + "”将追加到输出台账底部；程序只写序号、编码、名称、履约开始月份、当月电量和有效系数，其余业务资料留空待人工补齐。");

                List<LedgerCustomerRow> sameNameRows;
                if (ledgerRowsByName.TryGetValue(newCustomer.CustomerNameKey, out sameNameRows))
                {
                    plan.AliasCandidateRows++;
                    AddIssue(
                        plan,
                        ProvinceStage1LedgerUpdateIssueKinds.PossibleAlias,
                        "同名不同编码",
                        "警告",
                        newCustomer.CustomerName,
                        "电量表用户编号“"
                        + newCustomer.Code
                        + "”与台账其他编码同名（"
                        + string.Join("、", sameNameRows.Select(row => row.Code))
                        + "）；程序不会自动合并，确认后仍按新编码追加。");
                }

                if (newCustomer.CoefficientConflictRows > 0)
                {
                    AddIssue(
                        plan,
                        ProvinceStage1LedgerUpdateIssueKinds.CoefficientConflict,
                        "峰平谷系数冲突",
                        "提示",
                        newCustomer.CustomerName,
                        "同一用户编号的后续计量点存在不同系数；系数不参与代理费计算，本次按原始行顺序采用首个有效完整系数。");
                }
            }

            foreach (var ledgerRow in ledgerRows.Where(row => !powerRowsByCode.ContainsKey(row.CodeKey)))
            {
                plan.LedgerOnlyCustomers.Add(ledgerRow.CustomerName + "（" + ledgerRow.Code + "）");
                AddIssue(
                    plan,
                    ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower,
                    "台账客户本月无电量",
                    "提示",
                    ledgerRow.CustomerName,
                    "该台账客户不在当月零售结算明细中；继续后会把目标月总量、峰、平、谷写为0，并保留原有峰平谷系数。");
            }

            foreach (var sameNameGroup in powerData.CustomerRows
                .GroupBy(row => row.CustomerNameKey, StringComparer.Ordinal)
                .Where(group => group.Count() > 1))
            {
                plan.AliasCandidateRows++;
                AddIssue(
                    plan,
                    ProvinceStage1LedgerUpdateIssueKinds.PossibleAlias,
                    "电量表同名不同编码",
                    "警告",
                    sameNameGroup.First().CustomerName,
                    "电量表中同一归一化名称对应多个用户编号（"
                    + string.Join("、", sameNameGroup.Select(row => row.Code))
                    + "）；程序将保留为独立客户，不会自动合并。");
            }

            return plan;
        }

        private static IXLWorksheet FindLedgerWorksheet(XLWorkbook workbook)
        {
            var exact = workbook.Worksheets
                .FirstOrDefault(item => string.Equals(item.Name, LedgerSheetName, StringComparison.Ordinal));
            if (exact != null)
            {
                return exact;
            }

            var candidates = workbook.Worksheets
                .Where(item => FindHeaderColumn(item, "电力用户编码") > 0
                    && FindHeaderColumn(item, "电力用户名称") > 0)
                .ToList();
            if (candidates.Count == 1)
            {
                return candidates[0];
            }

            throw new InvalidOperationException(
                "广东台账中未唯一找到主 sheet“"
                + LedgerSheetName
                + "”或同时包含“电力用户编码/电力用户名称”的候选 sheet。");
        }

        private static FixedMap FindFixedMap(IXLWorksheet worksheet)
        {
            var requiredHeaders = new[]
            {
                "序号",
                "电力用户编码",
                "电力用户名称",
                "履约开始月份"
            };
            var lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 1, 10);
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var row = 1; row <= lastRow; row++)
            {
                var columns = requiredHeaders.ToDictionary(
                    header => header,
                    header => new List<int>(),
                    StringComparer.Ordinal);
                for (var column = 1; column <= lastColumn; column++)
                {
                    var text = CellText(worksheet.Cell(row, column));
                    List<int> matches;
                    if (columns.TryGetValue(text, out matches))
                    {
                        matches.Add(column);
                    }
                }

                if (!requiredHeaders.All(header => columns[header].Count > 0))
                {
                    continue;
                }

                var duplicate = requiredHeaders.FirstOrDefault(header => columns[header].Count > 1);
                if (duplicate != null)
                {
                    throw new InvalidOperationException("广东台账固定资料区存在重复表头“" + duplicate + "”。");
                }

                return new FixedMap
                {
                    HeaderRow = row,
                    DataStartRow = Math.Max(4, row + 1),
                    SequenceColumn = columns["序号"][0],
                    CodeColumn = columns["电力用户编码"][0],
                    CustomerNameColumn = columns["电力用户名称"][0],
                    PerformanceStartMonthColumn = columns["履约开始月份"][0]
                };
            }

            throw new InvalidOperationException(
                "广东台账缺少固定资料表头：序号、电力用户编码、电力用户名称、履约开始月份。");
        }

        private static MonthMap ResolveMonthMap(
            IXLWorksheet worksheet,
            int month,
            bool createTargetMonth,
            IList<string> warnings)
        {
            int targetStart;
            if (TryFindMonthStartColumn(worksheet, month, out targetStart))
            {
                ValidateMonthBlock(worksheet, month, targetStart);
                return MonthMap.Create(month, targetStart, true);
            }

            if (month <= 1)
            {
                throw new InvalidOperationException("广东台账中未找到1月区块，且无法从上月复制。");
            }

            var sourceMonth = month - 1;
            var sourceStart = FindMonthStartColumn(worksheet, sourceMonth);
            ValidateMonthBlock(worksheet, sourceMonth, sourceStart);
            var sourceWidth = ExpectedMonthBlockWidth(sourceMonth);
            var targetWidth = ExpectedMonthBlockWidth(month);
            if (sourceWidth != targetWidth)
            {
                throw new InvalidOperationException(
                    "广东台账"
                    + sourceMonth.ToString(CultureInfo.InvariantCulture)
                    + "月与"
                    + month.ToString(CultureInfo.InvariantCulture)
                    + "月区块宽度不同，不能自动跨越历史非标准月份复制；请提供已包含目标月份结构的台账。");
            }

            targetStart = sourceStart + sourceWidth;
            EnsureTargetAreaIsEmpty(worksheet, month, targetStart, targetWidth);
            if (createTargetMonth)
            {
                CopyMonthBlock(worksheet, sourceStart, targetStart, targetWidth, month);
                ValidateMonthBlock(worksheet, month, targetStart);
            }

            warnings.Add(
                "广东台账中未找到"
                + month.ToString(CultureInfo.InvariantCulture)
                + "月区块，"
                + (createTargetMonth ? "已" : "将")
                + "基于"
                + sourceMonth.ToString(CultureInfo.InvariantCulture)
                + "月的"
                + targetWidth.ToString(CultureInfo.InvariantCulture)
                + "列结构创建目标月份区块。");
            return MonthMap.Create(month, targetStart, false);
        }

        private static void ValidateMonthBlock(IXLWorksheet worksheet, int month, int start)
        {
            var width = ExpectedMonthBlockWidth(month);
            var duplicateLabels = FindMonthStartColumns(worksheet, month);
            if (duplicateLabels.Count != 1)
            {
                throw new InvalidOperationException(
                    "广东台账第1行存在"
                    + duplicateLabels.Count.ToString(CultureInfo.InvariantCulture)
                    + "个“"
                    + month.ToString(CultureInfo.InvariantCulture)
                    + "月”标题，无法唯一识别目标月份区块。");
            }

            for (var column = start + 1; column < start + width; column++)
            {
                if (IsMonthLabel(CellText(worksheet.Cell(1, column))))
                {
                    throw new InvalidOperationException(
                        "广东台账"
                        + month.ToString(CultureInfo.InvariantCulture)
                        + "月"
                        + width.ToString(CultureInfo.InvariantCulture)
                        + "列区块内部出现了另一个月份标题，结构不安全。");
                }
            }

            int previousStart;
            if (month > 1 && TryFindMonthStartColumn(worksheet, month - 1, out previousStart))
            {
                var expectedStart = previousStart + ExpectedMonthBlockWidth(month - 1);
                if (start != expectedStart)
                {
                    throw new InvalidOperationException(
                        "广东台账"
                        + month.ToString(CultureInfo.InvariantCulture)
                        + "月区块起始列与上月实际宽度不连续。");
                }
            }

            int nextStart;
            if (month < 12 && TryFindMonthStartColumn(worksheet, month + 1, out nextStart)
                && nextStart != start + width)
            {
                throw new InvalidOperationException(
                    "广东台账"
                    + month.ToString(CultureInfo.InvariantCulture)
                    + "月区块宽度不是预期的"
                    + width.ToString(CultureInfo.InvariantCulture)
                    + "列。");
            }

            if (!IsTotalHeader(HeaderText(worksheet, start))
                || !IsPeriodHeader(HeaderText(worksheet, start + 1), "峰")
                || !IsPeriodHeader(HeaderText(worksheet, start + 2), "平")
                || !IsPeriodHeader(HeaderText(worksheet, start + 3), "谷")
                || !IsCoefficientHeader(HeaderText(worksheet, start + 4), "峰_平")
                || !IsCoefficientHeader(HeaderText(worksheet, start + 5), "谷_平"))
            {
                throw new InvalidOperationException(
                    "广东台账"
                    + month.ToString(CultureInfo.InvariantCulture)
                    + "月区块前6列表头不符合预期：应依次为总实际电量、峰、平、谷、峰_平、谷_平。");
            }
        }

        private static string HeaderText(IXLWorksheet worksheet, int column)
        {
            var row3 = CellText(worksheet.Cell(3, column));
            return row3.Length > 0 ? row3 : CellText(worksheet.Cell(2, column));
        }

        private static bool IsTotalHeader(string value)
        {
            var key = TextUtil.CustomerKey(value);
            return key == "总实际用电量"
                || key == "总实际电量"
                || key == "总实际电量(兆瓦时)"
                || key == "总实际用电量(兆瓦时)";
        }

        private static bool IsPeriodHeader(string value, string period)
        {
            var key = TextUtil.CustomerKey(value);
            return key == period || key == period + "电量";
        }

        private static bool IsCoefficientHeader(string value, string expected)
        {
            return string.Equals(TextUtil.CustomerKey(value), expected, StringComparison.Ordinal);
        }

        private static int ExpectedMonthBlockWidth(int month)
        {
            if (month == 1)
            {
                return 30;
            }

            if (month == 2)
            {
                return 33;
            }

            return 32;
        }

        private static void EnsureTargetAreaIsEmpty(
            IXLWorksheet worksheet,
            int month,
            int targetStart,
            int width)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            if (targetStart <= lastColumn
                && HasContent(
                    worksheet,
                    1,
                    targetStart,
                    lastRow,
                    Math.Min(lastColumn, targetStart + width - 1)))
            {
                throw new InvalidOperationException(
                    "广东台账"
                    + month.ToString(CultureInfo.InvariantCulture)
                    + "月目标区块位置已有内容，但未识别为该月份，已停止以避免覆盖。");
            }
        }

        private static void CopyMonthBlock(
            IXLWorksheet worksheet,
            int sourceStart,
            int targetStart,
            int width,
            int targetMonth)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            var mergedAreas = worksheet.MergedRanges
                .Where(range =>
                    range.RangeAddress.FirstAddress.ColumnNumber >= sourceStart
                    && range.RangeAddress.LastAddress.ColumnNumber < sourceStart + width)
                .Select(range => new MergeArea
                {
                    FirstRow = range.RangeAddress.FirstAddress.RowNumber,
                    LastRow = range.RangeAddress.LastAddress.RowNumber,
                    FirstColumnOffset = range.RangeAddress.FirstAddress.ColumnNumber - sourceStart,
                    LastColumnOffset = range.RangeAddress.LastAddress.ColumnNumber - sourceStart
                })
                .ToList();

            worksheet.Range(1, sourceStart, lastRow, sourceStart + width - 1)
                .CopyTo(worksheet.Cell(1, targetStart));
            foreach (var area in mergedAreas)
            {
                var targetRange = worksheet.Range(
                    area.FirstRow,
                    targetStart + area.FirstColumnOffset,
                    area.LastRow,
                    targetStart + area.LastColumnOffset);
                if (!worksheet.MergedRanges.Any(range =>
                    range.RangeAddress.FirstAddress.RowNumber == area.FirstRow
                    && range.RangeAddress.LastAddress.RowNumber == area.LastRow
                    && range.RangeAddress.FirstAddress.ColumnNumber == targetStart + area.FirstColumnOffset
                    && range.RangeAddress.LastAddress.ColumnNumber == targetStart + area.LastColumnOffset))
                {
                    targetRange.Merge();
                }
            }

            worksheet.Cell(1, targetStart).Value =
                targetMonth.ToString(CultureInfo.InvariantCulture) + "月";
            for (var offset = 0; offset < width; offset++)
            {
                var sourceColumn = worksheet.Column(sourceStart + offset);
                var targetColumn = worksheet.Column(targetStart + offset);
                targetColumn.Width = sourceColumn.Width;
                if (sourceColumn.IsHidden)
                {
                    targetColumn.Hide();
                }
                else
                {
                    targetColumn.Unhide();
                }
            }
        }

        private static List<LedgerCustomerRow> ReadLedgerRows(IXLWorksheet worksheet, FixedMap map)
        {
            var rows = new List<LedgerCustomerRow>();
            var invalidRows = new List<string>();
            var codes = new Dictionary<string, int>(StringComparer.Ordinal);
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? map.DataStartRow;
            for (var row = map.DataStartRow; row <= lastRow; row++)
            {
                var codeText = CellText(worksheet.Cell(row, map.CodeColumn));
                var name = CellText(worksheet.Cell(row, map.CustomerNameColumn));
                var sequenceText = CellText(worksheet.Cell(row, map.SequenceColumn));
                if (codeText.Length == 0 && name.Length == 0 && sequenceText.Length == 0)
                {
                    continue;
                }

                decimal sequenceValue;
                var hasPositiveSequence = TryCellNumber(worksheet.Cell(row, map.SequenceColumn), out sequenceValue)
                    && sequenceValue > 0m
                    && sequenceValue == decimal.Truncate(sequenceValue);
                if (!hasPositiveSequence)
                {
                    if (codeText.Length == 0)
                    {
                        continue;
                    }

                    invalidRows.Add("第" + row.ToString(CultureInfo.InvariantCulture) + "行电力用户序号为空或不是正整数");
                    continue;
                }

                var codeKey = TextUtil.CustomerKey(codeText);
                var nameKey = TextUtil.CustomerKey(name);
                if (codeKey.Length == 0)
                {
                    invalidRows.Add("第" + row.ToString(CultureInfo.InvariantCulture) + "行电力用户编码为空");
                    continue;
                }

                if (nameKey.Length == 0)
                {
                    invalidRows.Add("第" + row.ToString(CultureInfo.InvariantCulture) + "行电力用户名称为空");
                    continue;
                }

                int existingRow;
                if (codes.TryGetValue(codeKey, out existingRow))
                {
                    invalidRows.Add(
                        "电力用户编码“"
                        + codeKey
                        + "”在第"
                        + existingRow.ToString(CultureInfo.InvariantCulture)
                        + "行和第"
                        + row.ToString(CultureInfo.InvariantCulture)
                        + "行重复");
                    continue;
                }

                codes.Add(codeKey, row);
                rows.Add(new LedgerCustomerRow
                {
                    RowNumber = row,
                    Sequence = decimal.ToInt32(sequenceValue),
                    Code = codeText,
                    CodeKey = codeKey,
                    CustomerName = name,
                    CustomerNameKey = nameKey
                });
            }

            if (invalidRows.Count > 0)
            {
                throw new InvalidOperationException(
                    "广东台账客户固定资料存在严重错误，已停止生成。"
                    + Environment.NewLine
                    + string.Join(Environment.NewLine, invalidRows.Take(10)));
            }

            if (rows.Count == 0)
            {
                throw new InvalidOperationException("广东台账没有可识别的客户数据行，无法安全选择新增行模板。");
            }

            return rows;
        }

        private static Dictionary<string, CoefficientSnapshot> SnapshotCoefficients(
            IXLWorksheet worksheet,
            MonthMap monthMap,
            IEnumerable<LedgerCustomerRow> ledgerRows)
        {
            return ledgerRows.ToDictionary(
                row => row.CodeKey,
                row => new CoefficientSnapshot
                {
                    PeakFlat = SnapshotCell(worksheet.Cell(row.RowNumber, monthMap.PeakFlatCoefficientColumn)),
                    ValleyFlat = SnapshotCell(worksheet.Cell(row.RowNumber, monthMap.ValleyFlatCoefficientColumn))
                },
                StringComparer.Ordinal);
        }

        private static CellSnapshot SnapshotCell(IXLCell cell)
        {
            return new CellSnapshot
            {
                IsEmpty = cell.IsEmpty(),
                FormulaA1 = cell.FormulaA1,
                Value = cell.Value
            };
        }

        private static void ClearTargetMonthValues(
            IXLWorksheet worksheet,
            MonthMap monthMap,
            IEnumerable<LedgerCustomerRow> ledgerRows)
        {
            foreach (var ledgerRow in ledgerRows)
            {
                worksheet.Range(
                        ledgerRow.RowNumber,
                        monthMap.TotalColumn,
                        ledgerRow.RowNumber,
                        monthMap.ValleyFlatCoefficientColumn)
                    .Clear(XLClearOptions.Contents);
                worksheet.Cell(ledgerRow.RowNumber, monthMap.TotalColumn).SetValue(0m);
                worksheet.Cell(ledgerRow.RowNumber, monthMap.PeakColumn).SetValue(0m);
                worksheet.Cell(ledgerRow.RowNumber, monthMap.FlatColumn).SetValue(0m);
                worksheet.Cell(ledgerRow.RowNumber, monthMap.ValleyColumn).SetValue(0m);
            }
        }

        private static void RestoreCoefficients(
            IXLWorksheet worksheet,
            MonthMap monthMap,
            int row,
            CoefficientSnapshot snapshot)
        {
            RestoreCell(worksheet.Cell(row, monthMap.PeakFlatCoefficientColumn), snapshot.PeakFlat);
            RestoreCell(worksheet.Cell(row, monthMap.ValleyFlatCoefficientColumn), snapshot.ValleyFlat);
        }

        private static void RestoreCell(IXLCell cell, CellSnapshot snapshot)
        {
            if (snapshot.IsEmpty)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(snapshot.FormulaA1))
            {
                cell.FormulaA1 = snapshot.FormulaA1;
            }
            else
            {
                cell.Value = snapshot.Value;
            }
        }

        private static void WritePower(
            IXLWorksheet worksheet,
            MonthMap map,
            int row,
            GuangdongStage1SourceReader.GuangdongPowerAggregateRow power)
        {
            worksheet.Cell(row, map.TotalColumn).SetValue(power.Total);
            worksheet.Cell(row, map.PeakColumn).SetValue(power.Peak);
            worksheet.Cell(row, map.FlatColumn).SetValue(power.Flat);
            worksheet.Cell(row, map.ValleyColumn).SetValue(power.Valley);
        }

        private static void WriteCoefficients(
            IXLWorksheet worksheet,
            MonthMap map,
            int row,
            GuangdongStage1SourceReader.GuangdongPowerAggregateRow power)
        {
            worksheet.Cell(row, map.PeakFlatCoefficientColumn).SetValue(power.PeakFlatCoefficient);
            worksheet.Cell(row, map.ValleyFlatCoefficientColumn).SetValue(power.ValleyFlatCoefficient);
        }

        private static void InsertNewCustomerRow(
            IXLWorksheet worksheet,
            FixedMap fixedMap,
            MonthMap monthMap,
            IList<LedgerCustomerRow> ledgerRows,
            GuangdongStage1SourceReader.GuangdongPowerAggregateRow power,
            int month)
        {
            var templateRow = ledgerRows.Max(row => row.RowNumber);
            var newRow = templateRow + 1;
            var lastColumn = Math.Max(
                worksheet.LastColumnUsed()?.ColumnNumber() ?? monthMap.ValleyFlatCoefficientColumn,
                monthMap.StartColumn + monthMap.Width - 1);
            worksheet.Row(templateRow).InsertRowsBelow(1);
            worksheet.Range(templateRow, 1, templateRow, lastColumn)
                .CopyTo(worksheet.Cell(newRow, 1));

            for (var column = 1; column <= lastColumn; column++)
            {
                var cell = worksheet.Cell(newRow, column);
                if (string.IsNullOrWhiteSpace(cell.FormulaA1))
                {
                    cell.Clear(XLClearOptions.Contents);
                }
            }

            var sequence = ledgerRows.Max(row => row.Sequence) + 1;
            worksheet.Cell(newRow, fixedMap.SequenceColumn).Value = sequence;
            worksheet.Cell(newRow, fixedMap.CodeColumn).Value = power.Code;
            worksheet.Cell(newRow, fixedMap.CodeColumn).Style.NumberFormat.Format = "@";
            worksheet.Cell(newRow, fixedMap.CustomerNameColumn).Value = power.CustomerName;
            worksheet.Cell(newRow, fixedMap.PerformanceStartMonthColumn).Value = SettlementYear * 100 + month;
            WritePower(worksheet, monthMap, newRow, power);
            if (power.HasCoefficientPair)
            {
                WriteCoefficients(worksheet, monthMap, newRow, power);
            }

            ledgerRows.Add(new LedgerCustomerRow
            {
                RowNumber = newRow,
                Sequence = sequence,
                Code = power.Code,
                CodeKey = power.Code,
                CustomerName = power.CustomerName,
                CustomerNameKey = power.CustomerNameKey
            });
        }

        private static void VerifyWrittenValues(
            IXLWorksheet worksheet,
            MonthMap monthMap,
            IEnumerable<LedgerCustomerRow> ledgerRows,
            IDictionary<string, GuangdongStage1SourceReader.GuangdongPowerAggregateRow> powerRowsByCode)
        {
            var ledgerRowsByCode = ledgerRows.ToDictionary(row => row.CodeKey, StringComparer.Ordinal);
            decimal writtenTotal = 0m;
            foreach (var powerRow in powerRowsByCode.Values)
            {
                LedgerCustomerRow ledgerRow;
                if (!ledgerRowsByCode.TryGetValue(powerRow.Code, out ledgerRow))
                {
                    throw new InvalidOperationException("广东台账写入后未找到用户编号“" + powerRow.Code + "”。");
                }

                if (!SamePowerVector(worksheet, ledgerRow.RowNumber, monthMap, powerRow))
                {
                    throw new InvalidOperationException(
                        "广东台账写入后用户“"
                        + powerRow.CustomerName
                        + "”（"
                        + powerRow.Code
                        + "）电量复核失败，未保存输出。");
                }

                if (powerRow.HasCoefficientPair
                    && !SameCoefficients(worksheet, ledgerRow.RowNumber, monthMap, powerRow))
                {
                    throw new InvalidOperationException(
                        "广东台账写入后用户“"
                        + powerRow.CustomerName
                        + "”（"
                        + powerRow.Code
                        + "）峰平谷系数复核失败，未保存输出。");
                }

                writtenTotal += powerRow.Total;
            }

            var sourceTotal = powerRowsByCode.Values.Sum(row => row.Total);
            if (Math.Abs(writtenTotal - sourceTotal) > GuangdongStage1SourceReader.PowerTolerance)
            {
                throw new InvalidOperationException("广东台账写入后全月总电量复核失败，未保存输出。");
            }

            foreach (var ledgerRow in ledgerRows.Where(row => !powerRowsByCode.ContainsKey(row.CodeKey)))
            {
                if (!IsZeroPower(worksheet, ledgerRow.RowNumber, monthMap))
                {
                    throw new InvalidOperationException(
                        "广东台账写入后，本月无电量客户“"
                        + ledgerRow.CustomerName
                        + "”的目标月电量不是0，未保存输出。");
                }
            }
        }

        private static bool SamePowerVector(
            IXLWorksheet worksheet,
            int row,
            MonthMap map,
            GuangdongStage1SourceReader.GuangdongPowerAggregateRow power)
        {
            decimal total;
            decimal peak;
            decimal flat;
            decimal valley;
            return TryCellNumber(worksheet.Cell(row, map.TotalColumn), out total)
                && TryCellNumber(worksheet.Cell(row, map.PeakColumn), out peak)
                && TryCellNumber(worksheet.Cell(row, map.FlatColumn), out flat)
                && TryCellNumber(worksheet.Cell(row, map.ValleyColumn), out valley)
                && SameNumber(total, power.Total)
                && SameNumber(peak, power.Peak)
                && SameNumber(flat, power.Flat)
                && SameNumber(valley, power.Valley);
        }

        private static bool SameCoefficients(
            IXLWorksheet worksheet,
            int row,
            MonthMap map,
            GuangdongStage1SourceReader.GuangdongPowerAggregateRow power)
        {
            decimal peakFlat;
            decimal valleyFlat;
            return TryCellNumber(worksheet.Cell(row, map.PeakFlatCoefficientColumn), out peakFlat)
                && TryCellNumber(worksheet.Cell(row, map.ValleyFlatCoefficientColumn), out valleyFlat)
                && SameNumber(peakFlat, power.PeakFlatCoefficient)
                && SameNumber(valleyFlat, power.ValleyFlatCoefficient);
        }

        private static bool IsBlankPower(IXLWorksheet worksheet, int row, MonthMap map)
        {
            return worksheet.Cell(row, map.TotalColumn).IsEmpty()
                && worksheet.Cell(row, map.PeakColumn).IsEmpty()
                && worksheet.Cell(row, map.FlatColumn).IsEmpty()
                && worksheet.Cell(row, map.ValleyColumn).IsEmpty();
        }

        private static bool IsZeroPower(IXLWorksheet worksheet, int row, MonthMap map)
        {
            decimal total;
            decimal peak;
            decimal flat;
            decimal valley;
            return TryCellNumber(worksheet.Cell(row, map.TotalColumn), out total)
                && TryCellNumber(worksheet.Cell(row, map.PeakColumn), out peak)
                && TryCellNumber(worksheet.Cell(row, map.FlatColumn), out flat)
                && TryCellNumber(worksheet.Cell(row, map.ValleyColumn), out valley)
                && SameNumber(total, 0m)
                && SameNumber(peak, 0m)
                && SameNumber(flat, 0m)
                && SameNumber(valley, 0m);
        }

        private static bool TryCellNumber(IXLCell cell, out decimal value)
        {
            if (cell.IsEmpty())
            {
                value = 0m;
                return false;
            }

            if (cell.TryGetValue(out value))
            {
                return true;
            }

            return decimal.TryParse(
                CellText(cell).Replace(",", string.Empty),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out value);
        }

        private static bool SameNumber(decimal left, decimal right)
        {
            return Math.Abs(left - right) <= GuangdongStage1SourceReader.PowerTolerance;
        }

        private static int FindHeaderColumn(IXLWorksheet worksheet, string headerText)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            var lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 1, 10);
            for (var row = 1; row <= lastRow; row++)
            {
                for (var column = 1; column <= lastColumn; column++)
                {
                    if (CellText(worksheet.Cell(row, column)) == headerText)
                    {
                        return column;
                    }
                }
            }

            return 0;
        }

        private static int FindMonthStartColumn(IXLWorksheet worksheet, int month)
        {
            int column;
            if (TryFindMonthStartColumn(worksheet, month, out column))
            {
                return column;
            }

            throw new InvalidOperationException(
                "广东台账中未找到"
                + month.ToString(CultureInfo.InvariantCulture)
                + "月区块。");
        }

        private static bool TryFindMonthStartColumn(IXLWorksheet worksheet, int month, out int column)
        {
            var columns = FindMonthStartColumns(worksheet, month);
            if (columns.Count > 1)
            {
                throw new InvalidOperationException(
                    "广东台账第1行存在多个“"
                    + month.ToString(CultureInfo.InvariantCulture)
                    + "月”标题，无法唯一识别月份区块。");
            }

            if (columns.Count == 1)
            {
                column = columns[0];
                return true;
            }

            column = 0;
            return false;
        }

        private static List<int> FindMonthStartColumns(IXLWorksheet worksheet, int month)
        {
            var label = month.ToString(CultureInfo.InvariantCulture) + "月";
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            var columns = new List<int>();
            for (var column = 1; column <= lastColumn; column++)
            {
                if (CellText(worksheet.Cell(1, column)) == label)
                {
                    columns.Add(column);
                }
            }

            return columns;
        }

        private static bool IsMonthLabel(string value)
        {
            if (!value.EndsWith("月", StringComparison.Ordinal))
            {
                return false;
            }

            int month;
            return int.TryParse(
                    value.Substring(0, value.Length - 1),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out month)
                && month >= 1
                && month <= 12;
        }

        private static bool HasContent(
            IXLWorksheet worksheet,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn)
        {
            for (var row = firstRow; row <= lastRow; row++)
            {
                for (var column = firstColumn; column <= lastColumn; column++)
                {
                    if (!worksheet.Cell(row, column).IsEmpty())
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static void AddIssue(
            ProvinceStage1LedgerUpdatePlan plan,
            string kind,
            string category,
            string severity,
            string customerName,
            string message)
        {
            plan.Issues.Add(new ProvinceStage1LedgerUpdateIssue
            {
                Kind = kind,
                Category = category,
                Severity = severity,
                CustomerName = customerName,
                Message = message
            });
        }

        private static void RejectUnsupportedCustomerDecisions(ProvinceStage1LedgerUpdateOptions options)
        {
            if ((options.CustomerDecisions != null && options.CustomerDecisions.Count > 0)
                || (options.ManualCustomerMatches != null && options.ManualCustomerMatches.Count > 0))
            {
                throw new InvalidOperationException(
                    "广东阶段一只按电力用户编码精确匹配；不接受基于客户名称的人工匹配或跳过决定。");
            }
        }

        private static void EnsureOutputDoesNotOverwriteInput(
            string outputPath,
            params string[] inputPaths)
        {
            var normalizedOutput = Path.GetFullPath(outputPath);
            if (inputPaths.Any(path => string.Equals(
                normalizedOutput,
                Path.GetFullPath(path),
                StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException("广东台账更新输出路径不能与台账或零售结算明细输入文件相同。");
            }
        }

        private static string CellText(IXLCell cell)
        {
            return TextUtil.S(cell.GetFormattedString());
        }

        private static string UniquePath(string path)
        {
            if (!File.Exists(path))
            {
                return path;
            }

            var directory = Path.GetDirectoryName(path);
            var stem = Path.GetFileNameWithoutExtension(path);
            var extension = Path.GetExtension(path);
            var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
            var candidate = Path.Combine(directory, stem + "-" + timestamp + extension);
            var index = 1;
            while (File.Exists(candidate))
            {
                candidate = Path.Combine(directory, stem + "-" + timestamp + "-" + index + extension);
                index++;
            }

            return candidate;
        }

        private sealed class LedgerUpdateContext
        {
            public IXLWorksheet Worksheet { get; set; }
            public FixedMap FixedMap { get; set; }
            public MonthMap MonthMap { get; set; }
            public List<LedgerCustomerRow> LedgerRows { get; set; }
            public GuangdongStage1SourceReader.GuangdongStage1DataSet PowerData { get; set; }
            public Dictionary<string, GuangdongStage1SourceReader.GuangdongPowerAggregateRow> PowerRowsByCode { get; set; }
            public List<GuangdongStage1SourceReader.GuangdongPowerAggregateRow> NewCustomerRows { get; set; }
            public ProvinceStage1LedgerUpdatePlan Plan { get; set; }
        }

        private sealed class FixedMap
        {
            public int HeaderRow { get; set; }
            public int DataStartRow { get; set; }
            public int SequenceColumn { get; set; }
            public int CodeColumn { get; set; }
            public int CustomerNameColumn { get; set; }
            public int PerformanceStartMonthColumn { get; set; }
        }

        private sealed class MonthMap
        {
            public int Month { get; private set; }
            public int StartColumn { get; private set; }
            public int Width { get; private set; }
            public bool TargetMonthAlreadyPresent { get; private set; }
            public int TotalColumn => StartColumn;
            public int PeakColumn => StartColumn + 1;
            public int FlatColumn => StartColumn + 2;
            public int ValleyColumn => StartColumn + 3;
            public int PeakFlatCoefficientColumn => StartColumn + PowerColumnCount;
            public int ValleyFlatCoefficientColumn => StartColumn + PowerColumnCount + 1;

            public static MonthMap Create(int month, int startColumn, bool alreadyPresent)
            {
                return new MonthMap
                {
                    Month = month,
                    StartColumn = startColumn,
                    Width = ExpectedMonthBlockWidth(month),
                    TargetMonthAlreadyPresent = alreadyPresent
                };
            }
        }

        private sealed class LedgerCustomerRow
        {
            public int RowNumber { get; set; }
            public int Sequence { get; set; }
            public string Code { get; set; }
            public string CodeKey { get; set; }
            public string CustomerName { get; set; }
            public string CustomerNameKey { get; set; }
        }

        private sealed class CoefficientSnapshot
        {
            public CellSnapshot PeakFlat { get; set; }
            public CellSnapshot ValleyFlat { get; set; }
        }

        private sealed class CellSnapshot
        {
            public bool IsEmpty { get; set; }
            public string FormulaA1 { get; set; }
            public XLCellValue Value { get; set; }
        }

        private sealed class MergeArea
        {
            public int FirstRow { get; set; }
            public int LastRow { get; set; }
            public int FirstColumnOffset { get; set; }
            public int LastColumnOffset { get; set; }
        }
    }
}
