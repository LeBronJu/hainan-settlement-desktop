using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal sealed class ChongqingLedgerStage1Updater
    {
        private const string Unit = "兆瓦时";
        private const int MonthBlockWidth = 30;
        private const int MonthPowerColumnCount = 5;
        private readonly ChongqingPowerCleanGenerator _powerCleanGenerator;

        public ChongqingLedgerStage1Updater(ChongqingPowerCleanGenerator powerCleanGenerator)
        {
            _powerCleanGenerator = powerCleanGenerator ?? throw new ArgumentNullException(nameof(powerCleanGenerator));
        }

        public ProvinceStage1LedgerUpdatePlan Plan(ProvinceStage1LedgerUpdateOptions options)
        {
            using (var stream = File.Open(options.LedgerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                return BuildContext(workbook, options).Plan;
            }
        }

        public ProvinceStage1LedgerUpdateResult Update(ProvinceStage1LedgerUpdateOptions options)
        {
            using (var stream = File.Open(options.LedgerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var context = BuildContext(workbook, options);
                var outputPath = UniquePath(Path.Combine(options.OutputDirectory, options.Month + "月重庆售电结算台账-阶段一更新.xlsx"));
                var reportPath = UniquePath(Path.Combine(options.OutputDirectory, options.Month + "月重庆阶段一台账更新报告.json"));
                FileAccessGuard.RequireWritableWorkbook(outputPath, "重庆台账更新输出文件");

                var updatedRows = 0;
                foreach (var ledgerRow in context.LedgerRows)
                {
                    ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow powerRow;
                    if (!context.PowerRowsByLedgerKey.TryGetValue(ledgerRow.Key, out powerRow))
                    {
                        continue;
                    }

                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.TotalColumn).Value = powerRow.Total;
                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.SharpColumn).Value = powerRow.Sharp;
                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.PeakColumn).Value = powerRow.Peak;
                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.FlatColumn).Value = powerRow.Flat;
                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.ValleyColumn).Value = powerRow.Valley;
                    updatedRows++;
                }

                foreach (var newCustomer in context.NewCustomerRows)
                {
                    InsertNewCustomerRow(context.Worksheet, context.Map, context.LedgerRows, newCustomer);
                    updatedRows++;
                }

                workbook.SaveAs(outputPath);

                var result = new ProvinceStage1LedgerUpdateResult
                {
                    Province = ProvinceCode.Chongqing,
                    Month = options.Month,
                    Unit = Unit,
                    LedgerPath = options.LedgerPath,
                    RawDetailPath = options.RawDetailPath,
                    OutputLedgerPath = outputPath,
                    ReportPath = reportPath,
                    LedgerCustomerRows = context.Plan.LedgerCustomerRows,
                    PowerCustomerRows = context.Plan.PowerCustomerRows,
                    MatchedRows = context.Plan.MatchedRows,
                    UpdatedPowerRows = updatedRows,
                    ManualMatchedRows = context.Plan.ManualMatchedRows,
                    CreatedCustomerRows = context.Plan.CreatedCustomerRows,
                    SkippedCustomerRows = context.Plan.SkippedCustomerRows,
                    MultiAccountRows = context.Plan.MultiAccountRows,
                    SkippedRows = context.Plan.MissingInLedgerRows + context.Plan.SkippedCustomerRows,
                    TotalPower = Math.Round(context.PowerData.CustomerRows.Sum(row => row.Total), 4),
                    CustomerDecisions = context.Plan.CustomerDecisions,
                    ManualCustomerMatches = context.Plan.ManualCustomerMatches,
                    Warnings = context.Plan.Warnings,
                    Issues = context.Plan.Issues
                };
                WriteReport(reportPath, result);
                return result;
            }
        }

        private LedgerUpdateContext BuildContext(XLWorkbook workbook, ProvinceStage1LedgerUpdateOptions options)
        {
            var powerData = _powerCleanGenerator.ReadData(new ProvinceStage1CleanOptions
            {
                Province = ProvinceCode.Chongqing,
                Month = options.Month,
                RawDetailPath = options.RawDetailPath,
                OutputDirectory = options.OutputDirectory
            });
            var worksheet = FindLedgerWorksheet(workbook);
            var ledgerWarnings = new List<string>();
            var map = FindLedgerMap(worksheet, options.Month, ledgerWarnings);
            var ledgerRows = ReadLedgerRows(worksheet, map);
            var powerRowsByKey = powerData.CustomerRows.ToDictionary(row => TextUtil.CustomerKey(row.CustomerName));
            var ledgerRowsByKey = ledgerRows.ToDictionary(row => row.Key);
            var accountNumbersByKey = powerData.AccountRows
                .GroupBy(row => TextUtil.CustomerKey(row.CustomerName))
                .ToDictionary(
                    group => group.Key,
                    group => group
                        .Select(row => TextUtil.S(row.AccountNumber))
                        .Where(value => value.Length > 0)
                        .Distinct(StringComparer.Ordinal)
                        .OrderBy(value => value, StringComparer.Ordinal)
                        .ToList());
            var exactMatchedKeys = ledgerRowsByKey.Keys.Intersect(powerRowsByKey.Keys).ToList();
            var customerDecisions = ResolveCustomerDecisions(options, ledgerRowsByKey, powerRowsByKey, exactMatchedKeys);
            var manualMatches = customerDecisions
                .Where(decision => decision.DecisionKind == ProvinceStage1CustomerDecisionKind.MatchExisting)
                .Select(decision => new ProvinceStage1CustomerMatch
                {
                    SourceCustomerName = decision.SourceCustomerName,
                    TargetCustomerName = decision.TargetCustomerName
                })
                .ToList();
            var newCustomerRows = customerDecisions
                .Where(decision => decision.DecisionKind == ProvinceStage1CustomerDecisionKind.CreateNew)
                .Select(decision => new NewCustomerPowerRow
                {
                    CustomerName = decision.SourceCustomerName,
                    PowerRow = powerRowsByKey[TextUtil.CustomerKey(decision.SourceCustomerName)]
                })
                .ToList();
            var powerRowsByLedgerKey = new Dictionary<string, ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow>();
            var powerKeyByLedgerKey = new Dictionary<string, string>();

            foreach (var key in exactMatchedKeys)
            {
                powerRowsByLedgerKey[key] = powerRowsByKey[key];
                powerKeyByLedgerKey[key] = key;
            }

            foreach (var match in manualMatches)
            {
                var sourceKey = TextUtil.CustomerKey(match.SourceCustomerName);
                var targetKey = TextUtil.CustomerKey(match.TargetCustomerName);
                powerRowsByLedgerKey[targetKey] = powerRowsByKey[sourceKey];
                powerKeyByLedgerKey[targetKey] = sourceKey;
            }

            var plan = BuildPlan(
                options,
                powerData,
                worksheet,
                map,
                ledgerRows,
                ledgerRowsByKey,
                powerRowsByKey,
                powerRowsByLedgerKey,
                powerKeyByLedgerKey,
                accountNumbersByKey,
                ledgerWarnings,
                customerDecisions,
                manualMatches);
            return new LedgerUpdateContext
            {
                Worksheet = worksheet,
                Map = map,
                PowerData = powerData,
                LedgerRows = ledgerRows,
                PowerRowsByLedgerKey = powerRowsByLedgerKey,
                NewCustomerRows = newCustomerRows,
                AccountNumbersByKey = accountNumbersByKey,
                Plan = plan
            };
        }

        private static ProvinceStage1LedgerUpdatePlan BuildPlan(
            ProvinceStage1LedgerUpdateOptions options,
            ChongqingPowerCleanGenerator.ChongqingPowerDataSet powerData,
            IXLWorksheet worksheet,
            LedgerMap map,
            IList<LedgerCustomerRow> ledgerRows,
            IDictionary<string, LedgerCustomerRow> ledgerRowsByKey,
            IDictionary<string, ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow> powerRowsByKey,
            IDictionary<string, ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow> powerRowsByLedgerKey,
            IDictionary<string, string> powerKeyByLedgerKey,
            IDictionary<string, List<string>> accountNumbersByKey,
            IList<string> ledgerWarnings,
            IList<ProvinceStage1CustomerDecision> customerDecisions,
            IList<ProvinceStage1CustomerMatch> manualMatches)
        {
            var matchedKeys = powerRowsByLedgerKey.Keys.ToList();
            var matchedPowerKeys = new HashSet<string>(powerKeyByLedgerKey.Values);
            foreach (var decision in customerDecisions)
            {
                matchedPowerKeys.Add(TextUtil.CustomerKey(decision.SourceCustomerName));
            }

            var plan = new ProvinceStage1LedgerUpdatePlan
            {
                Province = ProvinceCode.Chongqing,
                Month = options.Month,
                Unit = Unit,
                LedgerCustomerRows = ledgerRows.Count,
                PowerCustomerRows = powerData.CustomerRows.Count,
                MatchedRows = matchedKeys.Count,
                ManualMatchedRows = manualMatches.Count,
                CreatedCustomerRows = customerDecisions.Count(decision => decision.DecisionKind == ProvinceStage1CustomerDecisionKind.CreateNew),
                SkippedCustomerRows = customerDecisions.Count(decision => decision.DecisionKind == ProvinceStage1CustomerDecisionKind.SkipWrite),
                CustomerDecisions = customerDecisions.ToList(),
                ManualCustomerMatches = manualMatches.ToList(),
                Warnings = powerData.Warnings.Concat(ledgerWarnings).ToList()
            };

            if (powerData.Month > 0 && powerData.Month != options.Month)
            {
                AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.MonthMismatch, "月份不一致", "警告", null, "电量确认单识别为" + powerData.Month + "月，界面选择为" + options.Month + "月。");
            }

            foreach (var key in matchedKeys)
            {
                var ledgerRow = ledgerRowsByKey[key];
                var powerRow = powerRowsByLedgerKey[key];
                var powerKey = powerKeyByLedgerKey[key];
                var accounts = accountNumbersByKey.ContainsKey(powerKey) ? accountNumbersByKey[powerKey] : new List<string>();
                if (accounts.Count > 1)
                {
                    plan.MultiAccountRows++;
                    AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.MultiAccountCustomer, "多户号客户", "提示", ledgerRow.CustomerName, "该客户在电量明细中存在多个户号；本次仅写入汇总电量，不会写入电力用户编码列。");
                }

                if (powerKey != key)
                {
                    AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.ManualMatchedCustomer, "人工匹配客户", "警告", ledgerRow.CustomerName, "电量客户“" + powerRow.CustomerName + "”将按本次人工确认写入该台账客户。");
                }

                if (!IsBlankPower(worksheet, ledgerRow.RowNumber, map) && !SamePowerVector(worksheet, ledgerRow.RowNumber, map, powerRow))
                {
                    plan.ExistingDifferentPowerRows++;
                    AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.ExistingPowerDifference, "已有电量差异", "警告", ledgerRow.CustomerName, "台账目标月份已有电量，且与清洗结果不一致；继续后会按清洗结果写入副本。");
                }
            }

            foreach (var decision in customerDecisions.Where(item => item.DecisionKind == ProvinceStage1CustomerDecisionKind.CreateNew))
            {
                var powerRow = powerRowsByKey[TextUtil.CustomerKey(decision.SourceCustomerName)];
                AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.CreatedCustomer, "新增客户到台账", "提示", powerRow.CustomerName, "该电量客户将新增到本次生成的台账副本，仅写入客户名称和目标月份电量。");
            }

            foreach (var decision in customerDecisions.Where(item => item.DecisionKind == ProvinceStage1CustomerDecisionKind.SkipWrite))
            {
                var powerRow = powerRowsByKey[TextUtil.CustomerKey(decision.SourceCustomerName)];
                AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.SkippedPowerCustomer, "本月不写入", "提示", powerRow.CustomerName, "该电量客户已按本次选择跳过，不会写入台账副本。");
            }

            var missingInLedger = powerRowsByKey.Keys.Except(matchedPowerKeys).ToList();
            foreach (var key in missingInLedger)
            {
                plan.MissingInLedgerRows++;
                plan.PowerOnlyCustomers.Add(powerRowsByKey[key].CustomerName);
                AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.PowerCustomerMissingInLedger, "电量客户不在台账", "警告", powerRowsByKey[key].CustomerName, "清洗结果中的客户在台账中找不到，继续后不会新增该客户行。");
            }

            var missingInPower = ledgerRowsByKey.Keys.Except(matchedKeys).ToList();
            foreach (var key in missingInPower)
            {
                plan.MissingInPowerRows++;
                plan.LedgerOnlyCustomers.Add(ledgerRowsByKey[key].CustomerName);
                AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower, "台账客户不在电量表", "提示", ledgerRowsByKey[key].CustomerName, "台账客户在清洗结果中找不到，继续后该行目标月份电量不会更新。");
            }

            foreach (var powerKey in missingInLedger)
            {
                var powerRow = powerRowsByKey[powerKey];
                foreach (var ledgerKey in missingInPower)
                {
                    if (SamePowerVector(worksheet, ledgerRowsByKey[ledgerKey].RowNumber, map, powerRow))
                    {
                        plan.AliasCandidateRows++;
                        AddIssue(plan, ProvinceStage1LedgerUpdateIssueKinds.PossibleAlias, "疑似名称差异", "警告", powerRow.CustomerName, "该电量客户与某个台账客户目标月份电量完全一致，可能是名称别名；请人工确认。");
                        break;
                    }
                }
            }

            return plan;
        }

        private static List<ProvinceStage1CustomerDecision> ResolveCustomerDecisions(
            ProvinceStage1LedgerUpdateOptions options,
            IDictionary<string, LedgerCustomerRow> ledgerRowsByKey,
            IDictionary<string, ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow> powerRowsByKey,
            IList<string> exactMatchedKeys)
        {
            var decisions = new List<ProvinceStage1CustomerDecision>();
            if (options.CustomerDecisions != null)
            {
                decisions.AddRange(options.CustomerDecisions.Where(decision => decision != null));
            }

            if ((options.CustomerDecisions == null || options.CustomerDecisions.Count == 0)
                && options.ManualCustomerMatches != null)
            {
                decisions.AddRange(options.ManualCustomerMatches
                    .Where(match => match != null)
                    .Select(match => new ProvinceStage1CustomerDecision
                    {
                        SourceCustomerName = match.SourceCustomerName,
                        TargetCustomerName = match.TargetCustomerName,
                        DecisionKind = ProvinceStage1CustomerDecisionKind.MatchExisting
                    }));
            }

            var resolved = new List<ProvinceStage1CustomerDecision>();
            if (decisions.Count == 0)
            {
                return resolved;
            }

            var exactMatched = new HashSet<string>(exactMatchedKeys);
            var sourceKeys = new HashSet<string>();
            var targetKeys = new HashSet<string>();
            foreach (var decision in decisions)
            {
                var sourceName = TextUtil.S(decision.SourceCustomerName);
                if (string.IsNullOrWhiteSpace(sourceName))
                {
                    continue;
                }

                var sourceKey = TextUtil.CustomerKey(sourceName);
                if (!powerRowsByKey.ContainsKey(sourceKey))
                {
                    throw new InvalidOperationException("客户处理决定中的电量客户在清洗结果中不存在：" + sourceName);
                }

                if (exactMatched.Contains(sourceKey))
                {
                    throw new InvalidOperationException("客户处理决定中的电量客户已按名称精确匹配，无需再手动处理：" + sourceName);
                }

                if (!sourceKeys.Add(sourceKey))
                {
                    throw new InvalidOperationException("同一个电量客户不能重复设置处理决定：" + sourceName);
                }

                if (decision.DecisionKind == ProvinceStage1CustomerDecisionKind.MatchExisting)
                {
                    var targetName = TextUtil.S(decision.TargetCustomerName);
                    if (string.IsNullOrWhiteSpace(targetName))
                    {
                        throw new InvalidOperationException("匹配已有台账客户时必须选择台账客户：" + sourceName);
                    }

                    var targetKey = TextUtil.CustomerKey(targetName);
                    if (!ledgerRowsByKey.ContainsKey(targetKey))
                    {
                        throw new InvalidOperationException("客户处理决定中的台账客户不存在：" + targetName);
                    }

                    if (exactMatched.Contains(targetKey))
                    {
                        throw new InvalidOperationException("客户处理决定中的台账客户已按名称精确匹配，不能被其他电量客户覆盖：" + targetName);
                    }

                    if (!targetKeys.Add(targetKey))
                    {
                        throw new InvalidOperationException("同一个台账客户不能被多个电量客户人工匹配：" + targetName);
                    }

                    resolved.Add(new ProvinceStage1CustomerDecision
                    {
                        SourceCustomerName = powerRowsByKey[sourceKey].CustomerName,
                        TargetCustomerName = ledgerRowsByKey[targetKey].CustomerName,
                        DecisionKind = ProvinceStage1CustomerDecisionKind.MatchExisting
                    });
                    continue;
                }

                if (decision.DecisionKind != ProvinceStage1CustomerDecisionKind.CreateNew
                    && decision.DecisionKind != ProvinceStage1CustomerDecisionKind.SkipWrite)
                {
                    throw new InvalidOperationException("不支持的客户处理决定：" + decision.DecisionKind);
                }

                resolved.Add(new ProvinceStage1CustomerDecision
                {
                    SourceCustomerName = powerRowsByKey[sourceKey].CustomerName,
                    DecisionKind = decision.DecisionKind
                });
            }

            return resolved;
        }

        private static void InsertNewCustomerRow(
            IXLWorksheet worksheet,
            LedgerMap map,
            IList<LedgerCustomerRow> ledgerRows,
            NewCustomerPowerRow newCustomer)
        {
            var templateRowNumber = ledgerRows.Count == 0
                ? map.DataStartRow
                : ledgerRows.Max(row => row.RowNumber);
            var newRowNumber = ledgerRows.Count == 0
                ? map.DataStartRow
                : templateRowNumber + 1;
            var lastColumn = Math.Max(worksheet.LastColumnUsed()?.ColumnNumber() ?? map.ValleyColumn, map.ValleyColumn);

            if (ledgerRows.Count == 0)
            {
                worksheet.Row(newRowNumber).InsertRowsAbove(1);
                templateRowNumber = newRowNumber - 1;
            }
            else
            {
                worksheet.Row(templateRowNumber).InsertRowsBelow(1);
            }

            worksheet.Range(templateRowNumber, 1, templateRowNumber, lastColumn)
                .CopyTo(worksheet.Cell(newRowNumber, 1));

            for (var column = 1; column <= lastColumn; column++)
            {
                var cell = worksheet.Cell(newRowNumber, column);
                if (string.IsNullOrWhiteSpace(cell.FormulaA1))
                {
                    cell.Clear(XLClearOptions.Contents);
                }
            }

            worksheet.Cell(newRowNumber, map.CustomerNameColumn).Value = newCustomer.CustomerName;
            WritePower(worksheet, newRowNumber, map, newCustomer.PowerRow);

            ledgerRows.Add(new LedgerCustomerRow
            {
                RowNumber = newRowNumber,
                CustomerName = newCustomer.CustomerName,
                Key = TextUtil.CustomerKey(newCustomer.CustomerName)
            });
        }

        private static void WritePower(
            IXLWorksheet worksheet,
            int row,
            LedgerMap map,
            ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow powerRow)
        {
            worksheet.Cell(row, map.TotalColumn).Value = powerRow.Total;
            worksheet.Cell(row, map.SharpColumn).Value = powerRow.Sharp;
            worksheet.Cell(row, map.PeakColumn).Value = powerRow.Peak;
            worksheet.Cell(row, map.FlatColumn).Value = powerRow.Flat;
            worksheet.Cell(row, map.ValleyColumn).Value = powerRow.Valley;
        }

        private static IXLWorksheet FindLedgerWorksheet(XLWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                if (FindHeaderColumn(worksheet, "电力用户名称") > 0)
                {
                    return worksheet;
                }
            }

            throw new InvalidOperationException("重庆台账中未找到表头“电力用户名称”。");
        }

        private static LedgerMap FindLedgerMap(IXLWorksheet worksheet, int month, IList<string> warnings)
        {
            var customerNameColumn = FindHeaderColumn(worksheet, "电力用户名称");
            int totalColumn;
            if (!TryFindMonthStartColumn(worksheet, month, out totalColumn))
            {
                totalColumn = CopyPreviousMonthBlock(worksheet, month);
                warnings.Add("重庆台账中未找到" + month + "月电量区块，已基于" + (month - 1) + "月电量区块创建" + month + "月电量区块。");
            }

            if (CellText(worksheet.Cell(2, totalColumn)) != "总实际电量（兆瓦时）"
                || CellText(worksheet.Cell(2, totalColumn + 1)) != "实际电量（兆瓦时）"
                || CellText(worksheet.Cell(3, totalColumn + 1)) != "尖"
                || CellText(worksheet.Cell(3, totalColumn + 2)) != "峰"
                || CellText(worksheet.Cell(3, totalColumn + 3)) != "平"
                || CellText(worksheet.Cell(3, totalColumn + 4)) != "谷")
            {
                throw new InvalidOperationException("重庆台账" + month + "月电量区块表头不符合预期：应包含总实际电量以及尖/峰/平/谷。");
            }

            return new LedgerMap
            {
                CustomerNameColumn = customerNameColumn,
                TotalColumn = totalColumn,
                SharpColumn = totalColumn + 1,
                PeakColumn = totalColumn + 2,
                FlatColumn = totalColumn + 3,
                ValleyColumn = totalColumn + 4,
                DataStartRow = 4
            };
        }

        private static int FindHeaderColumn(IXLWorksheet worksheet, string headerText)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var row = 1; row <= Math.Min(10, worksheet.LastRowUsed()?.RowNumber() ?? 1); row++)
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

            throw new InvalidOperationException("重庆台账中未找到" + month.ToString(CultureInfo.InvariantCulture) + "月电量区块。");
        }

        private static bool TryFindMonthStartColumn(IXLWorksheet worksheet, int month, out int column)
        {
            var label = month.ToString(CultureInfo.InvariantCulture) + "月";
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (column = 1; column <= lastColumn; column++)
            {
                if (CellText(worksheet.Cell(1, column)) == label)
                {
                    return true;
                }
            }

            column = 0;
            return false;
        }

        private static int CopyPreviousMonthBlock(IXLWorksheet worksheet, int targetMonth)
        {
            if (targetMonth <= 1)
            {
                throw new InvalidOperationException("重庆台账中未找到" + targetMonth + "月电量区块，且无法从上月复制区块。");
            }

            var sourceMonth = targetMonth - 1;
            var sourceStart = FindMonthStartColumn(worksheet, sourceMonth);
            var targetStart = sourceStart + MonthBlockWidth;
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            if (targetStart <= lastColumn && HasContent(worksheet, 1, targetStart, lastRow, targetStart + MonthBlockWidth - 1))
            {
                throw new InvalidOperationException("重庆台账" + targetMonth + "月电量区块目标位置已有内容，但未识别为" + targetMonth + "月区块。");
            }

            var sourceRange = worksheet.Range(1, sourceStart, lastRow, sourceStart + MonthBlockWidth - 1);
            sourceRange.CopyTo(worksheet.Cell(1, targetStart));
            worksheet.Cell(1, targetStart).Value = targetMonth.ToString(CultureInfo.InvariantCulture) + "月";
            for (var offset = 0; offset < MonthBlockWidth; offset++)
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

            worksheet.Range(4, targetStart, lastRow, targetStart + MonthPowerColumnCount - 1).Clear(XLClearOptions.Contents);
            return targetStart;
        }

        private static bool HasContent(IXLWorksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
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

        private static List<LedgerCustomerRow> ReadLedgerRows(IXLWorksheet worksheet, LedgerMap map)
        {
            var rows = new List<LedgerCustomerRow>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? map.DataStartRow;
            for (var row = map.DataStartRow; row <= lastRow; row++)
            {
                var customerName = CellText(worksheet.Cell(row, map.CustomerNameColumn));
                if (string.IsNullOrWhiteSpace(customerName))
                {
                    continue;
                }

                rows.Add(new LedgerCustomerRow
                {
                    RowNumber = row,
                    CustomerName = customerName,
                    Key = TextUtil.CustomerKey(customerName)
                });
            }

            return rows;
        }

        private static bool SamePowerVector(
            IXLWorksheet worksheet,
            int row,
            LedgerMap map,
            ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow powerRow)
        {
            double total;
            double sharp;
            double peak;
            double flat;
            double valley;
            return TryCellNumber(worksheet.Cell(row, map.TotalColumn), out total)
                && TryCellNumber(worksheet.Cell(row, map.SharpColumn), out sharp)
                && TryCellNumber(worksheet.Cell(row, map.PeakColumn), out peak)
                && TryCellNumber(worksheet.Cell(row, map.FlatColumn), out flat)
                && TryCellNumber(worksheet.Cell(row, map.ValleyColumn), out valley)
                && SameNumber(total, powerRow.Total)
                && SameNumber(sharp, powerRow.Sharp)
                && SameNumber(peak, powerRow.Peak)
                && SameNumber(flat, powerRow.Flat)
                && SameNumber(valley, powerRow.Valley);
        }

        private static bool IsBlankPower(IXLWorksheet worksheet, int row, LedgerMap map)
        {
            return IsBlank(worksheet.Cell(row, map.TotalColumn))
                && IsBlank(worksheet.Cell(row, map.SharpColumn))
                && IsBlank(worksheet.Cell(row, map.PeakColumn))
                && IsBlank(worksheet.Cell(row, map.FlatColumn))
                && IsBlank(worksheet.Cell(row, map.ValleyColumn));
        }

        private static bool IsBlank(IXLCell cell)
        {
            return string.IsNullOrWhiteSpace(CellText(cell));
        }

        private static bool TryCellNumber(IXLCell cell, out double value)
        {
            value = 0d;
            if (cell.IsEmpty())
            {
                return false;
            }

            if (cell.DataType == XLDataType.Number)
            {
                value = cell.GetDouble();
                return true;
            }

            return double.TryParse(CellText(cell).Replace(",", string.Empty), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        private static bool SameNumber(double left, double right)
        {
            return Math.Abs(left - right) < 0.001;
        }

        private static void AddIssue(ProvinceStage1LedgerUpdatePlan plan, string kind, string category, string severity, string customerName, string message)
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

        private static string CellText(IXLCell cell)
        {
            return TextUtil.S(cell.GetFormattedString());
        }

        private static void WriteReport(string reportPath, ProvinceStage1LedgerUpdateResult result)
        {
            var payload = new
            {
                province = ProvinceDisplayNames.GetName(result.Province),
                month = result.Month,
                unit = result.Unit,
                ledgerPath = result.LedgerPath,
                rawDetailPath = result.RawDetailPath,
                outputLedgerPath = result.OutputLedgerPath,
                ledgerCustomerRows = result.LedgerCustomerRows,
                powerCustomerRows = result.PowerCustomerRows,
                matchedRows = result.MatchedRows,
                updatedPowerRows = result.UpdatedPowerRows,
                manualMatchedRows = result.ManualMatchedRows,
                createdCustomerRows = result.CreatedCustomerRows,
                skippedCustomerRows = result.SkippedCustomerRows,
                customerDecisions = result.CustomerDecisions.Select(decision => new
                {
                    sourceCustomerName = decision.SourceCustomerName,
                    decisionKind = decision.DecisionKind.ToString(),
                    targetCustomerName = decision.TargetCustomerName
                }),
                manualCustomerMatches = result.ManualCustomerMatches.Select(match => new
                {
                    sourceCustomerName = match.SourceCustomerName,
                    targetCustomerName = match.TargetCustomerName
                }),
                multiAccountRows = result.MultiAccountRows,
                skippedRows = result.SkippedRows,
                totalPower = result.TotalPower,
                warnings = result.Warnings,
                issues = result.Issues
            };
            File.WriteAllText(reportPath, JsonConvert.SerializeObject(payload, Formatting.Indented), Encoding.UTF8);
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
            public LedgerMap Map { get; set; }
            public ChongqingPowerCleanGenerator.ChongqingPowerDataSet PowerData { get; set; }
            public List<LedgerCustomerRow> LedgerRows { get; set; }
            public Dictionary<string, ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow> PowerRowsByLedgerKey { get; set; }
            public List<NewCustomerPowerRow> NewCustomerRows { get; set; }
            public Dictionary<string, List<string>> AccountNumbersByKey { get; set; }
            public ProvinceStage1LedgerUpdatePlan Plan { get; set; }
        }

        private sealed class LedgerMap
        {
            public int CustomerNameColumn { get; set; }
            public int TotalColumn { get; set; }
            public int SharpColumn { get; set; }
            public int PeakColumn { get; set; }
            public int FlatColumn { get; set; }
            public int ValleyColumn { get; set; }
            public int DataStartRow { get; set; }
        }

        private sealed class LedgerCustomerRow
        {
            public int RowNumber { get; set; }
            public string CustomerName { get; set; }
            public string Key { get; set; }
        }

        private sealed class NewCustomerPowerRow
        {
            public string CustomerName { get; set; }
            public ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow PowerRow { get; set; }
        }
    }
}
