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
                var codeFillRows = 0;
                foreach (var ledgerRow in context.LedgerRows)
                {
                    ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow powerRow;
                    if (!context.PowerRowsByKey.TryGetValue(ledgerRow.Key, out powerRow))
                    {
                        continue;
                    }

                    var accountText = AccountText(context.AccountNumbersByKey.ContainsKey(ledgerRow.Key)
                        ? context.AccountNumbersByKey[ledgerRow.Key]
                        : new List<string>());
                    if (!string.IsNullOrWhiteSpace(accountText) && CellText(context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.AccountCodeColumn)) != accountText)
                    {
                        context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.AccountCodeColumn).Value = accountText;
                        codeFillRows++;
                    }

                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.TotalColumn).Value = powerRow.Total;
                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.SharpColumn).Value = powerRow.Sharp;
                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.PeakColumn).Value = powerRow.Peak;
                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.FlatColumn).Value = powerRow.Flat;
                    context.Worksheet.Cell(ledgerRow.RowNumber, context.Map.ValleyColumn).Value = powerRow.Valley;
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
                    CodeFillRows = codeFillRows,
                    MultiAccountRows = context.Plan.MultiAccountRows,
                    SkippedRows = context.Plan.MissingInLedgerRows,
                    TotalPower = Math.Round(context.PowerData.CustomerRows.Sum(row => row.Total), 4),
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
            var map = FindLedgerMap(worksheet, options.Month);
            var ledgerRows = ReadLedgerRows(worksheet, map);
            var powerRowsByKey = powerData.CustomerRows.ToDictionary(row => TextUtil.CustomerKey(row.CustomerName));
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

            var plan = BuildPlan(options, powerData, worksheet, map, ledgerRows, powerRowsByKey, accountNumbersByKey);
            return new LedgerUpdateContext
            {
                Worksheet = worksheet,
                Map = map,
                PowerData = powerData,
                LedgerRows = ledgerRows,
                PowerRowsByKey = powerRowsByKey,
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
            IDictionary<string, ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow> powerRowsByKey,
            IDictionary<string, List<string>> accountNumbersByKey)
        {
            var ledgerRowsByKey = ledgerRows.ToDictionary(row => row.Key);
            var matchedKeys = ledgerRowsByKey.Keys.Intersect(powerRowsByKey.Keys).ToList();
            var plan = new ProvinceStage1LedgerUpdatePlan
            {
                Province = ProvinceCode.Chongqing,
                Month = options.Month,
                Unit = Unit,
                LedgerCustomerRows = ledgerRows.Count,
                PowerCustomerRows = powerData.CustomerRows.Count,
                MatchedRows = matchedKeys.Count,
                Warnings = new List<string>(powerData.Warnings)
            };

            if (powerData.Month > 0 && powerData.Month != options.Month)
            {
                AddIssue(plan, "月份不一致", "警告", null, "电量确认单识别为" + powerData.Month + "月，界面选择为" + options.Month + "月。");
            }

            foreach (var key in matchedKeys)
            {
                var ledgerRow = ledgerRowsByKey[key];
                var powerRow = powerRowsByKey[key];
                var accounts = accountNumbersByKey.ContainsKey(key) ? accountNumbersByKey[key] : new List<string>();
                var accountText = AccountText(accounts);
                if (accounts.Count > 1)
                {
                    plan.MultiAccountRows++;
                    AddIssue(plan, "多户号客户", "提示", ledgerRow.CustomerName, "该客户在电量明细中存在多个户号；继续后会用顿号合并写入电力用户编码列。");
                }

                if (!string.IsNullOrWhiteSpace(accountText))
                {
                    var currentAccountText = CellText(worksheet.Cell(ledgerRow.RowNumber, map.AccountCodeColumn));
                    if (string.IsNullOrWhiteSpace(currentAccountText))
                    {
                        plan.CodeFillRows++;
                    }
                    else if (!SameAccountSet(currentAccountText, accounts))
                    {
                        AddIssue(plan, "电力用户编码差异", "警告", ledgerRow.CustomerName, "台账现有电力用户编码与电量明细户号不一致；继续后会按电量明细户号写入副本。");
                    }
                }

                if (!IsBlankPower(worksheet, ledgerRow.RowNumber, map) && !SamePowerVector(worksheet, ledgerRow.RowNumber, map, powerRow))
                {
                    plan.ExistingDifferentPowerRows++;
                    AddIssue(plan, "已有电量差异", "警告", ledgerRow.CustomerName, "台账目标月份已有电量，且与清洗结果不一致；继续后会按清洗结果写入副本。");
                }
            }

            var missingInLedger = powerRowsByKey.Keys.Except(ledgerRowsByKey.Keys).ToList();
            foreach (var key in missingInLedger)
            {
                plan.MissingInLedgerRows++;
                AddIssue(plan, "电量客户不在台账", "警告", powerRowsByKey[key].CustomerName, "清洗结果中的客户在台账中找不到，继续后不会新增该客户行。");
            }

            var missingInPower = ledgerRowsByKey.Keys.Except(powerRowsByKey.Keys).ToList();
            foreach (var key in missingInPower)
            {
                plan.MissingInPowerRows++;
                AddIssue(plan, "台账客户不在电量表", "提示", ledgerRowsByKey[key].CustomerName, "台账客户在清洗结果中找不到，继续后该行目标月份电量不会更新。");
            }

            foreach (var powerKey in missingInLedger)
            {
                var powerRow = powerRowsByKey[powerKey];
                foreach (var ledgerKey in missingInPower)
                {
                    if (SamePowerVector(worksheet, ledgerRowsByKey[ledgerKey].RowNumber, map, powerRow))
                    {
                        plan.AliasCandidateRows++;
                        AddIssue(plan, "疑似名称差异", "警告", powerRow.CustomerName, "该电量客户与某个台账客户目标月份电量完全一致，可能是名称别名；请人工确认。");
                        break;
                    }
                }
            }

            return plan;
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

        private static LedgerMap FindLedgerMap(IXLWorksheet worksheet, int month)
        {
            var accountCodeColumn = FindHeaderColumn(worksheet, "电力用户编码");
            var customerNameColumn = FindHeaderColumn(worksheet, "电力用户名称");
            if (accountCodeColumn <= 0)
            {
                throw new InvalidOperationException("重庆台账中未找到表头“电力用户编码”。");
            }

            var totalColumn = FindMonthStartColumn(worksheet, month);
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
                AccountCodeColumn = accountCodeColumn,
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
            var label = month.ToString(CultureInfo.InvariantCulture) + "月";
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (CellText(worksheet.Cell(1, column)) == label)
                {
                    return column;
                }
            }

            throw new InvalidOperationException("重庆台账中未找到" + label + "电量区块。");
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

        private static string AccountText(IList<string> accounts)
        {
            return accounts == null || accounts.Count == 0 ? string.Empty : string.Join("、", accounts);
        }

        private static bool SameAccountSet(string currentText, IList<string> accounts)
        {
            var current = TextUtil.S(currentText)
                .Split(new[] { '、', ',', '，', ';', '；', '/', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(TextUtil.S)
                .Where(value => value.Length > 0)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToList();
            return current.SequenceEqual(accounts.OrderBy(value => value, StringComparer.Ordinal));
        }

        private static void AddIssue(ProvinceStage1LedgerUpdatePlan plan, string category, string severity, string customerName, string message)
        {
            plan.Issues.Add(new ProvinceStage1LedgerUpdateIssue
            {
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
                province = ProvinceStage1Service.ProvinceName(result.Province),
                month = result.Month,
                unit = result.Unit,
                ledgerPath = result.LedgerPath,
                rawDetailPath = result.RawDetailPath,
                outputLedgerPath = result.OutputLedgerPath,
                ledgerCustomerRows = result.LedgerCustomerRows,
                powerCustomerRows = result.PowerCustomerRows,
                matchedRows = result.MatchedRows,
                updatedPowerRows = result.UpdatedPowerRows,
                codeFillRows = result.CodeFillRows,
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
            public Dictionary<string, ChongqingPowerCleanGenerator.ChongqingPowerAggregateRow> PowerRowsByKey { get; set; }
            public Dictionary<string, List<string>> AccountNumbersByKey { get; set; }
            public ProvinceStage1LedgerUpdatePlan Plan { get; set; }
        }

        private sealed class LedgerMap
        {
            public int AccountCodeColumn { get; set; }
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
    }
}
