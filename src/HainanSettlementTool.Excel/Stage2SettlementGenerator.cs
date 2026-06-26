using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal sealed class Stage2SettlementGenerator
    {
        private const int DataStartRow = 5;
        private const int Year = 2026;
        private static readonly Dictionary<string, Tuple<int, string>> PaymentPartyOverrides =
            new Dictionary<string, Tuple<int, string>>
            {
                { PaymentKey("海南精研科技有限公司", "代理费"), Tuple.Create(3, "清能") }
            };

        public Stage2Report Generate(Stage2Options options)
        {
            Directory.CreateDirectory(options.OutputDirectory);
            var proxyRows = new List<DetailSettlementRow>();
            var interRows = new List<DetailSettlementRow>();
            ReadLedgerRows(options.LedgerPath, options.Month, proxyRows, interRows);

            var missingOwners = proxyRows
                .Concat(interRows)
                .Where(row => string.IsNullOrWhiteSpace(row.Owner))
                .Select(row => "第" + row.LedgerRow + "行 " + row.Customer + " 缺少负责人")
                .Distinct()
                .ToList();
            if (missingOwners.Count > 0 && !options.AllowMissingOwner)
            {
                throw new InvalidOperationException(options.Month + "月结算明细存在负责人缺失，先不要生成分表/汇总表：" + string.Join("、", missingOwners.Take(10)));
            }

            var auditIssues = BuildPreflightIssues(options, proxyRows, interRows);
            var totals = BuildSplitFiles(options, proxyRows, interRows, auditIssues);
            var warnings = new List<string>();
            var summaryPath = BuildSummary(options, totals, warnings);
            var report = CreateReport(options, proxyRows, interRows, totals, summaryPath, warnings, missingOwners, auditIssues);
            WriteReport(options, report);
            WriteWarnings(options, warnings);
            WriteAuditReport(options, report);
            return report;
        }

        public Stage2PreflightReport Analyze(Stage2Options options)
        {
            var proxyRows = new List<DetailSettlementRow>();
            var interRows = new List<DetailSettlementRow>();
            ReadLedgerRows(options.LedgerPath, options.Month, proxyRows, interRows);

            var report = new Stage2PreflightReport
            {
                Month = options.Month
            };
            report.Issues.AddRange(BuildPreflightIssues(options, proxyRows, interRows));
            return report;
        }

        private static void ReadLedgerRows(string ledgerPath, int month, List<DetailSettlementRow> proxyRows, List<DetailSettlementRow> interRows)
        {
            using (var workbook = new XLWorkbook(ledgerPath))
            {
                var worksheet = ClosedXmlUtil.MainSheet(workbook);
                var start = FindMonthStartColumn(worksheet, month);
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                for (var row = 4; row <= lastRow; row++)
                {
                    var customer = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                    if (string.IsNullOrWhiteSpace(customer))
                    {
                        continue;
                    }

                    var total = GetNumeric(worksheet, row, start);
                    if (total <= 0)
                    {
                        continue;
                    }

                    var owner = TextUtil.S(worksheet.Cell(row, 10).GetFormattedString());
                    var developer = TextUtil.S(worksheet.Cell(row, 8).GetFormattedString());
                    var interName = TextUtil.S(worksheet.Cell(row, 19).GetFormattedString());

                    var interRow = CreateDetailRow(worksheet, row, start, customer, owner, interName, "居间", start + 7, start + 8, start + 10, start + 12);
                    if (!string.IsNullOrWhiteSpace(interName) && HasSettlementAmount(interRow))
                    {
                        interRows.Add(interRow);
                    }

                    var proxyRow = CreateDetailRow(worksheet, row, start, customer, owner, developer, "代理", start + 13, start + 14, start + 16, start + 18);
                    if (!string.IsNullOrWhiteSpace(developer) && HasSettlementAmount(proxyRow))
                    {
                        proxyRows.Add(proxyRow);
                    }
                }
            }
        }

        private static DetailSettlementRow CreateDetailRow(
            IXLWorksheet worksheet,
            int ledgerRow,
            int start,
            string customer,
            string owner,
            string entity,
            string kind,
            int ratioColumn,
            int unitPriceColumn,
            int taxRateColumn,
            int cachedNetColumn)
        {
            var total = GetNumeric(worksheet, ledgerRow, start);
            var ratio = GetNumeric(worksheet, ledgerRow, ratioColumn);
            var unitPrice = GetNumeric(worksheet, ledgerRow, unitPriceColumn);
            var taxRate = GetNumeric(worksheet, ledgerRow, taxRateColumn);
            var ledgerNet = GetNumeric(worksheet, ledgerRow, cachedNetColumn);
            var amounts = Stage2SettlementCalculator.CalculateAmounts(total, ratio, unitPrice, taxRate);

            return new DetailSettlementRow
            {
                LedgerRow = ledgerRow,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                Kind = kind,
                Total = total,
                Sharp = GetNumeric(worksheet, ledgerRow, start + 1),
                Peak = GetNumeric(worksheet, ledgerRow, start + 2),
                Flat = GetNumeric(worksheet, ledgerRow, start + 3),
                Valley = GetNumeric(worksheet, ledgerRow, start + 4),
                PeakFlat = GetNumeric(worksheet, ledgerRow, start + 5),
                ValleyFlat = GetNumeric(worksheet, ledgerRow, start + 6),
                Ratio = ratio,
                UnitPrice = unitPrice,
                TaxRate = taxRate,
                LedgerNet = ledgerNet,
                Gross = amounts.Gross,
                Adjustment = amounts.Adjustment,
                AdjustedGross = amounts.AdjustedGross,
                TaxAmount = amounts.TaxAmount,
                CalculatedNet = amounts.CalculatedNet,
                ExpectedNet = amounts.ExpectedNet
            };
        }

        private static bool HasSettlementAmount(DetailSettlementRow row)
        {
            return Math.Abs(row.LedgerNet) > Stage2SettlementCalculator.AmountTolerance
                || Math.Abs(row.CalculatedNet) > Stage2SettlementCalculator.AmountTolerance;
        }

        private static double GetNumeric(IXLWorksheet worksheet, int row, int column)
        {
            var cell = worksheet.Cell(row, column);
            var value = ClosedXmlUtil.CellNumber(cell);
            if (value != 0)
            {
                return value;
            }

            var formula = TextUtil.S(cell.FormulaA1).Replace("$", string.Empty);
            var match = Regex.Match(formula, "^([A-Z]{1,3})(\\d+)$");
            if (!match.Success)
            {
                return value;
            }

            var targetColumn = ColumnNumber(match.Groups[1].Value);
            var targetRow = Convert.ToInt32(match.Groups[2].Value);
            if (targetRow == row && targetColumn == column)
            {
                return 0;
            }

            return ClosedXmlUtil.CellNumber(worksheet.Cell(targetRow, targetColumn));
        }

        private static int FindMonthStartColumn(IXLWorksheet worksheet, int month)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (TextUtil.S(worksheet.Cell(1, column).GetFormattedString()) == month + "月")
                {
                    return column;
                }
            }

            throw new InvalidOperationException("未找到 " + month + "月 的台账区块。");
        }

        private static List<Stage2CheckIssue> BuildPreflightIssues(Stage2Options options, IList<DetailSettlementRow> proxyRows, IList<DetailSettlementRow> interRows)
        {
            var issues = new List<Stage2CheckIssue>();
            var templateMap = BuildTemplateIndex(options.ProxyTemplateDirectory, options.IntermediaryTemplateDirectory);
            var grouped = proxyRows
                .Select(row => new { Key = Tuple.Create("代理", row.Owner, row.Entity), Row = row })
                .Concat(interRows.Select(row => new { Key = Tuple.Create("居间", row.Owner, row.Entity), Row = row }))
                .GroupBy(item => item.Key)
                .OrderBy(group => group.Key.Item1)
                .ThenBy(group => group.Key.Item2)
                .ThenBy(group => group.Key.Item3);

            foreach (var group in grouped)
            {
                var kind = group.Key.Item1;
                var owner = group.Key.Item2;
                var entity = group.Key.Item3;
                var templateKey = TemplateKey(kind, owner, entity);
                string templatePath;
                if (!templateMap.TryGetValue(templateKey, out templatePath))
                {
                    issues.Add(new Stage2CheckIssue
                    {
                        Severity = "提示",
                        Category = "未匹配到上月分表模板",
                        Kind = kind + "费",
                        Owner = owner,
                        Entity = entity,
                        Message = kind + "费主体“" + entity + "”未在上月分表文件夹中匹配到同名模板。",
                        Suggestion = "程序会复制同类型模板生成新分表；请确认这是新增关系，或检查负责人文件夹、分表文件名、A2代理名称是否和台账一致。"
                    });
                    continue;
                }

                CompareGroupWithTemplate(options.Month, kind, owner, entity, templatePath, group.Select(item => item.Row).ToList(), issues);
            }

            return issues;
        }

        private static void CompareGroupWithTemplate(
            int month,
            string kind,
            string owner,
            string entity,
            string templatePath,
            IList<DetailSettlementRow> currentRows,
            IList<Stage2CheckIssue> issues)
        {
            try
            {
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var previousSheet = PreviousMonthSheet(workbook, month, month + "月") ?? LastMonthSheet(workbook);
                    var previousRows = ReadPreviousDetails(previousSheet);
                    foreach (var row in currentRows)
                    {
                        PreviousDetailRow previous;
                        if (!previousRows.TryGetValue(NormalizeName(row.Customer), out previous))
                        {
                            issues.Add(new Stage2CheckIssue
                            {
                                Severity = "提示",
                                Category = "客户本月新增到分表",
                                Kind = kind + "费",
                                Customer = row.Customer,
                                Owner = owner,
                                Entity = entity,
                                LedgerRow = row.LedgerRow,
                                TemplateFile = templatePath,
                                SheetName = previousSheet.Name,
                                Message = kind + "费主体“" + entity + "”下的客户“" + row.Customer + "”在上月分表未找到。",
                                Suggestion = "如果这是本月新增客户，可以继续；否则请检查台账客户名称和上月分表客户名称是否一致。"
                            });
                            continue;
                        }

                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "电量比例", previous.Ratio, row.Ratio);
                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "利润单价", previous.UnitPrice, row.UnitPrice);
                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "税率", previous.TaxRate, row.TaxRate);
                    }
                }
            }
            catch (Exception ex)
            {
                issues.Add(new Stage2CheckIssue
                {
                    Severity = "提示",
                    Category = "上月分表预检失败",
                    Kind = kind + "费",
                    Owner = owner,
                    Entity = entity,
                    TemplateFile = templatePath,
                    Message = "读取上月分表模板失败：" + ex.Message,
                    Suggestion = "程序仍可继续生成；请确认该分表文件没有损坏，且包含上月明细和合计行。"
                });
            }
        }

        private static Dictionary<string, PreviousDetailRow> ReadPreviousDetails(IXLWorksheet worksheet)
        {
            var result = new Dictionary<string, PreviousDetailRow>();
            var totalRow = FindTotalRow(worksheet, DataStartRow);
            for (var row = DataStartRow; row < totalRow; row++)
            {
                var customer = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                if (string.IsNullOrWhiteSpace(customer))
                {
                    continue;
                }

                var key = NormalizeName(customer);
                if (result.ContainsKey(key))
                {
                    continue;
                }

                result[key] = new PreviousDetailRow
                {
                    Customer = customer,
                    Row = row,
                    SheetName = worksheet.Name,
                    Ratio = GetNumeric(worksheet, row, 10),
                    UnitPrice = GetNumeric(worksheet, row, 11),
                    TaxRate = GetNumeric(worksheet, row, 17)
                };
            }

            return result;
        }

        private static void AddValueChangeIssue(
            IList<Stage2CheckIssue> issues,
            string kind,
            string owner,
            string entity,
            DetailSettlementRow current,
            PreviousDetailRow previous,
            string templatePath,
            string fieldName,
            double previousValue,
            double currentValue)
        {
            if (Math.Abs(previousValue - currentValue) <= Stage2SettlementCalculator.AmountTolerance)
            {
                return;
            }

            issues.Add(new Stage2CheckIssue
            {
                Severity = "确认",
                Category = "关键字段较上月变化",
                Kind = kind + "费",
                Customer = current.Customer,
                Owner = owner,
                Entity = entity,
                LedgerRow = current.LedgerRow,
                TemplateFile = templatePath,
                SheetName = previous.SheetName,
                PreviousValue = fieldName + " 上月：" + Stage2SettlementCalculator.FormatAmount(previousValue),
                CurrentValue = fieldName + " 本月：" + Stage2SettlementCalculator.FormatAmount(currentValue),
                Message = kind + "费主体“" + entity + "”下的客户“" + current.Customer + "”" + fieldName + "发生变化（上月分表第" + previous.Row + "行，本月台账第" + current.LedgerRow + "行）。",
                Suggestion = "如果这是本月台账更新后的正常变化，请继续；否则请回到台账检查该客户的" + fieldName + "。"
            });
        }

        private static List<GroupSettlementTotal> BuildSplitFiles(
            Stage2Options options,
            IList<DetailSettlementRow> proxyRows,
            IList<DetailSettlementRow> interRows,
            IList<Stage2CheckIssue> auditIssues)
        {
            var templateMap = BuildTemplateIndex(options.ProxyTemplateDirectory, options.IntermediaryTemplateDirectory);
            var grouped = proxyRows
                .Select(row => new { Key = Tuple.Create("代理", row.Owner, row.Entity), Row = row })
                .Concat(interRows.Select(row => new { Key = Tuple.Create("居间", row.Owner, row.Entity), Row = row }))
                .GroupBy(item => item.Key)
                .OrderBy(group => group.Key.Item1)
                .ThenBy(group => group.Key.Item2)
                .ThenBy(group => group.Key.Item3);

            var totals = new List<GroupSettlementTotal>();
            foreach (var group in grouped)
            {
                var kind = group.Key.Item1;
                var owner = group.Key.Item2;
                var entity = group.Key.Item3;
                bool matchedTemplate;
                var outputPath = EnsureOutputWorkbook(templateMap, options, kind, owner, entity, out matchedTemplate);
                FileAccessGuard.RequireWritableWorkbook(outputPath, kind + "分表输出文件");

                using (var workbook = new XLWorkbook(outputPath))
                {
                    var displayEntity = matchedTemplate ? PriorSheetDisplayEntity(workbook, options.Month) : entity;
                    var worksheet = PrepareMonthSheet(workbook, options.Month);
                    WriteDetailSheet(worksheet, kind, entity, options.Month, group.Select(item => item.Row).ToList(), displayEntity, outputPath, auditIssues);
                    SaveWorkbook(workbook, outputPath);

                    totals.Add(new GroupSettlementTotal
                    {
                        Kind = kind + "费",
                        Owner = owner,
                        Entity = entity,
                        DisplayEntity = string.IsNullOrWhiteSpace(displayEntity) ? entity : displayEntity,
                        Rows = group.Count(),
                        ExpectedNet = Math.Round(group.Sum(item => item.Row.ExpectedNet), 4),
                        OutputFile = outputPath
                    });
                }
            }

            return totals;
        }

        private static Dictionary<string, string> BuildTemplateIndex(string proxyRoot, string interRoot)
        {
            var result = new Dictionary<string, string>();
            IndexTemplateRoot(result, "代理", proxyRoot);
            IndexTemplateRoot(result, "居间", interRoot);
            return result;
        }

        private static void IndexTemplateRoot(IDictionary<string, string> result, string kind, string root)
        {
            if (!Directory.Exists(root))
            {
                return;
            }

            foreach (var path in Directory.GetFiles(root, "*.xlsx", SearchOption.AllDirectories))
            {
                if (Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                {
                    continue;
                }

                try
                {
                    using (var workbook = new XLWorkbook(path))
                    {
                        var sheet = LastMonthSheet(workbook);
                        var rawEntity = TextUtil.S(sheet.Cell("A2").GetFormattedString()).Replace("代理名称:", string.Empty);
                        var owner = new DirectoryInfo(Path.GetDirectoryName(path)).Name.Replace(" - 海南2026", string.Empty);
                        result[TemplateKey(kind, owner, rawEntity)] = path;
                    }
                }
                catch
                {
                    // Ignore broken or unrelated workbooks in template folders.
                }
            }
        }

        private static IXLWorksheet LastMonthSheet(XLWorkbook workbook)
        {
            var monthSheets = workbook.Worksheets
                .Where(sheet => Regex.IsMatch(sheet.Name, "^\\d+月$"))
                .OrderBy(sheet => Convert.ToInt32(sheet.Name.Replace("月", string.Empty)))
                .ToList();
            return monthSheets.Count > 0 ? monthSheets.Last() : workbook.Worksheets.Last();
        }

        private static string EnsureOutputWorkbook(
            IDictionary<string, string> templateMap,
            Stage2Options options,
            string kind,
            string owner,
            string entity,
            out bool matchedTemplate)
        {
            string source;
            if (templateMap.TryGetValue(TemplateKey(kind, owner, entity), out source))
            {
                var sourceRoot = kind == "代理" ? options.ProxyTemplateDirectory : options.IntermediaryTemplateDirectory;
                var targetRoot = Path.Combine(options.OutputDirectory, kind == "代理" ? "2026年代理 - 海南" : "2026年居间 - 海南");
                var relative = RelativePath(sourceRoot, source);
                var target = Path.Combine(targetRoot, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                if (!File.Exists(target))
                {
                    File.Copy(source, target, false);
                }

                matchedTemplate = true;
                return target;
            }

            var baseRoot = Path.Combine(options.OutputDirectory, kind == "代理" ? "2026年代理 - 海南" : "2026年居间 - 海南");
            var folder = Path.Combine(baseRoot, TextUtil.SafeFileName(owner) + " - 海南2026");
            Directory.CreateDirectory(folder);
            var newTarget = Path.Combine(folder, TextUtil.SafeFileName(entity) + " 2026海南.xlsx");
            if (File.Exists(newTarget))
            {
                matchedTemplate = false;
                return newTarget;
            }

            var candidate = templateMap
                .Where(item => item.Key.StartsWith(kind + "|", StringComparison.Ordinal))
                .Select(item => item.Value)
                .FirstOrDefault();
            if (string.IsNullOrWhiteSpace(candidate))
            {
                throw new InvalidOperationException("没有可用的" + kind + "分表模板。");
            }

            File.Copy(candidate, newTarget, false);
            matchedTemplate = false;
            return newTarget;
        }

        private static IXLWorksheet PrepareMonthSheet(XLWorkbook workbook, int month)
        {
            var title = month + "月";
            var existing = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == title);
            var source = PreviousMonthSheet(workbook, month, title);
            if (existing != null)
            {
                if (source == null)
                {
                    var tempTitle = title + "__template";
                    while (workbook.Worksheets.Any(sheet => sheet.Name == tempTitle))
                    {
                        tempTitle += "_";
                    }

                    var clone = existing.CopyTo(tempTitle);
                    existing.Delete();
                    clone.Name = title;
                    return clone;
                }

                existing.Delete();
            }

            if (source == null)
            {
                source = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "3月") ?? workbook.Worksheets.Last();
            }

            return source.CopyTo(title);
        }

        private static IXLWorksheet PreviousMonthSheet(XLWorkbook workbook, int month, string targetTitle)
        {
            var candidates = workbook.Worksheets
                .Select(sheet =>
                {
                    int sheetMonth;
                    return new { Sheet = sheet, Matched = TryParseMonthSheet(sheet.Name, out sheetMonth), Month = sheetMonth };
                })
                .Where(item => item.Matched && item.Month < month && item.Sheet.Name != targetTitle)
                .OrderBy(item => item.Month)
                .ToList();
            return candidates.Count == 0 ? null : candidates.Last().Sheet;
        }

        private static bool TryParseMonthSheet(string name, out int month)
        {
            month = 0;
            var match = Regex.Match(TextUtil.S(name), "^(\\d{1,2})月$");
            return match.Success && int.TryParse(match.Groups[1].Value, out month);
        }

        private static string PriorSheetDisplayEntity(XLWorkbook workbook, int month)
        {
            foreach (var candidate in new[] { (month - 1) + "月", "3月", "2月", "1月" })
            {
                var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == candidate);
                if (sheet == null)
                {
                    continue;
                }

                var text = TextUtil.S(sheet.Cell("A2").GetFormattedString());
                if (text.StartsWith("代理名称:", StringComparison.Ordinal))
                {
                    return text.Replace("代理名称:", string.Empty);
                }
            }

            return null;
        }

        private static void WriteDetailSheet(
            IXLWorksheet worksheet,
            string kind,
            string entity,
            int month,
            IList<DetailSettlementRow> rows,
            string displayEntity,
            string outputPath,
            IList<Stage2CheckIssue> auditIssues)
        {
            SetTopTitles(worksheet, kind, entity, month, displayEntity);
            var totalRow = AdjustDetailRows(worksheet, rows.Count);
            for (var index = 0; index < rows.Count; index++)
            {
                var row = rows[index];
                var excelRow = DataStartRow + index;
                worksheet.Cell(excelRow, 1).Value = index + 1;
                worksheet.Cell(excelRow, 2).Value = row.Customer;
                worksheet.Cell(excelRow, 3).Value = Math.Round(row.Total, 4);
                worksheet.Cell(excelRow, 4).Value = Math.Round(row.Sharp, 4);
                worksheet.Cell(excelRow, 5).Value = Math.Round(row.Peak, 4);
                worksheet.Cell(excelRow, 6).Value = Math.Round(row.Flat, 4);
                worksheet.Cell(excelRow, 7).Value = Math.Round(row.Valley, 4);
                worksheet.Cell(excelRow, 8).Value = Math.Round(row.PeakFlat, 4);
                worksheet.Cell(excelRow, 9).Value = Math.Round(row.ValleyFlat, 4);
                worksheet.Cell(excelRow, 10).Value = Math.Round(row.Ratio, 4);
                worksheet.Cell(excelRow, 11).Value = Math.Round(row.UnitPrice, 4);
                worksheet.Cell(excelRow, 14).Clear(XLClearOptions.Contents);
                SetDetailFormula(worksheet.Cell(excelRow, 12), "ROUND(C" + excelRow + "*J" + excelRow + "*K" + excelRow + ",4)");
                SetDetailFormula(worksheet.Cell(excelRow, 13), "L" + excelRow + "-N" + excelRow);
                SetDetailFormula(worksheet.Cell(excelRow, 15), "ROUND(M" + excelRow + "/1.13*Q" + excelRow + ",4)");
                SetDetailFormula(worksheet.Cell(excelRow, 16), "M" + excelRow + "-O" + excelRow);
                worksheet.Cell(excelRow, 17).Value = row.TaxRate;
                for (var column = 18; column <= (worksheet.LastColumnUsed()?.ColumnNumber() ?? 18); column++)
                {
                    worksheet.Cell(excelRow, column).Clear(XLClearOptions.Contents);
                }

                AddLedgerDifferenceIssue(row, kind, outputPath, worksheet.Name, auditIssues);
            }

            if (rows.Count > 0)
            {
                var last = DataStartRow + rows.Count - 1;
                worksheet.Cell(totalRow, 1).Value = "合计";
                for (var column = 3; column <= 7; column++)
                {
                    var letter = ClosedXmlUtil.ColumnLetter(column);
                    SetDetailFormula(worksheet.Cell(totalRow, column), "SUM(" + letter + DataStartRow + ":" + letter + last + ")");
                }

                for (var column = 12; column <= 16; column++)
                {
                    var letter = ClosedXmlUtil.ColumnLetter(column);
                    SetDetailFormula(worksheet.Cell(totalRow, column), "SUM(" + letter + DataStartRow + ":" + letter + last + ")");
                }
            }

            RepairDetailTotalRowStyles(worksheet, month, totalRow);
            UpdateDetailSignatureDate(worksheet, totalRow);
        }

        private static void SetDetailFormula(IXLCell cell, string formula)
        {
            cell.FormulaA1 = formula;
        }

        private static void RepairDetailTotalRowStyles(IXLWorksheet worksheet, int month, int totalRow)
        {
            for (var column = 1; column <= 17; column++)
            {
                var target = worksheet.Cell(totalRow, column);
                if (HasTemplateStyle(target))
                {
                    continue;
                }

                var source = FindPriorDetailTotalStyleCell(worksheet, month, column)
                    ?? FindDetailRowStyleFallback(worksheet, totalRow, column);
                if (source != null)
                {
                    target.Style = source.Style;
                }
            }
        }

        private static IXLCell FindPriorDetailTotalStyleCell(IXLWorksheet worksheet, int month, int column)
        {
            for (var candidateMonth = month - 1; candidateMonth >= 1; candidateMonth--)
            {
                var candidate = worksheet.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == candidateMonth + "月");
                if (candidate == null)
                {
                    continue;
                }

                int totalRow;
                try
                {
                    totalRow = FindTotalRow(candidate, DataStartRow);
                }
                catch
                {
                    continue;
                }

                var source = candidate.Cell(totalRow, column);
                if (HasTemplateStyle(source))
                {
                    return source;
                }
            }

            return null;
        }

        private static IXLCell FindDetailRowStyleFallback(IXLWorksheet worksheet, int totalRow, int column)
        {
            var row = Math.Max(DataStartRow, totalRow - 1);
            var source = worksheet.Cell(row, column);
            return HasTemplateStyle(source) ? source : null;
        }

        private static bool HasTemplateStyle(IXLCell cell)
        {
            return cell.Style.Alignment.Horizontal != XLAlignmentHorizontalValues.General
                || cell.Style.Border.LeftBorder != XLBorderStyleValues.None
                || cell.Style.Border.RightBorder != XLBorderStyleValues.None
                || cell.Style.Border.TopBorder != XLBorderStyleValues.None
                || cell.Style.Border.BottomBorder != XLBorderStyleValues.None;
        }

        private static void UpdateDetailSignatureDate(IXLWorksheet worksheet, int totalRow)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.Address.RowNumber <= totalRow)
                {
                    continue;
                }

                var text = TextUtil.S(cell.GetFormattedString());
                if (!text.Contains("日期：") || text.Contains("结算日期"))
                {
                    continue;
                }

                var updated = ShiftSignatureDateText(text, 1);
                if (!string.IsNullOrWhiteSpace(updated))
                {
                    cell.Value = updated;
                }
            }
        }

        private static string ShiftSignatureDateText(string text, int months)
        {
            var match = Regex.Match(TextUtil.S(text), "日期：\\s*(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日");
            if (!match.Success)
            {
                return null;
            }

            var date = new DateTime(
                Convert.ToInt32(match.Groups[1].Value),
                Convert.ToInt32(match.Groups[2].Value),
                Convert.ToInt32(match.Groups[3].Value)).AddMonths(months);
            return "日期：" + date.Year + "年" + date.Month.ToString("00") + "月" + date.Day.ToString("00") + "日";
        }

        private static void AddLedgerDifferenceIssue(
            DetailSettlementRow row,
            string kind,
            string outputPath,
            string sheetName,
            IList<Stage2CheckIssue> auditIssues)
        {
            var issue = Stage2SettlementCalculator.CreateLedgerDifferenceIssue(row, kind, outputPath, sheetName);
            if (issue == null)
            {
                return;
            }

            auditIssues.Add(issue);
        }

        private static void SetTopTitles(IXLWorksheet worksheet, string kind, string entity, int month, string displayEntity)
        {
            worksheet.Cell("A1").Value = kind + "费用结算单";
            worksheet.Cell("A2").Value = "代理名称:" + (string.IsNullOrWhiteSpace(displayEntity) ? entity : displayEntity);
            SetFirstCellContaining(worksheet, "所属期", "所属期：" + Year + " 年 " + month.ToString("00") + " 月");
            var nextMonth = month == 12 ? 1 : month + 1;
            var year = month == 12 ? Year + 1 : Year;
            SetFirstCellContaining(worksheet, "结算日期", "结算日期：" + year + " 年 " + nextMonth.ToString("00") + " 月 15 日");
        }

        private static void SetFirstCellContaining(IXLWorksheet worksheet, string needle, string value)
        {
            var maxRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 1, 8);
            var maxColumn = Math.Min(worksheet.LastColumnUsed()?.ColumnNumber() ?? 1, 24);
            for (var row = 1; row <= maxRow; row++)
            {
                for (var column = 1; column <= maxColumn; column++)
                {
                    if (TextUtil.S(worksheet.Cell(row, column).GetFormattedString()).Contains(needle))
                    {
                        worksheet.Cell(row, column).Value = value;
                        return;
                    }
                }
            }
        }

        private static int AdjustDetailRows(IXLWorksheet worksheet, int count)
        {
            var totalRow = FindTotalRow(worksheet, DataStartRow);
            var existing = totalRow - DataStartRow;
            if (count > existing)
            {
                worksheet.Row(totalRow).InsertRowsAbove(count - existing);
                for (var row = totalRow; row < totalRow + count - existing; row++)
                {
                    worksheet.Row(DataStartRow).CopyTo(worksheet.Row(row));
                }
            }
            else if (count < existing)
            {
                worksheet.Rows(DataStartRow + count, totalRow - 1).Delete();
            }

            return FindTotalRow(worksheet, DataStartRow);
        }

        private static string BuildSummary(Stage2Options options, IList<GroupSettlementTotal> totals, IList<string> warnings)
        {
            var outputName = string.IsNullOrWhiteSpace(options.OutputSummaryName)
                ? "【2026年海南省代理费汇总表-" + options.Month + "月自动化】.xlsx"
                : options.OutputSummaryName;
            var outputPath = Path.Combine(options.OutputDirectory, outputName);
            FileAccessGuard.RequireWritableWorkbook(outputPath, "输出汇总表");
            File.Copy(options.SummaryTemplatePath, outputPath, true);

            using (var workbook = new XLWorkbook(outputPath))
            {
                var mainSheetName = ResolveSummarySheetName(workbook, "main", true);
                var qingnengSheetName = ResolveSummarySheetName(workbook, "qingneng", false);
                var qinghuiSheetName = ResolveSummarySheetName(workbook, "qinghui", false);
                var mainMeta = ReadSummaryMeta(workbook.Worksheet(mainSheetName));
                var partyByKey = mainMeta.ToDictionary(
                    item => SummaryKey(item.Entity, item.Kind),
                    item => PaymentPartyFor(item.Entity, item.Kind, options.Month, string.IsNullOrWhiteSpace(item.PaymentParty) ? "清辉" : item.PaymentParty));

                WriteSummarySheet(workbook.Worksheet(mainSheetName), totals, options.Month, null, warnings);

                if (!string.IsNullOrWhiteSpace(qingnengSheetName))
                {
                    var qnTotals = totals.Where(total => PartyForSummaryTotal(total, partyByKey) == "清能").ToList();
                    var allowed = new HashSet<string>(partyByKey.Where(item => item.Value == "清能").Select(item => item.Key));
                    foreach (var total in qnTotals)
                    {
                        allowed.Add(SummaryKey(total.Entity, total.Kind));
                    }
                    WriteSummarySheet(workbook.Worksheet(qingnengSheetName), qnTotals, options.Month, allowed, warnings);
                }

                if (!string.IsNullOrWhiteSpace(qinghuiSheetName))
                {
                    var qhTotals = totals.Where(total => PartyForSummaryTotal(total, partyByKey) == "清辉").ToList();
                    var allowed = new HashSet<string>(partyByKey.Where(item => item.Value == "清辉").Select(item => item.Key));
                    foreach (var total in qhTotals)
                    {
                        allowed.Add(SummaryKey(total.Entity, total.Kind));
                    }
                    WriteSummarySheet(workbook.Worksheet(qinghuiSheetName), qhTotals, options.Month, allowed, warnings);
                }

                SaveWorkbook(workbook, outputPath);
            }

            return outputPath;
        }

        private static void WriteSummarySheet(
            IXLWorksheet worksheet,
            IList<GroupSettlementTotal> totals,
            int month,
            ISet<string> allowedKeys,
            IList<string> warnings)
        {
            var startRow = 4;
            DeleteSummaryRowsNotAllowed(worksheet, startRow, allowedKeys);

            var monthColumn = InsertMonthBlock(worksheet, month);
            var cumulativeColumn = monthColumn + 6;
            var totalByKey = totals.ToDictionary(total => SummaryKey(total.Entity, total.Kind), total => total);
            var existingMeta = ReadSummaryMeta(worksheet);
            var knownKeys = new HashSet<string>(existingMeta.Select(item => SummaryKey(item.Entity, item.Kind)));
            var newTotals = totals.Where(total => !knownKeys.Contains(SummaryKey(total.Entity, total.Kind))).ToList();
            var newKeys = new HashSet<string>(newTotals.Select(total => SummaryKey(total.Entity, total.Kind)));
            InsertNewSummaryRows(worksheet, startRow, newTotals, warnings);

            var dataRows = ReadSummaryMeta(worksheet).OrderBy(item => item.Row).ToList();
            for (var index = 0; index < dataRows.Count; index++)
            {
                var row = dataRows[index].Row;
                var info = dataRows[index];
                GroupSettlementTotal total;
                totalByKey.TryGetValue(SummaryKey(info.Entity, info.Kind), out total);
                worksheet.Cell(row, 1).Value = index + 1;
                WriteSummaryValues(worksheet, row, monthColumn, cumulativeColumn, total, info.Entity, info.Kind, month, newKeys.Contains(SummaryKey(info.Entity, info.Kind)));
            }

            var totalRow = FindSummaryTotalRow(worksheet, startRow);
            worksheet.Cell(totalRow, 1).Value = "合计";
            WriteSummaryTotalRow(worksheet, totalRow, cumulativeColumn);
            ApplySummaryDateFormats(worksheet, startRow, totalRow, cumulativeColumn);
            UpdateSummarySignatureDate(worksheet, month);
            ApplySummaryHeaderMerges(worksheet, monthColumn, cumulativeColumn + 9);
        }

        private static void WriteSummaryValues(
            IXLWorksheet worksheet,
            int row,
            int monthColumn,
            int cumulativeColumn,
            GroupSettlementTotal total,
            string entity,
            string kind,
            int month,
            bool isNewRow)
        {
            var proxyValue = total != null && total.Kind == "代理费" ? total.ExpectedNet : (double?)null;
            var interValue = total != null && total.Kind == "居间费" ? total.ExpectedNet : (double?)null;
            SetNullableNumber(worksheet.Cell(row, monthColumn), proxyValue);
            SetNullableNumber(worksheet.Cell(row, monthColumn + 1), interValue);
            worksheet.Cell(row, monthColumn + 2).Clear(XLClearOptions.Contents);

            var fee = (proxyValue ?? 0) + (interValue ?? 0);
            var loanTotalCell = worksheet.Cell(row, cumulativeColumn + 1);
            var hasLoan = CellHasContent(loanTotalCell);
            var loanTotal = hasLoan ? ClosedXmlUtil.CellNumber(loanTotalCell) : 0;
            var previousDeducted = hasLoan ? ClosedXmlUtil.CellNumber(worksheet.Cell(row, cumulativeColumn + 2)) : 0;
            var monthlyDeduction = ParseMonthlyDeduction(TextUtil.S(worksheet.Cell(row, cumulativeColumn + 5).GetFormattedString()));
            var remaining = Math.Max(loanTotal - previousDeducted, 0);
            var deduction = 0d;
            if (remaining > 0 && fee > 0)
            {
                deduction = Math.Round(Math.Min(Math.Min(fee, remaining), monthlyDeduction == 0 ? remaining : monthlyDeduction), 4);
            }

            SetNullableNumber(worksheet.Cell(row, monthColumn + 3), deduction == 0 ? (double?)null : deduction);
            worksheet.Cell(row, monthColumn + 4).FormulaA1 = SumFormula(row, monthColumn, monthColumn + 2);
            worksheet.Cell(row, monthColumn + 5).FormulaA1 = ClosedXmlUtil.ColumnLetter(monthColumn + 4) + row + "-" + ClosedXmlUtil.ColumnLetter(monthColumn + 3) + row;
            worksheet.Cell(row, cumulativeColumn).FormulaA1 = SumEverySix(row, 16, cumulativeColumn - 1);
            if (hasLoan)
            {
                worksheet.Cell(row, cumulativeColumn + 2).FormulaA1 = SumEverySix(row, 15, cumulativeColumn - 1);
                worksheet.Cell(row, cumulativeColumn + 3).FormulaA1 = ClosedXmlUtil.ColumnLetter(cumulativeColumn + 1) + row + "-" + ClosedXmlUtil.ColumnLetter(cumulativeColumn + 2) + row;
            }
            else
            {
                worksheet.Cell(row, cumulativeColumn + 2).Clear(XLClearOptions.Contents);
                worksheet.Cell(row, cumulativeColumn + 3).Clear(XLClearOptions.Contents);
            }

            if (isNewRow && !CellHasContent(worksheet.Cell(row, cumulativeColumn + 7)))
            {
                worksheet.Cell(row, cumulativeColumn + 7).Value = new DateTime(Year, month, 1);
            }

            var currentParty = TextUtil.S(worksheet.Cell(row, cumulativeColumn + 8).GetFormattedString());
            worksheet.Cell(row, cumulativeColumn + 8).Value = PaymentPartyFor(entity, kind, month, string.IsNullOrWhiteSpace(currentParty) ? "清辉" : currentParty);
        }

        private static int InsertMonthBlock(IXLWorksheet worksheet, int month)
        {
            var cumulativeColumn = SummaryColumn(worksheet, "累计代理费总计");
            var insertAt = cumulativeColumn;
            worksheet.Column(insertAt).InsertColumnsBefore(6);
            for (var offset = 0; offset < 6; offset++)
            {
                var sourceColumn = insertAt - 6 + offset;
                var targetColumn = insertAt + offset;
                worksheet.Column(sourceColumn).CopyTo(worksheet.Column(targetColumn));
                worksheet.Column(targetColumn).Unhide();
            }

            for (var column = insertAt - 6; column < insertAt; column++)
            {
                worksheet.Column(column).Hide();
            }

            for (var column = insertAt; column < insertAt + 6; column++)
            {
                worksheet.Column(column).Unhide();
            }

            worksheet.Cell(2, insertAt).Value = Year + "年" + month + "月";
            var labels = new[] { "代理费", "居间费", "退补电费", "当月抵扣", "费用合计" };
            for (var index = 0; index < labels.Length; index++)
            {
                worksheet.Cell(3, insertAt + index).Value = labels[index];
            }
            worksheet.Cell(2, insertAt + 5).Value = "当月实际支付";
            worksheet.Cell(3, insertAt + 5).Clear(XLClearOptions.Contents);
            return insertAt;
        }

        private static void DeleteSummaryRowsNotAllowed(IXLWorksheet worksheet, int startRow, ISet<string> allowedKeys)
        {
            if (allowedKeys == null)
            {
                return;
            }

            var totalRow = FindSummaryTotalRow(worksheet, startRow);
            for (var row = totalRow - 1; row >= startRow; row--)
            {
                var entity = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                var kind = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                if (string.IsNullOrWhiteSpace(entity) || !allowedKeys.Contains(SummaryKey(entity, kind)))
                {
                    worksheet.Row(row).Delete();
                }
            }
        }

        private static void InsertNewSummaryRows(IXLWorksheet worksheet, int startRow, IList<GroupSettlementTotal> newTotals, IList<string> warnings)
        {
            if (newTotals.Count == 0)
            {
                return;
            }

            var totalRow = FindSummaryTotalRow(worksheet, startRow);
            var templateRow = Math.Max(startRow, totalRow - 1);
            foreach (var total in newTotals)
            {
                worksheet.Row(totalRow).InsertRowsAbove(1);
                worksheet.Row(templateRow).CopyTo(worksheet.Row(totalRow));
                worksheet.Row(totalRow).Clear(XLClearOptions.Contents);
                worksheet.Cell(totalRow, 1).Value = totalRow - 3;
                worksheet.Cell(totalRow, 2).Value = total.Entity;
                worksheet.Cell(totalRow, 3).Value = total.Kind;
                worksheet.Cell(totalRow, 4).Value = "否";
                worksheet.Cell(totalRow, 6).Value = total.Entity;
                worksheet.Cell(totalRow, 7).Value = "平台";
                worksheet.Cell(totalRow, 8).Value = 0;
                worksheet.Cell(totalRow, 9).Value = 0.13;
                worksheet.Cell(totalRow, 10).Value = 0.13;
                worksheet.Cell(totalRow, 11).Value = total.Owner;
                warnings.Add("新增汇总主体：" + total.Kind + " " + total.Entity + "（负责人：" + total.Owner + "）");
                totalRow++;
            }
        }

        private static void WriteSummaryTotalRow(IXLWorksheet worksheet, int totalRow, int cumulativeColumn)
        {
            for (var column = 12; column <= cumulativeColumn + 3; column++)
            {
                var letter = ClosedXmlUtil.ColumnLetter(column);
                worksheet.Cell(totalRow, column).FormulaA1 = "SUM(" + letter + "4:" + letter + (totalRow - 1) + ")";
            }

            for (var column = cumulativeColumn + 4; column <= Math.Min(cumulativeColumn + 9, worksheet.LastColumnUsed()?.ColumnNumber() ?? cumulativeColumn + 9); column++)
            {
                worksheet.Cell(totalRow, column).Clear(XLClearOptions.Contents);
            }
        }

        private static void ApplySummaryDateFormats(IXLWorksheet worksheet, int startRow, int totalRow, int cumulativeColumn)
        {
            foreach (var column in new[] { cumulativeColumn + 4, cumulativeColumn + 6, cumulativeColumn + 7 })
            {
                for (var row = startRow; row < totalRow; row++)
                {
                    if (CellHasContent(worksheet.Cell(row, column)))
                    {
                        worksheet.Cell(row, column).Style.DateFormat.Format = "yyyy年m月";
                    }
                }
            }
        }

        private static void UpdateSummarySignatureDate(IXLWorksheet worksheet, int month)
        {
            var date = new DateTime(Year, month, 8).AddMonths(2);
            var text = "日期：" + date.Year + "年" + date.Month.ToString("00") + "月" + date.Day.ToString("00") + "日";
            IXLCell target = null;
            var maxColumn = 0;
            foreach (var cell in worksheet.CellsUsed())
            {
                if (!TextUtil.S(cell.GetFormattedString()).Contains("日期："))
                {
                    continue;
                }

                if (cell.Address.ColumnNumber > maxColumn)
                {
                    target = cell;
                    maxColumn = cell.Address.ColumnNumber;
                }
            }

            if (target != null)
            {
                target.Value = text;
            }
        }

        private static void ApplySummaryHeaderMerges(IXLWorksheet worksheet, int monthColumn, int lastRelevantColumn)
        {
            UnmergeIntersecting(worksheet, 1, 1, 1, lastRelevantColumn);
            worksheet.Range(1, 1, 1, lastRelevantColumn).Merge();

            UnmergeIntersecting(worksheet, 2, monthColumn, 2, monthColumn + 4);
            worksheet.Range(2, monthColumn, 2, monthColumn + 4).Merge();

            UnmergeIntersecting(worksheet, 2, monthColumn + 5, 3, monthColumn + 5);
            worksheet.Range(2, monthColumn + 5, 3, monthColumn + 5).Merge();
        }

        private static void UnmergeIntersecting(IXLWorksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            var ranges = worksheet.MergedRanges
                .Where(range => RangeIntersects(range, firstRow, firstColumn, lastRow, lastColumn))
                .ToList();
            foreach (var range in ranges)
            {
                range.Unmerge();
            }
        }

        private static bool RangeIntersects(IXLRange range, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            var address = range.RangeAddress;
            return address.FirstAddress.RowNumber <= lastRow
                && address.LastAddress.RowNumber >= firstRow
                && address.FirstAddress.ColumnNumber <= lastColumn
                && address.LastAddress.ColumnNumber >= firstColumn;
        }

        private static IList<SummaryMetaRow> ReadSummaryMeta(IXLWorksheet worksheet)
        {
            var cumulativeColumn = SummaryColumn(worksheet, "累计代理费总计");
            var totalRow = FindSummaryTotalRow(worksheet, 4);
            var rows = new List<SummaryMetaRow>();
            for (var row = 4; row < totalRow; row++)
            {
                var entity = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                var kind = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                if (string.IsNullOrWhiteSpace(entity) || entity.Contains("审核"))
                {
                    continue;
                }

                rows.Add(new SummaryMetaRow
                {
                    Row = row,
                    Entity = entity,
                    Kind = kind,
                    PaymentParty = TextUtil.S(worksheet.Cell(row, cumulativeColumn + 8).GetFormattedString())
                });
            }

            return rows;
        }

        private static bool CellHasContent(IXLCell cell)
        {
            if (!string.IsNullOrWhiteSpace(cell.FormulaA1))
            {
                return true;
            }

            return !string.IsNullOrWhiteSpace(cell.GetFormattedString());
        }

        private static string ResolveSummarySheetName(XLWorkbook workbook, string role, bool required)
        {
            var names = workbook.Worksheets.Select(sheet => sheet.Name).ToList();
            Func<string, string> clean = name => Regex.Replace(TextUtil.S(name), "\\s+", string.Empty);
            if (role == "main")
            {
                foreach (var exact in new[] { "汇总表", "代理费汇总表" })
                {
                    var matched = names.FirstOrDefault(name => clean(name) == exact);
                    if (matched != null)
                    {
                        return matched;
                    }
                }

                var candidate = names.FirstOrDefault(name => clean(name).Contains("汇总表") && !clean(name).Contains("清能") && !clean(name).Contains("清辉") && SummarySheetHasMarker(workbook.Worksheet(name)));
                if (candidate != null)
                {
                    return candidate;
                }
            }
            else if (role == "qingneng")
            {
                var candidate = names.FirstOrDefault(name => clean(name).Contains("清能") && clean(name).Contains("汇总") && SummarySheetHasMarker(workbook.Worksheet(name)));
                if (candidate != null)
                {
                    return candidate;
                }
            }
            else if (role == "qinghui")
            {
                var candidate = names.FirstOrDefault(name => clean(name).Contains("清辉") && clean(name).Contains("汇总") && SummarySheetHasMarker(workbook.Worksheet(name)));
                if (candidate != null)
                {
                    return candidate;
                }
            }

            if (required)
            {
                throw new InvalidOperationException("选择的汇总表模板缺少" + role + "汇总表。当前工作表：" + string.Join("、", names) + "。请在“上月/修正版汇总表”选择代理费汇总表文件。");
            }

            return null;
        }

        private static bool SummarySheetHasMarker(IXLWorksheet worksheet)
        {
            try
            {
                SummaryColumn(worksheet, "累计代理费总计");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static int SummaryColumn(IXLWorksheet worksheet, string header)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (TextUtil.S(worksheet.Cell(2, column).GetFormattedString()) == header)
                {
                    return column;
                }
            }

            throw new InvalidOperationException(worksheet.Name + " 未找到列：" + header);
        }

        private static int FindSummaryTotalRow(IXLWorksheet worksheet, int startRow)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? startRow;
            for (var row = startRow; row <= lastRow; row++)
            {
                if (TextUtil.S(worksheet.Cell(row, 1).GetFormattedString()) == "合计")
                {
                    return row;
                }
            }

            for (var row = startRow; row <= lastRow; row++)
            {
                for (var column = 12; column <= Math.Min(worksheet.LastColumnUsed()?.ColumnNumber() ?? 12, 80); column++)
                {
                    var formula = TextUtil.S(worksheet.Cell(row, column).FormulaA1);
                    if (formula.StartsWith("SUM(", StringComparison.OrdinalIgnoreCase))
                    {
                        return row;
                    }
                }
            }

            throw new InvalidOperationException(worksheet.Name + " 未找到合计行。");
        }

        private static int FindTotalRow(IXLWorksheet worksheet, int startRow)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? startRow;
            for (var row = startRow; row <= lastRow; row++)
            {
                if (TextUtil.S(worksheet.Cell(row, 1).GetFormattedString()) == "合计")
                {
                    return row;
                }
            }

            throw new InvalidOperationException(worksheet.Name + " 未找到合计行。");
        }

        private static Stage2Report CreateReport(
            Stage2Options options,
            IList<DetailSettlementRow> proxyRows,
            IList<DetailSettlementRow> interRows,
            IList<GroupSettlementTotal> totals,
            string summaryPath,
            IList<string> warnings,
            IList<string> missingOwners,
            IList<Stage2CheckIssue> auditIssues)
        {
            var reportPath = Path.Combine(options.OutputDirectory, options.Month + "月结算生成总报告.json");
            var report = new Stage2Report
            {
                Month = options.Month,
                Ledger = options.LedgerPath,
                ProxyTemplateDirectory = options.ProxyTemplateDirectory,
                IntermediaryTemplateDirectory = options.IntermediaryTemplateDirectory,
                SummaryTemplate = options.SummaryTemplatePath,
                OutputDirectory = options.OutputDirectory,
                Summary = summaryPath,
                ReportPath = reportPath,
                ProxyRows = proxyRows.Count,
                IntermediaryRows = interRows.Count,
                ProxyGroups = totals.Count(total => total.Kind == "代理费"),
                IntermediaryGroups = totals.Count(total => total.Kind == "居间费"),
                ProxyTotal = Math.Round(totals.Where(total => total.Kind == "代理费").Sum(total => total.ExpectedNet), 4),
                IntermediaryTotal = Math.Round(totals.Where(total => total.Kind == "居间费").Sum(total => total.ExpectedNet), 4)
            };
            report.Groups.AddRange(totals);
            report.Warnings.AddRange(warnings);
            report.MissingOwners.AddRange(missingOwners);
            report.AuditIssues.AddRange(auditIssues);
            return report;
        }

        private static void WriteReport(Stage2Options options, Stage2Report report)
        {
            File.WriteAllText(report.ReportPath, JsonConvert.SerializeObject(report, Formatting.Indented), System.Text.Encoding.UTF8);
        }

        private static void WriteWarnings(Stage2Options options, IList<string> warnings)
        {
            var path = Path.Combine(options.OutputDirectory, "自动生成汇总提示.txt");
            if (warnings.Count > 0)
            {
                File.WriteAllLines(path, warnings, System.Text.Encoding.UTF8);
            }
            else if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        private static void WriteAuditReport(Stage2Options options, Stage2Report report)
        {
            var path = Path.Combine(options.OutputDirectory, "阶段二校验报告.txt");
            if (report.AuditIssues.Count == 0 && report.Warnings.Count == 0 && report.MissingOwners.Count == 0)
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                return;
            }

            var lines = new List<string>
            {
                "阶段二校验报告",
                "结算月份：2026年" + options.Month + "月",
                "说明：文件已照常生成；当前分表和汇总表金额采用分表自算结果。",
                "提示：如果确认台账金额才是正确结果，请同步检查/修改对应分表和汇总表。",
                string.Empty
            };

            if (report.AuditIssues.Count > 0)
            {
                lines.Add("一、校验问题");
                for (var index = 0; index < report.AuditIssues.Count; index++)
                {
                    var issue = report.AuditIssues[index];
                    lines.Add((index + 1) + ". [" + issue.Severity + "] " + issue.Category);
                    if (!string.IsNullOrWhiteSpace(issue.Kind))
                    {
                        lines.Add("   类型：" + issue.Kind);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Owner) || !string.IsNullOrWhiteSpace(issue.Entity))
                    {
                        lines.Add("   负责人/主体：" + issue.Owner + " / " + issue.Entity);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Customer))
                    {
                        lines.Add("   客户：" + issue.Customer);
                    }
                    if (issue.LedgerRow > 0)
                    {
                        lines.Add("   台账行：" + issue.LedgerRow);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.PreviousValue) || !string.IsNullOrWhiteSpace(issue.CurrentValue))
                    {
                        lines.Add("   对比：" + issue.PreviousValue + "；" + issue.CurrentValue);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Message))
                    {
                        lines.Add("   问题：" + issue.Message);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Suggestion))
                    {
                        lines.Add("   建议：" + issue.Suggestion);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.TemplateFile))
                    {
                        lines.Add("   文件：" + issue.TemplateFile);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.SheetName))
                    {
                        lines.Add("   工作表：" + issue.SheetName);
                    }
                }
                lines.Add(string.Empty);
            }

            if (report.Warnings.Count > 0)
            {
                lines.Add("二、自动生成汇总提示");
                foreach (var warning in report.Warnings)
                {
                    lines.Add("- " + warning);
                }
                lines.Add(string.Empty);
            }

            if (report.MissingOwners.Count > 0)
            {
                lines.Add("三、负责人缺失");
                foreach (var missingOwner in report.MissingOwners)
                {
                    lines.Add("- " + missingOwner);
                }
            }

            File.WriteAllLines(path, lines, System.Text.Encoding.UTF8);
        }

        private static string PartyForSummaryTotal(GroupSettlementTotal total, IDictionary<string, string> partyByKey)
        {
            string party;
            return partyByKey.TryGetValue(SummaryKey(total.Entity, total.Kind), out party) ? party : "清辉";
        }

        private static string PaymentPartyFor(string entity, string kind, int month, string defaultParty)
        {
            Tuple<int, string> overrideValue;
            if (PaymentPartyOverrides.TryGetValue(PaymentKey(entity, kind), out overrideValue) && month >= overrideValue.Item1)
            {
                return overrideValue.Item2;
            }

            return string.IsNullOrWhiteSpace(defaultParty) ? "清辉" : defaultParty;
        }

        private static double ParseMonthlyDeduction(string text)
        {
            var matchWan = Regex.Match(TextUtil.S(text), "每月扣除([0-9.]+)万");
            if (matchWan.Success)
            {
                return Convert.ToDouble(matchWan.Groups[1].Value);
            }

            var matchYuan = Regex.Match(TextUtil.S(text), "每月扣除([0-9.]+)元");
            if (matchYuan.Success)
            {
                return Math.Round(Convert.ToDouble(matchYuan.Groups[1].Value) / 10000, 4);
            }

            return 0;
        }

        private static string SumEverySix(int row, int firstColumn, int beforeColumn)
        {
            var parts = new List<string>();
            for (var column = firstColumn; column < beforeColumn; column += 6)
            {
                parts.Add(ClosedXmlUtil.ColumnLetter(column) + row);
            }

            return parts.Count == 0 ? "0" : string.Join("+", parts);
        }

        private static string SumFormula(int row, int startColumn, int endColumn)
        {
            var parts = new List<string>();
            for (var column = startColumn; column <= endColumn; column++)
            {
                parts.Add(ClosedXmlUtil.ColumnLetter(column) + row);
            }

            return string.Join("+", parts);
        }

        private static void SetNullableNumber(IXLCell cell, double? value)
        {
            if (value.HasValue)
            {
                cell.Value = value.Value;
            }
            else
            {
                cell.Clear(XLClearOptions.Contents);
            }
        }

        private static void SaveWorkbook(XLWorkbook workbook, string outputPath)
        {
            workbook.CalculateMode = XLCalculateMode.Auto;
            workbook.SaveAs(outputPath, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        private static string TemplateKey(string kind, string owner, string entity)
        {
            return kind + "|" + NormalizeName(owner) + "|" + NormalizeName(entity);
        }

        private static string SummaryKey(string entity, string kind)
        {
            return NormalizeName(entity) + "|" + TextUtil.S(kind);
        }

        private static string PaymentKey(string entity, string kind)
        {
            return NormalizeName(entity) + "|" + TextUtil.S(kind);
        }

        private static string NormalizeName(string value)
        {
            var text = TextUtil.S(value);
            var match = Regex.Match(text, "^[\\u4e00-\\u9fa5]{2,4}（(.+)）$");
            if (match.Success)
            {
                text = match.Groups[1].Value;
            }

            text = Regex.Replace(text, "\\s+", string.Empty);
            text = text.Replace("（个体工商户）", string.Empty);
            text = text.Replace("(个体工商户)", string.Empty);
            text = text.Replace("绿洲森焱", "绿舟森焱");
            return text;
        }

        private static int ColumnNumber(string columnName)
        {
            var sum = 0;
            foreach (var c in columnName)
            {
                sum *= 26;
                sum += c - 'A' + 1;
            }

            return sum;
        }

        private static string RelativePath(string root, string path)
        {
            var rootUri = new Uri(AppendDirectorySeparatorChar(Path.GetFullPath(root)));
            var pathUri = new Uri(Path.GetFullPath(path));
            return Uri.UnescapeDataString(rootUri.MakeRelativeUri(pathUri).ToString()).Replace('/', Path.DirectorySeparatorChar);
        }

        private static string AppendDirectorySeparatorChar(string path)
        {
            return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ? path : path + Path.DirectorySeparatorChar;
        }

        private sealed class SummaryMetaRow
        {
            public int Row { get; set; }
            public string Entity { get; set; }
            public string Kind { get; set; }
            public string PaymentParty { get; set; }
        }

        private sealed class PreviousDetailRow
        {
            public int Row { get; set; }
            public string Customer { get; set; }
            public string SheetName { get; set; }
            public double Ratio { get; set; }
            public double UnitPrice { get; set; }
            public double TaxRate { get; set; }
        }
    }
}
