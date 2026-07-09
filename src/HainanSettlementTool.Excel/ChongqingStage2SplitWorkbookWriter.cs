using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class ChongqingStage2SplitWorkbookWriter
    {
        public static List<GroupSettlementTotal> BuildSplitFiles(
            ChongqingStage2Options options,
            IList<ChongqingSettlementDetail> details,
            IList<string> warnings,
            IList<ChongqingStage2CheckIssue> auditIssues)
        {
            var templateMap = BuildTemplateIndex(options, warnings);
            var groups = new List<GroupSettlementTotal>();
            foreach (var group in details
                .GroupBy(detail => ChongqingStage2Keys.TemplateKey(detail.Kind, detail.Owner, detail.Entity))
                .OrderBy(group => group.First().Kind)
                .ThenBy(group => group.First().Owner)
                .ThenBy(group => group.First().Entity))
            {
                var first = group.First();
                bool matchedTemplate;
                var outputPath = EnsureOutputWorkbook(templateMap, options, first.Kind, first.Owner, first.Entity, warnings, out matchedTemplate);
                FileAccessGuard.RequireWritableWorkbook(outputPath, ChongqingStage2ExcelUtil.KindShort(first.Kind) + "分表输出文件");

                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = PrepareMonthSheet(workbook, options.Month);
                    WriteSplitSheet(worksheet, first.Kind, first.Entity, options.Month, group.ToList(), outputPath, warnings, auditIssues);
                    if (!matchedTemplate)
                    {
                        KeepOnlyCurrentMonthSheet(workbook, worksheet);
                    }

                    ChongqingStage2ExcelUtil.SaveWorkbook(workbook, outputPath);
                }

                groups.Add(new GroupSettlementTotal
                {
                    Kind = first.Kind,
                    Owner = first.Owner,
                    Entity = first.Entity,
                    DisplayEntity = first.Entity,
                    Rows = group.Count(),
                    ExpectedNet = Math.Round(group.Sum(item => item.ExpectedNet), 4),
                    OutputFile = outputPath
                });
            }

            return groups;
        }

        private static Dictionary<string, string> BuildTemplateIndex(ChongqingStage2Options options, IList<string> warnings)
        {
            var result = new Dictionary<string, string>();
            IndexTemplateRoot(result, ChongqingStage2SettlementKinds.Proxy, options.ProxyTemplateDirectory, warnings);
            IndexTemplateRoot(result, ChongqingStage2SettlementKinds.Intermediary, options.IntermediaryTemplateDirectory, warnings);
            IndexTemplateRoot(result, ChongqingStage2SettlementKinds.Refund, options.RefundTemplateDirectory, warnings);
            return result;
        }

        private static void IndexTemplateRoot(IDictionary<string, string> result, string kind, string root, IList<string> warnings)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return;
            }

            foreach (var path in Directory.GetFiles(root, "*.xlsx", SearchOption.AllDirectories))
            {
                var name = Path.GetFileName(path);
                if (name.StartsWith("~$", StringComparison.Ordinal) || name.StartsWith("._", StringComparison.Ordinal))
                {
                    continue;
                }

                try
                {
                    using (var workbook = ChongqingStage2ExcelUtil.OpenWorkbookShared(path))
                    {
                        var entity = ExtractEntityFromWorkbook(workbook);
                        if (string.IsNullOrWhiteSpace(entity))
                        {
                            warnings.Add("重庆" + ChongqingStage2ExcelUtil.KindShort(kind) + "模板未识别到有效主体抬头，已跳过：" + path);
                            continue;
                        }

                        var owner = new DirectoryInfo(Path.GetDirectoryName(path)).Name;
                        result[ChongqingStage2Keys.TemplateKey(kind, owner, entity)] = path;
                        result[ChongqingStage2Keys.TemplateKey(kind, string.Empty, entity)] = path;
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add("重庆" + ChongqingStage2ExcelUtil.KindShort(kind) + "模板读取失败，已跳过：" + path + "；" + ex.Message);
                }
            }
        }

        private static string ExtractEntityFromWorkbook(XLWorkbook workbook)
        {
            foreach (var sheet in workbook.Worksheets
                .Select(sheet =>
                {
                    int sheetMonth;
                    return new { Sheet = sheet, Matched = TryParseMonthSheet(sheet.Name, out sheetMonth), Month = sheetMonth };
                })
                .OrderByDescending(item => item.Matched)
                .ThenByDescending(item => item.Month)
                .Select(item => item.Sheet))
            {
                var entity = ExtractEntityFromSplitSheet(sheet);
                if (!string.IsNullOrWhiteSpace(entity))
                {
                    return entity;
                }
            }

            return string.Empty;
        }

        private static string EnsureOutputWorkbook(
            IDictionary<string, string> templateMap,
            ChongqingStage2Options options,
            string kind,
            string owner,
            string entity,
            IList<string> warnings,
            out bool matchedTemplate)
        {
            string source;
            if (templateMap.TryGetValue(ChongqingStage2Keys.TemplateKey(kind, owner, entity), out source)
                || templateMap.TryGetValue(ChongqingStage2Keys.TemplateKey(kind, string.Empty, entity), out source))
            {
                var sourceRoot = ChongqingStage2ExcelUtil.TemplateRootFor(options, kind);
                var targetRoot = ChongqingStage2ExcelUtil.OutputRootFor(options, kind);
                var relative = ChongqingStage2ExcelUtil.RelativePath(sourceRoot, source);
                var target = Path.Combine(targetRoot, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                if (!File.Exists(target))
                {
                    ChongqingStage2ExcelUtil.CopyWorkbookShared(source, target, overwrite: false);
                }

                matchedTemplate = true;
                return target;
            }

            source = templateMap
                .Where(item => item.Key.StartsWith(kind + "|", StringComparison.Ordinal))
                .Select(item => item.Value)
                .FirstOrDefault();
            if (string.IsNullOrWhiteSpace(source))
            {
                throw new InvalidOperationException("没有可用的重庆" + ChongqingStage2ExcelUtil.KindShort(kind) + "分表模板。");
            }

            var folder = Path.Combine(ChongqingStage2ExcelUtil.OutputRootFor(options, kind), TextUtil.SafeFileName(owner));
            Directory.CreateDirectory(folder);
            var targetPath = Path.Combine(folder, TextUtil.SafeFileName(entity) + " 2026重庆.xlsx");
            if (!File.Exists(targetPath))
            {
                ChongqingStage2ExcelUtil.CopyWorkbookShared(source, targetPath, overwrite: false);
            }

            matchedTemplate = false;
            warnings.Add("新增" + kind + "分表主体：" + entity + "（负责人：" + owner + "），已借用同类模板生成，请人工复核抬头、签章和付款信息。");
            return targetPath;
        }

        private static void WriteSplitSheet(
            IXLWorksheet worksheet,
            string kind,
            string entity,
            int month,
            IList<ChongqingSettlementDetail> rows,
            string outputPath,
            IList<string> warnings,
            IList<ChongqingStage2CheckIssue> auditIssues)
        {
            SetSplitTopTitles(worksheet, kind, entity, month);
            var totalRow = AdjustDetailRows(worksheet, rows.Count);
            for (var index = 0; index < rows.Count; index++)
            {
                var excelRow = ChongqingStage2Layout.DetailDataStartRow + index;
                if (kind == ChongqingStage2SettlementKinds.Refund)
                {
                    WriteRefundDetailRow(worksheet, excelRow, index + 1, rows[index]);
                }
                else
                {
                    WriteProxyLikeDetailRow(worksheet, excelRow, index + 1, rows[index]);
                }

                AddLedgerDifferenceIssue(rows[index], outputPath, worksheet.Name, auditIssues);
            }

            WriteSplitTotalRow(worksheet, totalRow, kind, rows.Count);
            if (kind == ChongqingStage2SettlementKinds.Refund)
            {
                SyncRefundExtraPowerRows(worksheet, rows, outputPath, warnings);
            }
        }

        private static void WriteProxyLikeDetailRow(IXLWorksheet worksheet, int row, int sequence, ChongqingSettlementDetail detail)
        {
            worksheet.Cell(row, 1).Value = sequence;
            worksheet.Cell(row, 2).Value = detail.Customer;
            worksheet.Cell(row, 3).Value = Math.Round(detail.Total, 4);
            worksheet.Cell(row, 4).Value = Math.Round(detail.Sharp, 4);
            worksheet.Cell(row, 5).Value = Math.Round(detail.Peak, 4);
            worksheet.Cell(row, 6).Value = Math.Round(detail.Flat, 4);
            worksheet.Cell(row, 7).Value = Math.Round(detail.Valley, 4);
            worksheet.Cell(row, 8).Value = Math.Round(detail.Ratio, 6);
            worksheet.Cell(row, 9).Value = Math.Round(detail.UnitPrice, 6);
            worksheet.Cell(row, 10).FormulaA1 = "ROUND(C" + row + "*H" + row + "*I" + row + "/10,4)";
            worksheet.Cell(row, 11).FormulaA1 = "J" + row + "-L" + row;
            worksheet.Cell(row, 12).Value = Math.Round(detail.RecoverShortfall, 4);
            worksheet.Cell(row, 13).FormulaA1 = "ROUND(K" + row + "/1.13*O" + row + ",4)";
            worksheet.Cell(row, 14).FormulaA1 = "K" + row + "-M" + row;
            worksheet.Cell(row, 15).Value = Math.Round(detail.TaxRate, 6);
        }

        private static void WriteRefundDetailRow(IXLWorksheet worksheet, int row, int sequence, ChongqingSettlementDetail detail)
        {
            worksheet.Cell(row, 1).Value = sequence;
            worksheet.Cell(row, 2).Value = detail.Customer;
            worksheet.Cell(row, 3).FormulaA1 = "SUM(D" + row + ":G" + row + ")";
            worksheet.Cell(row, 4).Value = Math.Round(detail.Sharp, 4);
            worksheet.Cell(row, 5).Value = Math.Round(detail.Peak, 4);
            worksheet.Cell(row, 6).Value = Math.Round(detail.Flat, 4);
            worksheet.Cell(row, 7).Value = Math.Round(detail.Valley, 4);
            worksheet.Cell(row, 8).Value = Math.Round(detail.Ratio, 6);
            worksheet.Cell(row, 9).Value = Math.Round(detail.RefundSharpPrice, 6);
            worksheet.Cell(row, 10).Value = Math.Round(detail.RefundPeakPrice, 6);
            worksheet.Cell(row, 11).Value = Math.Round(detail.RefundFlatPrice, 6);
            worksheet.Cell(row, 12).Value = Math.Round(detail.RefundValleyPrice, 6);
            worksheet.Cell(row, 13).FormulaA1 = "ROUND((D" + row + "*H" + row + "*I" + row + "+E" + row + "*H" + row + "*J" + row + "+F" + row + "*H" + row + "*K" + row + "+G" + row + "*H" + row + "*L" + row + ")/10,4)";
            worksheet.Cell(row, 14).FormulaA1 = "ROUND(M" + row + "/1.13*P" + row + ",4)";
            worksheet.Cell(row, 15).FormulaA1 = "M" + row + "-N" + row;
            worksheet.Cell(row, 16).Value = Math.Round(detail.TaxRate, 6);
        }

        private static void WriteSplitTotalRow(IXLWorksheet worksheet, int totalRow, string kind, int rowCount)
        {
            worksheet.Cell(totalRow, 1).Value = "合计";
            worksheet.Cell(totalRow, 2).Clear(XLClearOptions.Contents);
            if (rowCount == 0)
            {
                return;
            }

            var last = ChongqingStage2Layout.DetailDataStartRow + rowCount - 1;
            var sumColumns = kind == ChongqingStage2SettlementKinds.Refund
                ? new[] { 3, 4, 5, 6, 7, 13, 14, 15 }
                : new[] { 3, 4, 5, 6, 7, 10, 11, 12, 13, 14 };
            foreach (var column in sumColumns)
            {
                var letter = ClosedXmlUtil.ColumnLetter(column);
                worksheet.Cell(totalRow, column).FormulaA1 = "SUM(" + letter + ChongqingStage2Layout.DetailDataStartRow + ":" + letter + last + ")";
            }
        }

        private static IXLWorksheet PrepareMonthSheet(XLWorkbook workbook, int month)
        {
            var title = ChongqingStage2ExcelUtil.MonthSheetName(month);
            var existing = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == title);
            if (existing != null)
            {
                return existing;
            }

            var source = PreviousMonthSheet(workbook, month, title);
            if (source == null)
            {
                source = LastMonthSheet(workbook);
            }

            return source.CopyTo(title);
        }

        private static IXLWorksheet LastMonthSheet(XLWorkbook workbook)
        {
            var monthSheets = workbook.Worksheets
                .Select(sheet =>
                {
                    int sheetMonth;
                    return new { Sheet = sheet, Matched = TryParseMonthSheet(sheet.Name, out sheetMonth), Month = sheetMonth };
                })
                .Where(item => item.Matched)
                .OrderBy(item => item.Month)
                .ToList();
            return monthSheets.Count == 0 ? workbook.Worksheets.Last() : monthSheets.Last().Sheet;
        }

        private static IXLWorksheet PreviousMonthSheet(XLWorkbook workbook, int month, string targetTitle)
        {
            return workbook.Worksheets
                .Select(sheet =>
                {
                    int sheetMonth;
                    return new { Sheet = sheet, Matched = TryParseMonthSheet(sheet.Name, out sheetMonth), Month = sheetMonth };
                })
                .Where(item => item.Matched && item.Month < month && item.Sheet.Name != targetTitle)
                .OrderBy(item => item.Month)
                .Select(item => item.Sheet)
                .LastOrDefault();
        }

        private static bool TryParseMonthSheet(string name, out int month)
        {
            month = 0;
            var match = Regex.Match(TextUtil.S(name), "^(\\d{1,2})(?:\\s*\\(\\d+\\))?$");
            return match.Success && int.TryParse(match.Groups[1].Value, out month);
        }

        private static void KeepOnlyCurrentMonthSheet(XLWorkbook workbook, IXLWorksheet currentMonthSheet)
        {
            var currentName = currentMonthSheet.Name;
            foreach (var sheet in workbook.Worksheets.Where(sheet => sheet.Name != currentName).ToList())
            {
                sheet.Delete();
            }
        }

        private static int AdjustDetailRows(IXLWorksheet worksheet, int count)
        {
            var totalRow = FindSplitTotalRow(worksheet);
            var existing = totalRow - ChongqingStage2Layout.DetailDataStartRow;
            if (count > existing)
            {
                worksheet.Row(totalRow).InsertRowsAbove(count - existing);
                for (var row = totalRow; row < totalRow + count - existing; row++)
                {
                    worksheet.Row(ChongqingStage2Layout.DetailDataStartRow).CopyTo(worksheet.Row(row));
                }
            }
            else if (count < existing)
            {
                worksheet.Rows(ChongqingStage2Layout.DetailDataStartRow + count, totalRow - 1).Delete();
            }

            return FindSplitTotalRow(worksheet);
        }

        private static int FindSplitTotalRow(IXLWorksheet worksheet)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? ChongqingStage2Layout.DetailDataStartRow;
            for (var row = ChongqingStage2Layout.DetailDataStartRow; row <= lastRow; row++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 1)) == "合计")
                {
                    return row;
                }

                for (var column = 3; column <= Math.Min(16, worksheet.LastColumnUsed()?.ColumnNumber() ?? 16); column++)
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

        private static void SyncRefundExtraPowerRows(
            IXLWorksheet worksheet,
            IList<ChongqingSettlementDetail> rows,
            string outputPath,
            IList<string> warnings)
        {
            var firstTotalRow = FindSplitTotalRow(worksheet);
            var extraRows = FindRefundExtraPowerRows(worksheet, firstTotalRow);
            if (extraRows.Count == 0)
            {
                return;
            }

            var entity = rows.Count > 0 ? rows[0].Entity : ExtractEntityFromSplitSheet(worksheet);
            if (rows.Count != 1)
            {
                warnings.Add("重庆退补分表“" + entity + "”" + worksheet.Name + "月检测到" + extraRows.Count + "行额外扣减块，但当前主体有" + rows.Count + "条退补明细，无法安全自动匹配，已保留额外块 C-G 原值；请人工处理。文件：" + outputPath);
                return;
            }

            var detail = rows[0];
            foreach (var row in extraRows)
            {
                worksheet.Cell(row, 3).FormulaA1 = "SUM(D" + row + ":G" + row + ")";
                worksheet.Cell(row, 4).Value = Math.Round(detail.Sharp, 4);
                worksheet.Cell(row, 5).Value = Math.Round(detail.Peak, 4);
                worksheet.Cell(row, 6).Value = Math.Round(detail.Flat, 4);
                worksheet.Cell(row, 7).Value = Math.Round(detail.Valley, 4);
            }

            warnings.Add("重庆退补分表“" + entity + "”" + worksheet.Name + "月检测到" + extraRows.Count + "行额外扣减块，已同步 C-G 当月电量；H列以后、汇总表当月抵扣和实际支付仍按模板保留，请人工复核。文件：" + outputPath);
        }

        private static List<int> FindRefundExtraPowerRows(IXLWorksheet worksheet, int firstTotalRow)
        {
            var result = new List<int>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? firstTotalRow;
            for (var row = firstTotalRow + 1; row <= lastRow; row++)
            {
                var label = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2));
                if (label.IndexOf("当月应扣", StringComparison.Ordinal) >= 0
                    || label.IndexOf("电表改造", StringComparison.Ordinal) >= 0)
                {
                    result.Add(row);
                }
            }

            return result;
        }

        private static void SetSplitTopTitles(IXLWorksheet worksheet, string kind, string entity, int month)
        {
            if (kind == ChongqingStage2SettlementKinds.Refund)
            {
                worksheet.Cell("A1").Value = "退补电费结算单";
                worksheet.Cell("A2").Value = "名称:" + entity;
            }
            else
            {
                worksheet.Cell("A1").Value = ChongqingStage2ExcelUtil.KindShort(kind) + "费用结算单";
                worksheet.Cell("A2").Value = "代理名称:" + entity;
            }

            SetFirstCellContaining(worksheet, "所属期", "所属期：" + ChongqingStage2Layout.Year + " 年 " + month.ToString("00", CultureInfo.InvariantCulture) + " 月");
            var nextMonth = month == 12 ? 1 : month + 1;
            var year = month == 12 ? ChongqingStage2Layout.Year + 1 : ChongqingStage2Layout.Year;
            SetFirstCellContaining(worksheet, "结算日期", "结算日期：" + year + " 年 " + nextMonth.ToString("00", CultureInfo.InvariantCulture) + " 月 15 日");
        }

        private static void SetFirstCellContaining(IXLWorksheet worksheet, string needle, string value)
        {
            var maxRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 1, 8);
            var maxColumn = Math.Min(worksheet.LastColumnUsed()?.ColumnNumber() ?? 1, 24);
            for (var row = 1; row <= maxRow; row++)
            {
                for (var column = 1; column <= maxColumn; column++)
                {
                    if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, column)).Contains(needle))
                    {
                        worksheet.Cell(row, column).Value = value;
                        return;
                    }
                }
            }
        }

        private static string ExtractEntityFromSplitSheet(IXLWorksheet sheet)
        {
            var text = ChongqingStage2ExcelUtil.CellText(sheet.Cell("A2"));
            foreach (var prefix in new[] { "代理名称:", "代理名称：", "居间名称:", "居间名称：", "名称:", "名称：" })
            {
                if (text.StartsWith(prefix, StringComparison.Ordinal))
                {
                    return TextUtil.S(text.Substring(prefix.Length));
                }
            }

            return string.Empty;
        }

        private static void AddLedgerDifferenceIssue(
            ChongqingSettlementDetail detail,
            string outputPath,
            string sheetName,
            IList<ChongqingStage2CheckIssue> auditIssues)
        {
            if (Math.Abs(detail.LedgerNet - detail.ExpectedNet) <= Stage2SettlementCalculator.AmountTolerance)
            {
                return;
            }

            auditIssues.Add(new ChongqingStage2CheckIssue
            {
                Severity = "确认",
                Category = "台账与分表金额不一致",
                Kind = "台账与分表金额不一致",
                SettlementKind = detail.Kind,
                Customer = detail.Customer,
                Owner = detail.Owner,
                Entity = detail.Entity,
                LedgerRow = detail.LedgerRow,
                TemplateFile = outputPath,
                SheetName = sheetName,
                PreviousValue = "台账：" + Stage2SettlementCalculator.FormatAmount(detail.LedgerNet),
                CurrentValue = "分表自算：" + Stage2SettlementCalculator.FormatAmount(detail.ExpectedNet),
                Message = "重庆" + detail.Kind + "主体“" + detail.Entity + "”下客户“" + detail.Customer + "”金额与分表公式结果不一致。",
                Suggestion = "请检查台账第" + detail.LedgerRow + "行电量、占比、单价、税点、少回收电能量电费和公式缓存；本次汇总按分表公式结果写入。"
            });
        }
    }
}
