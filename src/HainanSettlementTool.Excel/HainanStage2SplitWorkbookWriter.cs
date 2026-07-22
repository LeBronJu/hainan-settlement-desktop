using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class HainanStage2SplitWorkbookWriter
    {
        internal static void VerifyGeneratedSplitWorkbooks(
            IList<HainanStage2SubjectGroup> subjectGroups,
            IList<GroupSettlementTotal> totals,
            int month)
        {
            var expectedGroups = subjectGroups.ToDictionary(
                group => HainanStage2ExcelUtil.SummaryKey(group.Entity, group.SettlementKind));
            var generatedGroups = totals
                .GroupBy(total => HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind))
                .ToDictionary(group => group.Key, group => group.ToList());
            if (generatedGroups.Any(group => group.Value.Count != 1)
                || !new HashSet<string>(expectedGroups.Keys).SetEquals(generatedGroups.Keys))
            {
                throw new InvalidDataException("海南阶段二分表主体集合与预检结果不一致。");
            }

            var outputPaths = totals
                .Select(total => Path.GetFullPath(total.OutputFile))
                .ToList();
            if (outputPaths.Count != outputPaths.Distinct(StringComparer.OrdinalIgnoreCase).Count())
            {
                throw new InvalidDataException("海南阶段二有多个主体指向同一分表文件。");
            }

            foreach (var pair in expectedGroups)
            {
                var expected = pair.Value;
                var generated = generatedGroups[pair.Key].Single();
                if (generated.Rows != expected.Rows.Count
                    || Math.Abs(generated.ExpectedNet - expected.Rows.Sum(row => row.ExpectedNet))
                        > Stage2SettlementCalculator.AmountTolerance)
                {
                    throw new InvalidDataException("海南阶段二分表分组摘要与台账不一致：" + generated.Kind + " " + generated.Entity);
                }

                using (var stream = File.Open(generated.OutputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var workbook = new XLWorkbook(stream))
                {
                    var sheetName = month + "月";
                    var matches = workbook.Worksheets.Where(sheet => sheet.Name == sheetName).ToList();
                    if (matches.Count != 1)
                    {
                        throw new InvalidDataException("海南阶段二分表目标月工作表数量不是 1：" + generated.OutputFile);
                    }

                    var worksheet = matches[0];
                    var displayEntity = TextUtil.S(worksheet.Cell("A2").GetFormattedString());
                    if (displayEntity.StartsWith("代理名称:", StringComparison.Ordinal))
                    {
                        displayEntity = displayEntity.Substring("代理名称:".Length);
                    }

                    if (HainanStage2ExcelUtil.NormalizeName(displayEntity)
                        != HainanStage2ExcelUtil.NormalizeName(generated.DisplayEntity))
                    {
                        throw new InvalidDataException("海南阶段二分表主体与分组不一致：" + generated.OutputFile);
                    }

                    var totalRow = HainanStage2ExcelUtil.FindTotalRow(worksheet, HainanStage2ExcelUtil.DataStartRow);
                    if (totalRow - HainanStage2ExcelUtil.DataStartRow != expected.Rows.Count)
                    {
                        throw new InvalidDataException("海南阶段二分表明细行数与台账不一致：" + generated.OutputFile);
                    }

                    for (var index = 0; index < expected.Rows.Count; index++)
                    {
                        var expectedRow = expected.Rows[index];
                        var row = HainanStage2ExcelUtil.DataStartRow + index;
                        if (HainanStage2ExcelUtil.NormalizeName(worksheet.Cell(row, 2).GetFormattedString())
                            != HainanStage2ExcelUtil.NormalizeName(expectedRow.Customer)
                            || !AmountsEqual(ClosedXmlUtil.CellNumber(worksheet.Cell(row, 3)), expectedRow.Total)
                            || !AmountsEqual(ClosedXmlUtil.CellNumber(worksheet.Cell(row, 10)), expectedRow.Ratio)
                            || !AmountsEqual(ClosedXmlUtil.CellNumber(worksheet.Cell(row, 11)), expectedRow.UnitPrice)
                            || !AmountsEqual(ClosedXmlUtil.CellNumber(worksheet.Cell(row, 16)), expectedRow.ExpectedNet)
                            || !HainanStage2ExcelUtil.TaxRatesEqual(
                                ClosedXmlUtil.CellNumber(worksheet.Cell(row, 17)),
                                expectedRow.TaxRate))
                        {
                            throw new InvalidDataException(
                                "海南阶段二分表明细与台账不一致："
                                + generated.Kind + " " + generated.Entity + "，第" + row + "行。");
                        }
                    }

                    if (!AmountsEqual(
                        ClosedXmlUtil.CellNumber(worksheet.Cell(totalRow, 16)),
                        generated.ExpectedNet))
                    {
                        throw new InvalidDataException("海南阶段二分表合计与分组金额不一致：" + generated.OutputFile);
                    }
                }
            }
        }

        internal static List<GroupSettlementTotal> BuildSplitFiles(
            HainanStage2Options options,
            IList<HainanStage2SubjectGroup> subjectGroups,
            HainanStage2TemplateCatalog templateCatalog)
        {
            var totals = new List<GroupSettlementTotal>();
            foreach (var group in subjectGroups
                .OrderBy(item => item.Kind)
                .ThenBy(item => item.FirstLedgerRow))
            {
                bool matchedTemplate;
                var outputPath = EnsureOutputWorkbook(templateCatalog, options, group.Kind, group.Owner, group.Entity, out matchedTemplate);
                FileAccessGuard.RequireWritableWorkbook(outputPath, group.Kind + "分表输出文件");

                using (var workbook = new XLWorkbook(outputPath))
                {
                    var displayEntity = matchedTemplate ? PriorSheetDisplayEntity(workbook, options.Month) : group.Entity;
                    var worksheet = PrepareMonthSheet(workbook, options.Month);
                    WriteDetailSheet(worksheet, group.Kind, group.Entity, options.Month, group.Rows, displayEntity);
                    if (!matchedTemplate)
                    {
                        KeepOnlyCurrentMonthSheet(workbook, worksheet);
                    }

                    HainanStage2ExcelUtil.SaveWorkbook(workbook, outputPath);

                    totals.Add(new GroupSettlementTotal
                    {
                        Kind = group.SettlementKind,
                        Owner = group.Owner,
                        Entity = group.Entity,
                        DisplayEntity = string.IsNullOrWhiteSpace(displayEntity) ? group.Entity : displayEntity,
                        Rows = group.Rows.Count,
                        ExpectedNet = Math.Round(group.Rows.Sum(row => row.ExpectedNet), 4),
                        OutputFile = outputPath
                    });
                }
            }

            return totals;
        }

        internal static string PlanOutputPath(
            HainanStage2Options options,
            HainanStage2SubjectGroup group,
            HainanStage2TemplateCatalog templateCatalog)
        {
            var exactCandidates = templateCatalog.ExactCandidates(group.Kind, group.Entity);
            var baseRoot = Path.Combine(
                options.OutputDirectory,
                group.Kind == "代理" ? "2026年代理 - 海南" : "2026年居间 - 海南");
            var folder = Path.Combine(baseRoot, TextUtil.SafeFileName(group.Owner) + " - 海南2026");
            var fileName = exactCandidates.Count == 1
                ? Path.GetFileName(exactCandidates[0].Path)
                : TextUtil.SafeFileName(group.Entity) + " 2026海南.xlsx";
            return Path.GetFullPath(Path.Combine(folder, fileName));
        }

        private static string EnsureOutputWorkbook(
            HainanStage2TemplateCatalog templateCatalog,
            HainanStage2Options options,
            string kind,
            string owner,
            string entity,
            out bool matchedTemplate)
        {
            var exactCandidates = templateCatalog.ExactCandidates(kind, entity);
            if (exactCandidates.Count > 1)
            {
                throw new InvalidOperationException(kind + "费主体“" + entity + "”匹配到多个上月分表，已停止生成。请先处理重复模板。");
            }

            var group = new HainanStage2SubjectGroup
            {
                Kind = kind,
                Entity = entity,
                Owner = owner
            };
            var target = PlanOutputPath(options, group, templateCatalog);
            var folder = Path.GetDirectoryName(target);
            Directory.CreateDirectory(folder);

            if (exactCandidates.Count == 1)
            {
                var source = exactCandidates[0].Path;
                if (!File.Exists(target))
                {
                    File.Copy(source, target, false);
                }

                matchedTemplate = true;
                return target;
            }

            var newTarget = target;
            if (File.Exists(newTarget))
            {
                matchedTemplate = false;
                return newTarget;
            }

            var candidate = templateCatalog.UniqueForKind(kind);
            if (candidate == null)
            {
                throw new InvalidOperationException("没有唯一可用的" + kind + "分表模板。");
            }

            File.Copy(candidate.Path, newTarget, false);
            matchedTemplate = false;
            return newTarget;
        }

        private static IXLWorksheet PrepareMonthSheet(XLWorkbook workbook, int month)
        {
            var title = month + "月";
            var existing = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == title);
            var source = HainanStage2ExcelUtil.PreviousMonthSheet(workbook, month, title);
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

        private static void KeepOnlyCurrentMonthSheet(XLWorkbook workbook, IXLWorksheet currentMonthSheet)
        {
            var currentName = currentMonthSheet.Name;
            foreach (var sheet in workbook.Worksheets.Where(sheet => sheet.Name != currentName).ToList())
            {
                sheet.Delete();
            }
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
            IList<HainanStage2DetailSettlementRow> rows,
            string displayEntity)
        {
            SetTopTitles(worksheet, kind, entity, month, displayEntity);
            var totalRow = AdjustDetailRows(worksheet, rows.Count);
            for (var index = 0; index < rows.Count; index++)
            {
                var row = rows[index];
                var excelRow = HainanStage2ExcelUtil.DataStartRow + index;
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
            }

            if (rows.Count > 0)
            {
                var last = HainanStage2ExcelUtil.DataStartRow + rows.Count - 1;
                worksheet.Cell(totalRow, 1).Value = "合计";
                for (var column = 3; column <= 7; column++)
                {
                    var letter = ClosedXmlUtil.ColumnLetter(column);
                    SetDetailFormula(worksheet.Cell(totalRow, column), "SUM(" + letter + HainanStage2ExcelUtil.DataStartRow + ":" + letter + last + ")");
                }

                for (var column = 12; column <= 16; column++)
                {
                    var letter = ClosedXmlUtil.ColumnLetter(column);
                    SetDetailFormula(worksheet.Cell(totalRow, column), "SUM(" + letter + HainanStage2ExcelUtil.DataStartRow + ":" + letter + last + ")");
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
                    totalRow = HainanStage2ExcelUtil.FindTotalRow(candidate, HainanStage2ExcelUtil.DataStartRow);
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
            var row = Math.Max(HainanStage2ExcelUtil.DataStartRow, totalRow - 1);
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
                if (text.Contains("结算日期"))
                {
                    continue;
                }

                if (text.Contains("日期："))
                {
                    var updated = ShiftSignatureDateText(text, 1);
                    if (!string.IsNullOrWhiteSpace(updated))
                    {
                        cell.Value = updated;
                        continue;
                    }
                }

                DateTime date;
                if (TryGetDateFormattedCellValue(cell, out date))
                {
                    cell.Value = date.AddMonths(1);
                }
            }
        }

        private static bool TryGetDateFormattedCellValue(IXLCell cell, out DateTime date)
        {
            date = default(DateTime);
            var format = TextUtil.S(cell.Style.DateFormat.Format).ToLowerInvariant();
            if (!format.Contains("y") || !format.Contains("m") || !format.Contains("d"))
            {
                return false;
            }

            try
            {
                date = cell.GetDateTime();
                return true;
            }
            catch
            {
            }

            var serial = ClosedXmlUtil.CellNumber(cell);
            if (serial <= 20000 || serial >= 60000)
            {
                return false;
            }

            date = DateTime.FromOADate(serial);
            return true;
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

        private static bool AmountsEqual(double left, double right)
        {
            return Math.Abs(left - right) <= Stage2SettlementCalculator.AmountTolerance;
        }

        private static void SetTopTitles(IXLWorksheet worksheet, string kind, string entity, int month, string displayEntity)
        {
            worksheet.Cell("A1").Value = kind + "费用结算单";
            worksheet.Cell("A2").Value = "代理名称:" + (string.IsNullOrWhiteSpace(displayEntity) ? entity : displayEntity);
            SetFirstCellContaining(worksheet, "所属期", "所属期：" + HainanStage2ExcelUtil.Year + " 年 " + month.ToString("00") + " 月");
            var nextMonth = month == 12 ? 1 : month + 1;
            var year = month == 12 ? HainanStage2ExcelUtil.Year + 1 : HainanStage2ExcelUtil.Year;
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
            var totalRow = HainanStage2ExcelUtil.FindTotalRow(worksheet, HainanStage2ExcelUtil.DataStartRow);
            var existing = totalRow - HainanStage2ExcelUtil.DataStartRow;
            if (count > existing)
            {
                worksheet.Row(totalRow).InsertRowsAbove(count - existing);
                for (var row = totalRow; row < totalRow + count - existing; row++)
                {
                    worksheet.Row(HainanStage2ExcelUtil.DataStartRow).CopyTo(worksheet.Row(row));
                }
            }
            else if (count < existing)
            {
                worksheet.Rows(HainanStage2ExcelUtil.DataStartRow + count, totalRow - 1).Delete();
            }

            return HainanStage2ExcelUtil.FindTotalRow(worksheet, HainanStage2ExcelUtil.DataStartRow);
        }
    }
}
