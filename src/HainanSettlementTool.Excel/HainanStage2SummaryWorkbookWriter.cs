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
    internal static class HainanStage2SummaryWorkbookWriter
    {
        internal static string BuildSummary(HainanStage2Options options, IList<GroupSettlementTotal> totals, IList<string> warnings)
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
                var partyByKey = BuildPaymentPartyIndex(options, totals, mainMeta);

                WriteSummarySheet(workbook.Worksheet(mainSheetName), totals, options.Month, null, warnings, partyByKey);

                if (!string.IsNullOrWhiteSpace(qingnengSheetName))
                {
                    var qnTotals = totals.Where(total => PartyForSummaryTotal(total, partyByKey) == HainanStage2PaymentParties.Qingneng).ToList();
                    var allowed = new HashSet<string>(partyByKey.Where(item => item.Value == HainanStage2PaymentParties.Qingneng).Select(item => item.Key));
                    foreach (var total in qnTotals)
                    {
                        allowed.Add(HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind));
                    }
                    WriteSummarySheet(workbook.Worksheet(qingnengSheetName), qnTotals, options.Month, allowed, warnings, partyByKey);
                }

                if (!string.IsNullOrWhiteSpace(qinghuiSheetName))
                {
                    var qhTotals = totals.Where(total => PartyForSummaryTotal(total, partyByKey) == HainanStage2PaymentParties.Qinghui).ToList();
                    var allowed = new HashSet<string>(partyByKey.Where(item => item.Value == HainanStage2PaymentParties.Qinghui).Select(item => item.Key));
                    foreach (var total in qhTotals)
                    {
                        allowed.Add(HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind));
                    }
                    WriteSummarySheet(workbook.Worksheet(qinghuiSheetName), qhTotals, options.Month, allowed, warnings, partyByKey);
                }

                HainanStage2ExcelUtil.SaveWorkbook(workbook, outputPath);
            }

            return outputPath;
        }

        private static void WriteSummarySheet(
            IXLWorksheet worksheet,
            IList<GroupSettlementTotal> totals,
            int month,
            ISet<string> allowedKeys,
            IList<string> warnings,
            IDictionary<string, string> paymentPartyByKey)
        {
            var startRow = 4;
            DeleteSummaryRowsNotAllowed(worksheet, startRow, allowedKeys);

            var monthColumn = InsertMonthBlock(worksheet, month);
            var cumulativeColumn = monthColumn + 6;
            var totalByKey = totals.ToDictionary(total => HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind), total => total);
            var existingMeta = ReadSummaryMeta(worksheet);
            var knownKeys = new HashSet<string>(existingMeta.Select(item => HainanStage2ExcelUtil.SummaryKey(item.Entity, item.Kind)));
            var newTotals = totals.Where(total => !knownKeys.Contains(HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind))).ToList();
            var newKeys = new HashSet<string>(newTotals.Select(total => HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind)));
            InsertNewSummaryRows(worksheet, startRow, newTotals, warnings, paymentPartyByKey);

            var dataRows = ReadSummaryMeta(worksheet).OrderBy(item => item.Row).ToList();
            for (var index = 0; index < dataRows.Count; index++)
            {
                var row = dataRows[index].Row;
                var info = dataRows[index];
                GroupSettlementTotal total;
                totalByKey.TryGetValue(HainanStage2ExcelUtil.SummaryKey(info.Entity, info.Kind), out total);
                worksheet.Cell(row, 1).Value = index + 1;
                WriteSummaryValues(worksheet, row, monthColumn, cumulativeColumn, total, info.Entity, info.Kind, month, newKeys.Contains(HainanStage2ExcelUtil.SummaryKey(info.Entity, info.Kind)), paymentPartyByKey);
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
            bool isNewRow,
            IDictionary<string, string> paymentPartyByKey)
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
                worksheet.Cell(row, cumulativeColumn + 7).Value = new DateTime(HainanStage2ExcelUtil.Year, month, 1);
            }

            var currentParty = TextUtil.S(worksheet.Cell(row, cumulativeColumn + 8).GetFormattedString());
            worksheet.Cell(row, cumulativeColumn + 8).Value = isNewRow
                ? PaymentPartyFromIndex(paymentPartyByKey, entity, kind)
                : PaymentPartyFor(entity, kind, month, string.IsNullOrWhiteSpace(currentParty) ? HainanStage2PaymentParties.Qinghui : currentParty);
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

            worksheet.Cell(2, insertAt).Value = HainanStage2ExcelUtil.Year + "年" + month + "月";
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
                if (string.IsNullOrWhiteSpace(entity) || !allowedKeys.Contains(HainanStage2ExcelUtil.SummaryKey(entity, kind)))
                {
                    worksheet.Row(row).Delete();
                }
            }
        }

        private static void InsertNewSummaryRows(
            IXLWorksheet worksheet,
            int startRow,
            IList<GroupSettlementTotal> newTotals,
            IList<string> warnings,
            IDictionary<string, string> paymentPartyByKey)
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
                warnings.Add("新增汇总主体：" + total.Kind + " " + total.Entity + "（负责人：" + total.Owner + "；支付方：" + PaymentPartyFromIndex(paymentPartyByKey, total.Entity, total.Kind) + "）");
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
            var date = new DateTime(HainanStage2ExcelUtil.Year, month, 8).AddMonths(2);
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
            worksheet.Cell(2, monthColumn + 5).Value = "当月实际支付";
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

        internal static IList<HainanStage2SummaryMetaRow> ReadSummaryMeta(IXLWorksheet worksheet)
        {
            var cumulativeColumn = SummaryColumn(worksheet, "累计代理费总计");
            var totalRow = FindSummaryTotalRow(worksheet, 4);
            var rows = new List<HainanStage2SummaryMetaRow>();
            for (var row = 4; row < totalRow; row++)
            {
                var entity = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                var kind = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                if (string.IsNullOrWhiteSpace(entity) || entity.Contains("审核"))
                {
                    continue;
                }

                rows.Add(new HainanStage2SummaryMetaRow
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

        internal static string ResolveSummarySheetName(XLWorkbook workbook, string role, bool required)
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

        private static Dictionary<string, string> BuildPaymentPartyIndex(
            HainanStage2Options options,
            IList<GroupSettlementTotal> totals,
            IList<HainanStage2SummaryMetaRow> mainMeta)
        {
            var partyByKey = mainMeta.ToDictionary(
                item => HainanStage2ExcelUtil.SummaryKey(item.Entity, item.Kind),
                item => PaymentPartyFor(item.Entity, item.Kind, options.Month, string.IsNullOrWhiteSpace(item.PaymentParty) ? HainanStage2PaymentParties.Qinghui : item.PaymentParty));

            foreach (var total in totals)
            {
                var key = HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind);
                if (partyByKey.ContainsKey(key))
                {
                    continue;
                }

                string paymentParty;
                if (!HainanStage2ExcelUtil.TryGetPaymentPartyOverride(total.Entity, total.Kind, options.Month, out paymentParty)
                    && !HainanStage2ExcelUtil.TryGetPaymentPartyDecision(options, total.Entity, total.Kind, out paymentParty))
                {
                    throw new InvalidOperationException("海南阶段二新增汇总主体支付方未选择：" + total.Kind + " " + total.Entity + "。请在预检窗口选择清能或清辉后再生成。");
                }

                partyByKey[key] = PaymentPartyFor(total.Entity, total.Kind, options.Month, paymentParty);
            }

            return partyByKey;
        }

        private static string PartyForSummaryTotal(GroupSettlementTotal total, IDictionary<string, string> partyByKey)
        {
            string party;
            if (partyByKey.TryGetValue(HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind), out party))
            {
                return party;
            }

            throw new InvalidOperationException("海南阶段二新增汇总主体支付方未选择：" + total.Kind + " " + total.Entity + "。请在预检窗口选择清能或清辉后再生成。");
        }

        private static string PaymentPartyFor(string entity, string kind, int month, string defaultParty)
        {
            string overrideParty;
            if (HainanStage2ExcelUtil.TryGetPaymentPartyOverride(entity, kind, month, out overrideParty))
            {
                return overrideParty;
            }

            return string.IsNullOrWhiteSpace(defaultParty) ? HainanStage2PaymentParties.Qinghui : defaultParty;
        }

        private static string PaymentPartyFromIndex(IDictionary<string, string> paymentPartyByKey, string entity, string kind)
        {
            string paymentParty;
            if (paymentPartyByKey != null && paymentPartyByKey.TryGetValue(HainanStage2ExcelUtil.SummaryKey(entity, kind), out paymentParty))
            {
                return paymentParty;
            }

            throw new InvalidOperationException("海南阶段二新增汇总主体支付方未选择：" + kind + " " + entity + "。请在预检窗口选择清能或清辉后再生成。");
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
    }
}
