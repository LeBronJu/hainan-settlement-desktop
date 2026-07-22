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
    internal static class HainanStage2ExcelUtil
    {
        internal const int DataStartRow = 5;
        internal const int Year = 2026;
        internal const int TaxRatePrecision = 10;
        private static readonly Dictionary<string, Tuple<int, string>> PaymentPartyOverrides =
            new Dictionary<string, Tuple<int, string>>
            {
        { PaymentKey("海南精研科技有限公司", "代理费"), Tuple.Create(3, HainanStage2PaymentParties.Qingneng) }
            };

        internal static double GetNumeric(IXLWorksheet worksheet, int row, int column)
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

        internal static double NormalizeTaxRate(double value)
        {
            return Math.Round(value, TaxRatePrecision);
        }

        internal static bool TaxRatesEqual(double left, double right)
        {
            return NormalizeTaxRate(left) == NormalizeTaxRate(right);
        }

        internal static IXLWorksheet LastMonthSheet(XLWorkbook workbook)
        {
            var monthSheets = workbook.Worksheets
                .Where(sheet => Regex.IsMatch(sheet.Name, "^\\d+月$"))
                .OrderBy(sheet => Convert.ToInt32(sheet.Name.Replace("月", string.Empty)))
                .ToList();
            return monthSheets.Count > 0 ? monthSheets.Last() : workbook.Worksheets.Last();
        }

        internal static IXLWorksheet PreviousMonthSheet(XLWorkbook workbook, int month, string targetTitle)
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

        internal static bool TryParseMonthSheet(string name, out int month)
        {
            month = 0;
            var match = Regex.Match(TextUtil.S(name), "^(\\d{1,2})月$");
            return match.Success && int.TryParse(match.Groups[1].Value, out month);
        }

        internal static int FindTotalRow(IXLWorksheet worksheet, int startRow)
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

        internal static void SaveWorkbook(XLWorkbook workbook, string outputPath)
        {
            workbook.CalculateMode = XLCalculateMode.Auto;
            workbook.SaveAs(outputPath, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        internal static string TemplateKey(string kind, string owner, string entity)
        {
            return kind + "|" + NormalizeName(owner) + "|" + NormalizeName(entity);
        }

        internal static string TemplateSubjectKey(string kind, string entity)
        {
            return TextUtil.S(kind) + "|" + NormalizeName(entity);
        }

        internal static string SummaryKey(string entity, string kind)
        {
            return NormalizeName(entity) + "|" + TextUtil.S(kind);
        }

        internal static string PaymentKey(string entity, string kind)
        {
            return NormalizeName(entity) + "|" + TextUtil.S(kind);
        }

        internal static string NormalizeName(string value)
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

        internal static int ColumnNumber(string columnName)
        {
            var sum = 0;
            foreach (var c in columnName)
            {
                sum *= 26;
                sum += c - 'A' + 1;
            }

            return sum;
        }

        internal static string RelativePath(string root, string path)
        {
            var rootUri = new Uri(AppendDirectorySeparatorChar(Path.GetFullPath(root)));
            var pathUri = new Uri(Path.GetFullPath(path));
            return Uri.UnescapeDataString(rootUri.MakeRelativeUri(pathUri).ToString()).Replace('/', Path.DirectorySeparatorChar);
        }

        internal static string AppendDirectorySeparatorChar(string path)
        {
            return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ? path : path + Path.DirectorySeparatorChar;
        }

        internal static bool TryGetPaymentPartyOverride(string entity, string kind, int month, out string paymentParty)
        {
            paymentParty = null;
            Tuple<int, string> overrideValue;
            if (PaymentPartyOverrides.TryGetValue(PaymentKey(entity, kind), out overrideValue) && month >= overrideValue.Item1)
            {
                paymentParty = overrideValue.Item2;
                return true;
            }

            return false;
        }

        internal static bool TryGetPaymentPartyDecision(HainanStage2Options options, string entity, string kind, out string paymentParty)
        {
            paymentParty = null;
            if (options == null)
            {
                return false;
            }

            var key = SummaryKey(entity, kind);
            var decision = options.SummarySubjectDecisions
                .Where(item => item != null)
                .FirstOrDefault(item => SummaryKey(item.Entity, item.SettlementKind) == key);
            if (decision == null || string.IsNullOrWhiteSpace(decision.PaymentParty))
            {
                return false;
            }

            paymentParty = decision.PaymentParty;
            return true;
        }
    }
}
