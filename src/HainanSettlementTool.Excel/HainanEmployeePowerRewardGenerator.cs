using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal sealed class HainanEmployeePowerRewardGenerator
    {
        public IList<HainanEmployeePowerRewardLedgerRow> ReadLedgerRows(HainanEmployeePowerRewardOptions options)
        {
            using (var workbook = new XLWorkbook(options.LedgerPath))
            {
                var worksheet = FindLedgerWorksheet(workbook);
                var fixedColumns = FindFixedColumns(worksheet);
                var monthColumns = FindMonthColumns(worksheet, options.StartMonth, options.EndMonth);
                var rows = new List<HainanEmployeePowerRewardLedgerRow>();
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 2;

                for (var rowNumber = 4; rowNumber <= lastRow; rowNumber++)
                {
                    var row = ReadRow(worksheet, rowNumber, fixedColumns, monthColumns);
                    if (IsBlankIdentityRow(row))
                    {
                        continue;
                    }

                    rows.Add(row);
                }

                return rows;
            }
        }

        public HainanEmployeePowerRewardOutput GenerateWorkbooks(HainanEmployeePowerRewardOptions options, HainanEmployeePowerRewardResult result)
        {
            Directory.CreateDirectory(options.OutputDirectory);

            var summaryPath = GenerateSummaryWorkbook(options, result);
            var personalPaths = GeneratePersonalWorkbooks(options, result);
            var reportPath = WriteReport(options, result, summaryPath, personalPaths);

            return new HainanEmployeePowerRewardOutput
            {
                SummaryPath = summaryPath,
                ReportPath = reportPath,
                PersonalWorkbookPaths = personalPaths
            };
        }

        private static string GenerateSummaryWorkbook(HainanEmployeePowerRewardOptions options, HainanEmployeePowerRewardResult result)
        {
            var outputName = string.IsNullOrWhiteSpace(options.OutputSummaryName)
                ? PeriodLabel(options.Year, result.Months) + "员工电量奖励-海南.xlsx"
                : options.OutputSummaryName;
            var outputPath = UniquePath(Path.Combine(options.OutputDirectory, outputName));

            using (var workbook = new XLWorkbook())
            {
                var detailSheet = workbook.AddWorksheet(SheetPrefix(result.Months) + "月企业用电量明细");
                detailSheet.Name = SheetPrefix(result.Months) + "月企业用电量明细";
                WriteDetailSheet(
                    detailSheet,
                    PeriodLabel(options.Year, result.Months) + "企业用电量明细",
                    result.Months,
                    result.Details,
                    personalFooter: false);

                var summarySheet = workbook.AddWorksheet(SheetPeriodLabel(result.Months) + "员工电量汇总");
                summarySheet.Name = SheetPeriodLabel(result.Months) + "员工电量汇总";
                WriteResponsiblePersonSummarySheet(
                    summarySheet,
                    PeriodLabel(options.Year, result.Months) + "员工电量奖励",
                    result.Months,
                    result.ResponsiblePersonSummaries);

                SaveWorkbook(workbook, outputPath);
            }

            return outputPath;
        }

        private static List<string> GeneratePersonalWorkbooks(HainanEmployeePowerRewardOptions options, HainanEmployeePowerRewardResult result)
        {
            var paths = new List<string>();
            foreach (var summary in result.ResponsiblePersonSummaries.OrderBy(row => row.ResponsiblePerson, StringComparer.Ordinal))
            {
                var details = result.Details
                    .Where(row => row.ResponsiblePerson == summary.ResponsiblePerson)
                    .OrderBy(row => row.SourceRow)
                    .ToList();
                var fileName = TextUtil.SafeFileName(summary.ResponsiblePerson)
                    + "-"
                    + PeriodLabel(options.Year, result.Months)
                    + "员工电量确认表-海南.xlsx";
                var outputPath = UniquePath(Path.Combine(options.OutputDirectory, fileName));

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet(SheetPrefix(result.Months) + "月企业用电明细");
                    worksheet.Name = SheetPrefix(result.Months) + "月企业用电明细";
                    WriteDetailSheet(
                        worksheet,
                        PeriodLabel(options.Year, result.Months) + "员工电量确认表(" + summary.ResponsiblePerson + ")",
                        result.Months,
                        details,
                        personalFooter: true);
                    SaveWorkbook(workbook, outputPath);
                }

                paths.Add(outputPath);
            }

            return paths;
        }

        private static string WriteReport(
            HainanEmployeePowerRewardOptions options,
            HainanEmployeePowerRewardResult result,
            string summaryPath,
            IList<string> personalPaths)
        {
            var reportPath = UniquePath(Path.Combine(options.OutputDirectory, PeriodLabel(options.Year, result.Months) + "员工电量奖励校验报告.json"));
            var payload = new
            {
                year = result.Year,
                months = result.Months,
                totalCustomers = result.TotalCustomers,
                employeeCount = result.ResponsiblePersonSummaries.Count,
                monthTotals = result.MonthTotals,
                totalPower = result.TotalPower,
                totalReward = result.TotalReward,
                summaryPath,
                personalWorkbookCount = personalPaths.Count,
                personalWorkbookPaths = personalPaths
            };
            File.WriteAllText(reportPath, JsonConvert.SerializeObject(payload, Formatting.Indented));
            return reportPath;
        }

        private static void WriteDetailSheet(
            IXLWorksheet worksheet,
            string title,
            IList<int> months,
            IList<HainanEmployeePowerRewardDetail> details,
            bool personalFooter)
        {
            var totalColumn = 8 + months.Count;
            var remarkColumn = totalColumn + 1;
            var lastColumn = remarkColumn;
            SetTitle(worksheet, title, lastColumn);
            WriteDetailHeaders(worksheet, months, totalColumn, remarkColumn);
            ApplyColumnWidths(worksheet, lastColumn);

            for (var index = 0; index < details.Count; index++)
            {
                var rowNumber = 3 + index;
                var detail = details[index];
                worksheet.Cell(rowNumber, 1).Value = index + 1;
                worksheet.Cell(rowNumber, 2).Value = detail.CustomerCode;
                worksheet.Cell(rowNumber, 3).Value = detail.CustomerName;
                worksheet.Cell(rowNumber, 4).Value = detail.ContractStartMonth;
                worksheet.Cell(rowNumber, 5).Value = detail.ProjectDeveloper;
                worksheet.Cell(rowNumber, 6).Value = detail.AgentType;
                worksheet.Cell(rowNumber, 7).Value = detail.ResponsiblePerson;

                for (var monthIndex = 0; monthIndex < months.Count; monthIndex++)
                {
                    var column = 8 + monthIndex;
                    worksheet.Cell(rowNumber, column).Value = detail.MonthlyPowers[months[monthIndex]];
                }

                worksheet.Cell(rowNumber, totalColumn).FormulaA1 = "SUM("
                    + ClosedXmlUtil.ColumnLetter(8)
                    + rowNumber
                    + ":"
                    + ClosedXmlUtil.ColumnLetter(totalColumn - 1)
                    + rowNumber
                    + ")";
                worksheet.Cell(rowNumber, remarkColumn).Value = string.Empty;
            }

            var totalRow = 3 + details.Count;
            worksheet.Cell(totalRow, 1).Value = "合计";
            for (var monthIndex = 0; monthIndex < months.Count; monthIndex++)
            {
                var column = 8 + monthIndex;
                var letter = ClosedXmlUtil.ColumnLetter(column);
                worksheet.Cell(totalRow, column).FormulaA1 = "SUM(" + letter + "3:" + letter + (totalRow - 1) + ")";
            }

            var totalLetter = ClosedXmlUtil.ColumnLetter(totalColumn);
            worksheet.Cell(totalRow, totalColumn).FormulaA1 = "SUM(" + totalLetter + "3:" + totalLetter + (totalRow - 1) + ")";
            worksheet.Cell(totalRow, remarkColumn).Value = string.Empty;
            ApplyDetailStyles(worksheet, totalRow, lastColumn, totalColumn, remarkColumn);

            var footerColumn = personalFooter ? Math.Max(1, totalColumn - 1) : totalColumn;
            worksheet.Cell(totalRow + 1, footerColumn).Value = "确认人：";
            worksheet.Cell(totalRow + 2, footerColumn).Value = "日  期：";
        }

        private static void WriteResponsiblePersonSummarySheet(
            IXLWorksheet worksheet,
            string title,
            IList<int> months,
            IList<HainanEmployeePowerRewardSummary> summaries)
        {
            var totalColumn = 3 + months.Count;
            var rewardColumn = totalColumn + 1;
            var remarkColumn = rewardColumn + 1;
            var lastColumn = remarkColumn;
            SetTitle(worksheet, title, lastColumn);
            WriteSummaryHeaders(worksheet, months, totalColumn, rewardColumn, remarkColumn);
            ApplySummaryColumnWidths(worksheet, lastColumn);

            for (var index = 0; index < summaries.Count; index++)
            {
                var rowNumber = 3 + index;
                var summary = summaries[index];
                worksheet.Cell(rowNumber, 1).Value = index + 1;
                worksheet.Cell(rowNumber, 2).Value = summary.ResponsiblePerson;
                for (var monthIndex = 0; monthIndex < months.Count; monthIndex++)
                {
                    worksheet.Cell(rowNumber, 3 + monthIndex).Value = summary.MonthlyPowers[months[monthIndex]];
                }

                worksheet.Cell(rowNumber, totalColumn).FormulaA1 = "SUM("
                    + ClosedXmlUtil.ColumnLetter(3)
                    + rowNumber
                    + ":"
                    + ClosedXmlUtil.ColumnLetter(totalColumn - 1)
                    + rowNumber
                    + ")";
                worksheet.Cell(rowNumber, rewardColumn).FormulaA1 = ClosedXmlUtil.ColumnLetter(totalColumn) + rowNumber + "*10000*0.0001";
                worksheet.Cell(rowNumber, remarkColumn).Value = string.Empty;
            }

            var totalRow = 3 + summaries.Count;
            worksheet.Range(totalRow, 1, totalRow, 2).Merge();
            worksheet.Cell(totalRow, 1).Value = "合计";
            for (var monthIndex = 0; monthIndex < months.Count; monthIndex++)
            {
                var column = 3 + monthIndex;
                worksheet.Cell(totalRow, column).FormulaA1 = SumFormula(column, 3, totalRow - 1);
            }

            worksheet.Cell(totalRow, totalColumn).FormulaA1 = SumFormula(totalColumn, 3, totalRow - 1);
            worksheet.Cell(totalRow, rewardColumn).FormulaA1 = SumFormula(rewardColumn, 3, totalRow - 1);
            worksheet.Cell(totalRow, remarkColumn).Value = string.Empty;

            var remarkRow = totalRow + 1;
            worksheet.Range(remarkRow, 1, remarkRow, lastColumn).Merge();
            worksheet.Cell(remarkRow, 1).Value = "备注：按0.0001元/千瓦时奖励。";
            ApplySummaryStyles(worksheet, summaries.Count, totalRow, remarkRow, lastColumn, totalColumn, rewardColumn);

            var footerRow = remarkRow;
            worksheet.Cell(footerRow + 1, 1).Value = "审批：";
            worksheet.Cell(footerRow + 1, 3).Value = "审核：";
            worksheet.Cell(footerRow + 1, Math.Min(lastColumn, 4)).Value = "财务中心复核：";
            worksheet.Cell(footerRow + 1, Math.Min(lastColumn, 6)).Value = "交易结算部复核：";
            worksheet.Cell(footerRow + 1, Math.Max(1, lastColumn - 1)).Value = "制表：";
            worksheet.Range(footerRow + 2, Math.Max(1, lastColumn - 2), footerRow + 2, lastColumn - 1).Merge();
            worksheet.Cell(footerRow + 2, Math.Max(1, lastColumn - 2)).Value = "日期：";
        }

        private static void WriteDetailHeaders(IXLWorksheet worksheet, IList<int> months, int totalColumn, int remarkColumn)
        {
            worksheet.Cell(2, 1).Value = "序号";
            worksheet.Cell(2, 2).Value = "用电企业编号";
            worksheet.Cell(2, 3).Value = "用电企业名称";
            worksheet.Cell(2, 4).Value = "履约开始月份";
            worksheet.Cell(2, 5).Value = "项目开发人";
            worksheet.Cell(2, 6).Value = "代理或自营";
            worksheet.Cell(2, 7).Value = "负责人";
            for (var index = 0; index < months.Count; index++)
            {
                worksheet.Cell(2, 8 + index).Value = months[index] + "月总实际电量（万千瓦时）";
            }

            worksheet.Cell(2, totalColumn).Value = "电量合计（万千瓦时）";
            worksheet.Cell(2, remarkColumn).Value = "备注";
        }

        private static void WriteSummaryHeaders(
            IXLWorksheet worksheet,
            IList<int> months,
            int totalColumn,
            int rewardColumn,
            int remarkColumn)
        {
            worksheet.Cell(2, 1).Value = "序号";
            worksheet.Cell(2, 2).Value = "负责人";
            for (var index = 0; index < months.Count; index++)
            {
                worksheet.Cell(2, 3 + index).Value = months[index] + "月总实际电量（万千瓦时）";
            }

            worksheet.Cell(2, totalColumn).Value = "电量合计（万千瓦时）";
            worksheet.Cell(2, rewardColumn).Value = "电量奖励（元）";
            worksheet.Cell(2, remarkColumn).Value = "备注";
        }

        private static void ApplyColumnWidths(IXLWorksheet worksheet, int lastColumn)
        {
            worksheet.Column(1).Width = 7;
            worksheet.Column(2).Width = 20;
            worksheet.Column(3).Width = 34;
            worksheet.Column(4).Width = 14;
            worksheet.Column(5).Width = 30;
            worksheet.Column(6).Width = 12;
            worksheet.Column(7).Width = 12;
            for (var column = 8; column < lastColumn; column++)
            {
                worksheet.Column(column).Width = 18;
            }

            worksheet.Column(lastColumn).Width = 16;
        }

        private static void ApplySummaryColumnWidths(IXLWorksheet worksheet, int lastColumn)
        {
            worksheet.Column(1).Width = 7;
            worksheet.Column(2).Width = 14;
            for (var column = 3; column <= lastColumn; column++)
            {
                worksheet.Column(column).Width = column == lastColumn ? 18 : 20;
            }
        }

        private static void ApplyDetailStyles(
            IXLWorksheet worksheet,
            int totalRow,
            int lastColumn,
            int totalColumn,
            int remarkColumn)
        {
            ApplyTitleAndHeaderStyles(worksheet, lastColumn);
            var table = worksheet.Range(1, 1, totalRow, lastColumn);
            ApplyTableBorder(table);
            table.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            table.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range(3, 3, totalRow - 1, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Range(3, 5, totalRow - 1, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Range(3, 8, totalRow, totalColumn).Style.NumberFormat.Format = "0.0000";
            worksheet.Range(totalRow, 1, totalRow, lastColumn).Style.Font.Bold = true;
            worksheet.Range(totalRow, 1, totalRow, lastColumn).Style.Fill.BackgroundColor = XLColor.FromHtml("#F3F7F8");
            worksheet.Column(remarkColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }

        private static void ApplySummaryStyles(
            IXLWorksheet worksheet,
            int summaryRows,
            int totalRow,
            int remarkRow,
            int lastColumn,
            int totalColumn,
            int rewardColumn)
        {
            ApplyTitleAndHeaderStyles(worksheet, lastColumn);

            var table = worksheet.Range(1, 1, remarkRow, lastColumn);
            ApplyTableBorder(table);
            table.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            table.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range(3, 3, totalRow, totalColumn).Style.NumberFormat.Format = "0.0000";
            worksheet.Range(3, rewardColumn, totalRow, rewardColumn).Style.NumberFormat.Format = "0.00";
            worksheet.Range(totalRow, 1, totalRow, lastColumn).Style.Font.Bold = true;
            worksheet.Range(totalRow, 1, totalRow, lastColumn).Style.Fill.BackgroundColor = XLColor.FromHtml("#F3F7F8");
            worksheet.Cell(remarkRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }

        private static void ApplyTitleAndHeaderStyles(IXLWorksheet worksheet, int lastColumn)
        {
            worksheet.Row(1).Height = 28;
            worksheet.Row(2).Height = 36;
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 1).Style.Font.FontSize = 16;
            worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            var header = worksheet.Range(2, 1, 2, lastColumn);
            header.Style.Font.Bold = true;
            header.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF4F4");
            header.Style.Alignment.WrapText = true;
            header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static void ApplyTableBorder(IXLRange range)
        {
            foreach (var cell in range.Cells())
            {
                cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.TopBorderColor = XLColor.Black;
                cell.Style.Border.BottomBorderColor = XLColor.Black;
                cell.Style.Border.LeftBorderColor = XLColor.Black;
                cell.Style.Border.RightBorderColor = XLColor.Black;
            }
        }

        private static void SetTitle(IXLWorksheet worksheet, string title, int lastColumn)
        {
            UnmergeIntersecting(worksheet, 1, 1, 1, Math.Max(lastColumn, worksheet.LastColumnUsed()?.ColumnNumber() ?? lastColumn));
            worksheet.Range(1, 1, 1, lastColumn).Merge();
            worksheet.Cell(1, 1).Value = title;
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

        private static string PeriodLabel(int year, IList<int> months)
        {
            if (months.Count == 1)
            {
                return year + "年" + months[0] + "月";
            }

            return year + "年" + months.First() + "-" + months.Last() + "月";
        }

        private static string SheetPrefix(IList<int> months)
        {
            if (months.Count == 1)
            {
                return months[0].ToString(CultureInfo.InvariantCulture);
            }

            return months.First() + "-" + months.Last();
        }

        private static string SheetPeriodLabel(IList<int> months)
        {
            if (months.Count == 1)
            {
                return months[0].ToString(CultureInfo.InvariantCulture) + "月";
            }

            return months.First().ToString(CultureInfo.InvariantCulture)
                + "月-"
                + months.Last().ToString(CultureInfo.InvariantCulture)
                + "月";
        }

        private static string SumFormula(int column, int firstRow, int lastRow)
        {
            if (lastRow < firstRow)
            {
                return "0";
            }

            var letter = ClosedXmlUtil.ColumnLetter(column);
            return "SUM(" + letter + firstRow + ":" + letter + lastRow + ")";
        }

        private static void SaveWorkbook(XLWorkbook workbook, string outputPath)
        {
            workbook.CalculateMode = XLCalculateMode.Auto;
            workbook.SaveAs(outputPath, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        private static HainanEmployeePowerRewardLedgerRow ReadRow(
            IXLWorksheet worksheet,
            int rowNumber,
            FixedColumns fixedColumns,
            IDictionary<int, int> monthColumns)
        {
            var monthlyPowers = new Dictionary<int, double>();
            foreach (var item in monthColumns)
            {
                monthlyPowers[item.Key] = ReadPower(worksheet.Cell(rowNumber, item.Value), rowNumber, item.Key);
            }

            return new HainanEmployeePowerRewardLedgerRow
            {
                SourceRow = rowNumber,
                Sequence = ReadSequence(worksheet.Cell(rowNumber, 1), rowNumber),
                CustomerCode = CellText(worksheet.Cell(rowNumber, fixedColumns.CustomerCode)),
                CustomerName = CellText(worksheet.Cell(rowNumber, fixedColumns.CustomerName)),
                ContractStartMonth = CellText(worksheet.Cell(rowNumber, fixedColumns.ContractStartMonth)),
                ProjectDeveloper = CellText(worksheet.Cell(rowNumber, fixedColumns.ProjectDeveloper)),
                AgentType = CellText(worksheet.Cell(rowNumber, fixedColumns.AgentType)),
                ResponsiblePerson = CellText(worksheet.Cell(rowNumber, fixedColumns.ResponsiblePerson)),
                MonthlyPowers = monthlyPowers
            };
        }

        private static IXLWorksheet FindLedgerWorksheet(XLWorkbook workbook)
        {
            var named = workbook.Worksheets.FirstOrDefault(ws => ws.Name == LedgerLayout.MainSheetName);
            if (named != null)
            {
                return named;
            }

            var matched = workbook.Worksheets.FirstOrDefault(HasRequiredLedgerHeaders);
            if (matched != null)
            {
                return matched;
            }

            throw new InvalidOperationException("找不到员工电量奖励可用的售电结算台账 sheet。");
        }

        private static bool HasRequiredLedgerHeaders(IXLWorksheet worksheet)
        {
            return FindHeaderColumnOrZero(worksheet, "用电企业编号") > 0
                && FindHeaderColumnOrZero(worksheet, "用电企业名称") > 0
                && FindHeaderColumnOrZero(worksheet, "负责人") > 0;
        }

        private static FixedColumns FindFixedColumns(IXLWorksheet worksheet)
        {
            return new FixedColumns
            {
                CustomerCode = FindHeaderColumn(worksheet, "用电企业编号"),
                CustomerName = FindHeaderColumn(worksheet, "用电企业名称"),
                ContractStartMonth = FindHeaderColumn(worksheet, "履约开始月份"),
                ProjectDeveloper = FindHeaderColumn(worksheet, "项目开发人"),
                AgentType = FindHeaderColumn(worksheet, "代理或自营"),
                ResponsiblePerson = FindHeaderColumn(worksheet, "负责人")
            };
        }

        private static int FindHeaderColumn(IXLWorksheet worksheet, string header)
        {
            var column = FindHeaderColumnOrZero(worksheet, header);
            if (column <= 0)
            {
                throw new InvalidOperationException("台账缺少表头：" + header);
            }

            return column;
        }

        private static int FindHeaderColumnOrZero(IXLWorksheet worksheet, string header)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (CellText(worksheet.Cell(2, column)) == header)
                {
                    return column;
                }
            }

            return 0;
        }

        private static Dictionary<int, int> FindMonthColumns(IXLWorksheet worksheet, int startMonth, int endMonth)
        {
            var result = new Dictionary<int, int>();
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var month = startMonth; month <= endMonth; month++)
            {
                var monthText = month + "月";
                var column = 0;
                for (var candidate = 1; candidate <= lastColumn; candidate++)
                {
                    var header1 = CellText(worksheet.Cell(1, candidate));
                    var header2 = CellText(worksheet.Cell(2, candidate));
                    if (header1 == monthText && header2.Contains("总实际电量"))
                    {
                        column = candidate;
                        break;
                    }
                }

                if (column <= 0)
                {
                    throw new InvalidOperationException("未找到" + monthText + "总实际电量列。");
                }

                result[month] = column;
            }

            return result;
        }

        private static int ReadSequence(IXLCell cell, int fallback)
        {
            var value = TextUtil.N(cell.GetFormattedString());
            return value > 0 ? Convert.ToInt32(value) : fallback;
        }

        private static double ReadPower(IXLCell cell, int rowNumber, int month)
        {
            if (cell.IsEmpty())
            {
                return 0d;
            }

            if (cell.DataType == XLDataType.Number)
            {
                return cell.GetDouble();
            }

            var text = CellText(cell).Replace(",", string.Empty);
            if (text.Length == 0)
            {
                return 0d;
            }

            double parsed;
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
            {
                return parsed;
            }

            throw new InvalidOperationException("台账第" + rowNumber + "行" + month + "月总实际电量不是数字：" + text);
        }

        private static string CellText(IXLCell cell)
        {
            return TextUtil.S(cell.GetFormattedString());
        }

        private static bool IsBlankIdentityRow(HainanEmployeePowerRewardLedgerRow row)
        {
            return string.IsNullOrWhiteSpace(row.CustomerCode)
                && string.IsNullOrWhiteSpace(row.CustomerName)
                && string.IsNullOrWhiteSpace(row.ProjectDeveloper)
                && string.IsNullOrWhiteSpace(row.AgentType)
                && string.IsNullOrWhiteSpace(row.ResponsiblePerson);
        }

        private sealed class FixedColumns
        {
            public int CustomerCode { get; set; }
            public int CustomerName { get; set; }
            public int ContractStartMonth { get; set; }
            public int ProjectDeveloper { get; set; }
            public int AgentType { get; set; }
            public int ResponsiblePerson { get; set; }
        }
    }
}
