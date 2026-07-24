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
    internal sealed class GuangdongPowerCleanGenerator
    {
        private readonly GuangdongStage1SourceReader _sourceReader;

        public GuangdongPowerCleanGenerator(GuangdongStage1SourceReader sourceReader)
        {
            _sourceReader = sourceReader ?? throw new ArgumentNullException(nameof(sourceReader));
        }

        public ProvinceStage1CleanResult Generate(ProvinceStage1CleanOptions options)
        {
            var data = ReadData(options);
            var outputPath = WriteCleanWorkbook(options, data);
            var reportPath = UniquePath(Path.Combine(
                options.OutputDirectory,
                MonthPrefix(data.Month) + "月广东零售侧用户电量校验报告.json"));
            var htmlReportPath = UniquePath(Path.Combine(
                options.OutputDirectory,
                MonthPrefix(data.Month) + "月广东零售侧用户电量校验报告.html"));
            var result = new ProvinceStage1CleanResult
            {
                Province = ProvinceCode.Guangdong,
                Month = data.Month,
                Unit = GuangdongStage1SourceReader.Unit,
                RawDetailPath = options.RawDetailPath,
                OutputWorkbookPath = outputPath,
                ReportPath = reportPath,
                HtmlReportPath = htmlReportPath,
                SourceSheetName = data.SourceSheetName,
                RawRows = data.RawRows,
                CustomerRows = data.CustomerRows.Count,
                AccountRows = 0,
                TotalPower = Convert.ToDouble(data.CustomerRows.Sum(row => row.Total), CultureInfo.InvariantCulture),
                Warnings = data.Warnings.ToList()
            };
            new GuangdongStage1ReportWriter().WriteCleanReport(result, data.CustomerRows);
            return result;
        }

        internal GuangdongStage1SourceReader.GuangdongStage1DataSet ReadData(ProvinceStage1CleanOptions options)
        {
            return _sourceReader.Read(options.RawDetailPath, options.Month);
        }

        internal string WriteCleanWorkbook(
            ProvinceStage1CleanOptions options,
            GuangdongStage1SourceReader.GuangdongStage1DataSet data)
        {
            var outputPath = UniquePath(Path.Combine(
                options.OutputDirectory,
                string.IsNullOrWhiteSpace(options.OutputWorkbookName)
                    ? MonthPrefix(data.Month) + "月广东零售侧用户电量数据处理表.xlsx"
                    : options.OutputWorkbookName));
            EnsureOutputDoesNotOverwriteInput(outputPath, options.RawDetailPath);
            FileAccessGuard.RequireWritableWorkbook(outputPath, "广东电量处理表输出文件");
            WriteWorkbookContents(outputPath, data.CustomerRows);
            return outputPath;
        }

        private static void WriteWorkbookContents(
            string outputPath,
            IList<GuangdongStage1SourceReader.GuangdongPowerAggregateRow> rows)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("用户电量汇总");
                var headers = new[]
                {
                    "用户名称",
                    "用户编号",
                    "总实际用电量",
                    "峰电量",
                    "平电量",
                    "谷电量",
                    "峰_平",
                    "谷_平"
                };
                for (var column = 1; column <= headers.Length; column++)
                {
                    worksheet.Cell(1, column).Value = headers[column - 1];
                }

                for (var index = 0; index < rows.Count; index++)
                {
                    var row = rows[index];
                    var excelRow = index + 2;
                    worksheet.Cell(excelRow, 1).Value = row.CustomerName;
                    worksheet.Cell(excelRow, 2).Value = row.Code;
                    worksheet.Cell(excelRow, 2).Style.NumberFormat.Format = "@";
                    worksheet.Cell(excelRow, 3).SetValue(row.Total);
                    worksheet.Cell(excelRow, 4).SetValue(row.Peak);
                    worksheet.Cell(excelRow, 5).SetValue(row.Flat);
                    worksheet.Cell(excelRow, 6).SetValue(row.Valley);
                    if (row.HasCoefficientPair)
                    {
                        worksheet.Cell(excelRow, 7).SetValue(row.PeakFlatCoefficient);
                        worksheet.Cell(excelRow, 8).SetValue(row.ValleyFlatCoefficient);
                    }
                }

                var lastRow = Math.Max(1, rows.Count + 1);
                worksheet.Range(1, 1, 1, 8).Style.Font.Bold = true;
                worksheet.Range(1, 1, 1, 8).Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF4F4");
                worksheet.Range(1, 1, lastRow, 8).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range(2, 3, Math.Max(2, lastRow), 6).Style.NumberFormat.Format = "0.#####";
                worksheet.Range(2, 7, Math.Max(2, lastRow), 8).Style.NumberFormat.Format = "0.####";
                ApplyTableBorders(worksheet.Range(1, 1, lastRow, 8));
                worksheet.SheetView.FreezeRows(1);
                worksheet.Column(1).Width = 42;
                worksheet.Column(2).Width = 24;
                for (var column = 3; column <= 8; column++)
                {
                    worksheet.Column(column).Width = 18;
                }

                workbook.SaveAs(outputPath);
            }
        }

        private static void ApplyTableBorders(IXLRange range)
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

        private static void EnsureOutputDoesNotOverwriteInput(string outputPath, string inputPath)
        {
            if (string.Equals(
                Path.GetFullPath(outputPath),
                Path.GetFullPath(inputPath),
                StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("广东电量处理表输出路径不能与原始零售结算明细相同。");
            }
        }

        private static string MonthPrefix(int month)
        {
            return month > 0
                ? month.ToString(CultureInfo.InvariantCulture)
                : "未识别月份";
        }

        internal static string UniquePath(string path)
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
    }
}
