using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelDataReader;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal sealed class ChongqingPowerCleanGenerator
    {
        private const string Unit = "兆瓦时";
        private static readonly string[] RequiredHeaders = { "用户名称", "户号", "时段", "用电量" };

        public ProvinceStage1CleanResult Generate(ProvinceStage1CleanOptions options)
        {
            var data = ReadData(options);
            var outputPath = UniquePath(Path.Combine(options.OutputDirectory, OutputWorkbookName(options, data.Month)));
            var reportPath = UniquePath(Path.Combine(options.OutputDirectory, ReportName(data.Month)));

            WriteWorkbook(outputPath, data.Month, data.CustomerRows, data.AccountRows);
            var result = new ProvinceStage1CleanResult
            {
                Province = ProvinceCode.Chongqing,
                Month = data.Month,
                Unit = Unit,
                RawDetailPath = options.RawDetailPath,
                OutputWorkbookPath = outputPath,
                ReportPath = reportPath,
                SourceSheetName = data.SourceSheetName,
                RawRows = data.RawRows,
                CustomerRows = data.CustomerRows.Count,
                AccountRows = data.AccountRows.Count,
                TotalPower = Math.Round(data.CustomerRows.Sum(row => row.Total), 4),
                Warnings = data.Warnings
            };
            WriteReport(reportPath, result);
            return result;
        }

        internal ChongqingPowerDataSet ReadData(ProvinceStage1CleanOptions options)
        {
            var source = ReadSource(options.RawDetailPath);
            var month = ResolveMonth(options, source.TitleText);
            return new ChongqingPowerDataSet
            {
                Month = month,
                Unit = Unit,
                SourceSheetName = source.SheetName,
                RawRows = source.Rows.Count,
                CustomerRows = AggregateCustomerRows(source.Rows),
                AccountRows = AggregateAccountRows(source.Rows),
                Warnings = source.Warnings
            };
        }

        private static SourceReadResult ReadSource(string path)
        {
            var extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension == ".xlsx")
            {
                return ReadXlsx(path);
            }

            if (extension == ".xls")
            {
                return ReadXls(path);
            }

            if (extension == ".csv")
            {
                return ReadCsv(path);
            }

            throw new NotSupportedException("重庆电量确认结算单只支持 .xlsx、.xls 或 .csv。");
        }

        private static SourceReadResult ReadXlsx(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, "sheet1", StringComparison.OrdinalIgnoreCase))
                    ?? workbook.Worksheets.First();
                var header = FindXlsxHeader(worksheet);
                var titleText = CellText(worksheet.Cell(1, 1));
                var rows = new List<ChongqingPowerRawRow>();
                var warnings = new List<string>();
                var invalidRows = new List<string>();
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? header.Row;

                for (var rowNumber = header.Row + 1; rowNumber <= lastRow; rowNumber++)
                {
                    var raw = ReadXlsxRow(worksheet, rowNumber, header.Columns);
                    AddValidRow(raw, rows, warnings, invalidRows);
                }

                ThrowIfInvalid(invalidRows);
                return new SourceReadResult
                {
                    SheetName = worksheet.Name,
                    TitleText = titleText,
                    Rows = rows,
                    Warnings = warnings
                };
            }
        }

        private static SourceReadResult ReadXls(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var selected = ReadSelectedXlsSheet(reader);
                ThrowIfInvalid(selected.InvalidRows);
                return new SourceReadResult
                {
                    SheetName = selected.SheetName,
                    TitleText = selected.TitleText,
                    Rows = selected.Rows,
                    Warnings = selected.Warnings
                };
            }
        }

        private static SourceReadResult ReadCsv(string path)
        {
            var lines = File.ReadAllLines(path, DetectCsvEncoding(path));
            var parsed = lines.Select(SplitCsvLine).ToList();
            var header = FindCsvHeader(parsed);
            var rows = new List<ChongqingPowerRawRow>();
            var warnings = new List<string>();
            var invalidRows = new List<string>();

            for (var index = header.Row; index < parsed.Count; index++)
            {
                var raw = ReadCsvRow(parsed[index], index + 1, header.Columns);
                AddValidRow(raw, rows, warnings, invalidRows);
            }

            ThrowIfInvalid(invalidRows);
            return new SourceReadResult
            {
                SheetName = "CSV",
                TitleText = parsed.Count > 0 && parsed[0].Count > 0 ? parsed[0][0] : string.Empty,
                Rows = rows,
                Warnings = warnings
            };
        }

        private static XlsSheetReadResult ReadSelectedXlsSheet(IExcelDataReader reader)
        {
            XlsSheetReadResult firstSheet = null;
            do
            {
                var result = ReadCurrentXlsSheet(reader);
                if (firstSheet == null)
                {
                    firstSheet = result;
                }

                if (string.Equals(reader.Name, "sheet1", StringComparison.OrdinalIgnoreCase))
                {
                    return result;
                }
            }
            while (reader.NextResult());

            return firstSheet;
        }

        private static XlsSheetReadResult ReadCurrentXlsSheet(IExcelDataReader reader)
        {
            var rawRows = new List<List<object>>();
            while (reader.Read())
            {
                var values = new List<object>();
                for (var index = 0; index < reader.FieldCount; index++)
                {
                    values.Add(reader.GetValue(index));
                }

                rawRows.Add(values);
            }

            var header = FindXlsHeader(rawRows);
            var rows = new List<ChongqingPowerRawRow>();
            var warnings = new List<string>();
            var invalidRows = new List<string>();
            for (var index = header.Row; index < rawRows.Count; index++)
            {
                var raw = ReadXlsRow(rawRows[index], index + 1, header.Columns);
                AddValidRow(raw, rows, warnings, invalidRows);
            }

            return new XlsSheetReadResult
            {
                SheetName = reader.Name,
                TitleText = rawRows.Count > 0 && rawRows[0].Count > 0 ? TextUtil.S(rawRows[0][0]) : string.Empty,
                Rows = rows,
                Warnings = warnings,
                InvalidRows = invalidRows
            };
        }

        private static HeaderMap FindXlsxHeader(IXLWorksheet worksheet)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            var lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 1, 30);
            for (var row = 1; row <= lastRow; row++)
            {
                var columns = new Dictionary<string, int>();
                for (var column = 1; column <= lastColumn; column++)
                {
                    var text = CellText(worksheet.Cell(row, column));
                    if (text.Length > 0 && !columns.ContainsKey(text))
                    {
                        columns[text] = column;
                    }
                }

                if (RequiredHeaders.All(columns.ContainsKey))
                {
                    return new HeaderMap { Row = row, Columns = columns };
                }
            }

            throw new InvalidOperationException("重庆电量确认结算单缺少必要表头：用户名称、户号、时段、用电量。");
        }

        private static HeaderMap FindXlsHeader(IList<List<object>> rows)
        {
            var maxRows = Math.Min(rows.Count, 30);
            for (var row = 0; row < maxRows; row++)
            {
                var columns = new Dictionary<string, int>();
                for (var column = 0; column < rows[row].Count; column++)
                {
                    var text = TextUtil.S(rows[row][column]);
                    if (text.Length > 0 && !columns.ContainsKey(text))
                    {
                        columns[text] = column;
                    }
                }

                if (RequiredHeaders.All(columns.ContainsKey))
                {
                    return new HeaderMap { Row = row + 1, Columns = columns };
                }
            }

            throw new InvalidOperationException("重庆电量确认结算单缺少必要表头：用户名称、户号、时段、用电量。");
        }

        private static HeaderMap FindCsvHeader(IList<List<string>> rows)
        {
            var maxRows = Math.Min(rows.Count, 30);
            for (var row = 0; row < maxRows; row++)
            {
                var columns = new Dictionary<string, int>();
                for (var column = 0; column < rows[row].Count; column++)
                {
                    var text = TextUtil.S(rows[row][column]);
                    if (text.Length > 0 && !columns.ContainsKey(text))
                    {
                        columns[text] = column;
                    }
                }

                if (RequiredHeaders.All(columns.ContainsKey))
                {
                    return new HeaderMap { Row = row + 1, Columns = columns };
                }
            }

            throw new InvalidOperationException("重庆电量确认结算单缺少必要表头：用户名称、户号、时段、用电量。");
        }

        private static ChongqingPowerRawRow ReadXlsxRow(
            IXLWorksheet worksheet,
            int rowNumber,
            IDictionary<string, int> columns)
        {
            double power;
            var powerText = CellText(worksheet.Cell(rowNumber, columns["用电量"]));
            var hasPower = TryCellNumber(worksheet.Cell(rowNumber, columns["用电量"]), out power);
            return new ChongqingPowerRawRow
            {
                SourceRow = rowNumber,
                CustomerName = CellText(worksheet.Cell(rowNumber, columns["用户名称"])),
                AccountNumber = CellText(worksheet.Cell(rowNumber, columns["户号"])),
                PeriodText = CellText(worksheet.Cell(rowNumber, columns["时段"])),
                PowerText = powerText,
                HasPower = hasPower,
                Power = power
            };
        }

        private static ChongqingPowerRawRow ReadXlsRow(
            IList<object> values,
            int rowNumber,
            IDictionary<string, int> columns)
        {
            double power;
            var powerValue = ColumnValue(values, columns["用电量"]);
            var hasPower = TryObjectNumber(powerValue, out power);
            return new ChongqingPowerRawRow
            {
                SourceRow = rowNumber,
                CustomerName = TextUtil.S(ColumnValue(values, columns["用户名称"])),
                AccountNumber = TextUtil.S(ColumnValue(values, columns["户号"])),
                PeriodText = TextUtil.S(ColumnValue(values, columns["时段"])),
                PowerText = TextUtil.S(powerValue),
                HasPower = hasPower,
                Power = power
            };
        }

        private static ChongqingPowerRawRow ReadCsvRow(
            IList<string> values,
            int rowNumber,
            IDictionary<string, int> columns)
        {
            double power;
            var powerText = ColumnString(values, columns["用电量"]);
            var hasPower = double.TryParse(powerText.Replace(",", string.Empty), NumberStyles.Any, CultureInfo.InvariantCulture, out power);
            return new ChongqingPowerRawRow
            {
                SourceRow = rowNumber,
                CustomerName = ColumnString(values, columns["用户名称"]),
                AccountNumber = ColumnString(values, columns["户号"]),
                PeriodText = ColumnString(values, columns["时段"]),
                PowerText = powerText,
                HasPower = hasPower,
                Power = power
            };
        }

        private static void AddValidRow(
            ChongqingPowerRawRow raw,
            IList<ChongqingPowerRawRow> rows,
            IList<string> warnings,
            IList<string> invalidRows)
        {
            if (IsIgnorableTailRow(raw))
            {
                warnings.Add("跳过非电量数据行：第" + raw.SourceRow + "行。");
                return;
            }

            var errors = new List<string>();
            if (string.IsNullOrWhiteSpace(raw.CustomerName))
            {
                errors.Add("用户名称为空");
            }

            if (string.IsNullOrWhiteSpace(raw.AccountNumber))
            {
                errors.Add("户号为空");
            }

            string period;
            if (!TryNormalizePeriod(raw.PeriodText, out period))
            {
                errors.Add("时段不在尖峰/高峰/平段/低谷范围内");
            }

            if (!raw.HasPower)
            {
                errors.Add("用电量不是数字");
            }
            else if (raw.Power < 0)
            {
                errors.Add("用电量为负数");
            }

            if (errors.Count > 0)
            {
                invalidRows.Add("第" + raw.SourceRow + "行：" + string.Join("；", errors));
                return;
            }

            raw.NormalizedPeriod = period;
            rows.Add(raw);
        }

        private static bool IsIgnorableTailRow(ChongqingPowerRawRow raw)
        {
            return string.IsNullOrWhiteSpace(raw.AccountNumber)
                && string.IsNullOrWhiteSpace(raw.PeriodText)
                && string.IsNullOrWhiteSpace(raw.PowerText);
        }

        private static bool TryNormalizePeriod(string text, out string period)
        {
            switch (TextUtil.S(text))
            {
                case "尖峰":
                    period = "尖";
                    return true;
                case "高峰":
                    period = "峰";
                    return true;
                case "平段":
                    period = "平";
                    return true;
                case "低谷":
                    period = "谷";
                    return true;
                default:
                    period = string.Empty;
                    return false;
            }
        }

        private static List<ChongqingPowerAggregateRow> AggregateCustomerRows(IEnumerable<ChongqingPowerRawRow> rows)
        {
            return rows
                .GroupBy(row => TextUtil.CustomerKey(row.CustomerName))
                .Select(group => Aggregate(group.First().CustomerName, null, group))
                .OrderBy(row => row.CustomerName, StringComparer.CurrentCulture)
                .ToList();
        }

        private static List<ChongqingPowerAggregateRow> AggregateAccountRows(IEnumerable<ChongqingPowerRawRow> rows)
        {
            return rows
                .GroupBy(row => TextUtil.CustomerKey(row.CustomerName) + "\u001f" + TextUtil.CustomerKey(row.AccountNumber))
                .Select(group => Aggregate(group.First().CustomerName, group.First().AccountNumber, group))
                .OrderBy(row => row.CustomerName, StringComparer.CurrentCulture)
                .ThenBy(row => row.AccountNumber, StringComparer.Ordinal)
                .ToList();
        }

        private static ChongqingPowerAggregateRow Aggregate(
            string customerName,
            string accountNumber,
            IEnumerable<ChongqingPowerRawRow> rows)
        {
            var result = new ChongqingPowerAggregateRow
            {
                CustomerName = customerName,
                AccountNumber = accountNumber
            };

            foreach (var row in rows)
            {
                switch (row.NormalizedPeriod)
                {
                    case "尖":
                        result.Sharp += row.Power;
                        break;
                    case "峰":
                        result.Peak += row.Power;
                        break;
                    case "平":
                        result.Flat += row.Power;
                        break;
                    case "谷":
                        result.Valley += row.Power;
                        break;
                }
            }

            result.Round();
            return result;
        }

        private static void WriteWorkbook(
            string outputPath,
            int month,
            IList<ChongqingPowerAggregateRow> customerRows,
            IList<ChongqingPowerAggregateRow> accountRows)
        {
            FileAccessGuard.RequireWritableWorkbook(outputPath, "重庆电量处理表输出文件");

            using (var workbook = new XLWorkbook())
            {
                var summary = workbook.AddWorksheet("用户电量汇总");
                WriteCustomerSummarySheet(summary, month, customerRows);

                var details = workbook.AddWorksheet("户号明细");
                WriteAccountDetailSheet(details, month, accountRows);

                workbook.SaveAs(outputPath);
            }
        }

        private static void WriteCustomerSummarySheet(
            IXLWorksheet worksheet,
            int month,
            IList<ChongqingPowerAggregateRow> rows)
        {
            worksheet.Range("A1:A2").Merge();
            worksheet.Range("B1:B2").Merge();
            worksheet.Range("C1:F1").Merge();
            worksheet.Cell("A1").Value = "用户名称";
            worksheet.Cell("B1").Value = "总实际电量（兆瓦时）";
            worksheet.Cell("C1").Value = "实际电量（兆瓦时）";
            worksheet.Cell("C2").Value = "尖";
            worksheet.Cell("D2").Value = "峰";
            worksheet.Cell("E2").Value = "平";
            worksheet.Cell("F2").Value = "谷";

            for (var index = 0; index < rows.Count; index++)
            {
                var row = rows[index];
                var excelRow = index + 3;
                worksheet.Cell(excelRow, 1).Value = row.CustomerName;
                worksheet.Cell(excelRow, 2).Value = row.Total;
                worksheet.Cell(excelRow, 3).Value = row.Sharp;
                worksheet.Cell(excelRow, 4).Value = row.Peak;
                worksheet.Cell(excelRow, 5).Value = row.Flat;
                worksheet.Cell(excelRow, 6).Value = row.Valley;
            }

            ApplySummaryStyles(worksheet, rows.Count + 2);
        }

        private static void WriteAccountDetailSheet(
            IXLWorksheet worksheet,
            int month,
            IList<ChongqingPowerAggregateRow> rows)
        {
            worksheet.Range("A1:G1").Merge();
            worksheet.Cell("A1").Value = MonthPrefix(month) + "月重庆零售侧用户电量户号明细";
            var headers = new[] { "用户名称", "户号", "总实际电量（兆瓦时）", "尖", "峰", "平", "谷" };
            for (var column = 1; column <= headers.Length; column++)
            {
                worksheet.Cell(2, column).Value = headers[column - 1];
            }

            for (var index = 0; index < rows.Count; index++)
            {
                var row = rows[index];
                var excelRow = index + 3;
                worksheet.Cell(excelRow, 1).Value = row.CustomerName;
                worksheet.Cell(excelRow, 2).Value = row.AccountNumber;
                worksheet.Cell(excelRow, 3).Value = row.Total;
                worksheet.Cell(excelRow, 4).Value = row.Sharp;
                worksheet.Cell(excelRow, 5).Value = row.Peak;
                worksheet.Cell(excelRow, 6).Value = row.Flat;
                worksheet.Cell(excelRow, 7).Value = row.Valley;
            }

            ApplyDetailStyles(worksheet, rows.Count + 2);
        }

        private static void ApplySummaryStyles(IXLWorksheet worksheet, int lastRow)
        {
            worksheet.Range("A1:F2").Style.Font.Bold = true;
            worksheet.Range("A1:F2").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF4F4");
            worksheet.Range("A1:F" + Math.Max(lastRow, 2)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("A1:F" + Math.Max(lastRow, 2)).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Range("B3:F" + Math.Max(lastRow, 3)).Style.NumberFormat.Format = "0.0000";
            ApplyTableBorder(worksheet.Range("A1:F" + Math.Max(lastRow, 2)));
            worksheet.Column(1).Width = 42;
            for (var column = 2; column <= 6; column++)
            {
                worksheet.Column(column).Width = 18;
            }
        }

        private static void ApplyDetailStyles(IXLWorksheet worksheet, int lastRow)
        {
            worksheet.Row(1).Height = 26;
            worksheet.Cell("A1").Style.Font.Bold = true;
            worksheet.Cell("A1").Style.Font.FontSize = 14;
            worksheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("A2:G2").Style.Font.Bold = true;
            worksheet.Range("A2:G2").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF4F4");
            worksheet.Range("A1:G" + Math.Max(lastRow, 2)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range("A1:G" + Math.Max(lastRow, 2)).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Range("C3:G" + Math.Max(lastRow, 3)).Style.NumberFormat.Format = "0.0000";
            ApplyTableBorder(worksheet.Range("A1:G" + Math.Max(lastRow, 2)));
            worksheet.Column(1).Width = 42;
            worksheet.Column(2).Width = 24;
            for (var column = 3; column <= 7; column++)
            {
                worksheet.Column(column).Width = 18;
            }
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

        private static void WriteReport(string reportPath, ProvinceStage1CleanResult result)
        {
            var payload = new
            {
                province = ProvinceDisplayNames.GetName(result.Province),
                month = result.Month,
                unit = result.Unit,
                sourceSheetName = result.SourceSheetName,
                rawRows = result.RawRows,
                customerRows = result.CustomerRows,
                accountRows = result.AccountRows,
                totalPower = result.TotalPower,
                outputWorkbookPath = result.OutputWorkbookPath,
                warnings = result.Warnings
            };
            File.WriteAllText(reportPath, JsonConvert.SerializeObject(payload, Formatting.Indented), Encoding.UTF8);
        }

        private static int ResolveMonth(ProvinceStage1CleanOptions options, string titleText)
        {
            var month = ExtractMonth(titleText);
            if (month > 0)
            {
                return month;
            }

            month = ExtractMonth(Path.GetFileNameWithoutExtension(options.RawDetailPath));
            if (month > 0)
            {
                return month;
            }

            return options.Month;
        }

        private static int ExtractMonth(string text)
        {
            var match = Regex.Match(TextUtil.S(text), @"(?<!\d)(0?[1-9]|1[0-2])月");
            if (!match.Success)
            {
                return 0;
            }

            int month;
            return int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out month) ? month : 0;
        }

        private static string OutputWorkbookName(ProvinceStage1CleanOptions options, int month)
        {
            if (!string.IsNullOrWhiteSpace(options.OutputWorkbookName))
            {
                return options.OutputWorkbookName;
            }

            return MonthPrefix(month) + "月重庆零售侧用户电量数据处理表.xlsx";
        }

        private static string ReportName(int month)
        {
            return MonthPrefix(month) + "月重庆零售侧用户电量校验报告.json";
        }

        private static string MonthPrefix(int month)
        {
            return month > 0 ? month.ToString(CultureInfo.InvariantCulture) : "未识别月份";
        }

        private static void ThrowIfInvalid(IList<string> invalidRows)
        {
            if (invalidRows.Count == 0)
            {
                return;
            }

            var preview = string.Join(Environment.NewLine, invalidRows.Take(10));
            throw new InvalidOperationException("重庆电量确认结算单存在严重数据错误，已停止生成。" + Environment.NewLine + preview);
        }

        private static bool TryCellNumber(IXLCell cell, out double value)
        {
            if (cell.IsEmpty())
            {
                value = 0d;
                return false;
            }

            if (cell.DataType == XLDataType.Number)
            {
                value = cell.GetDouble();
                return true;
            }

            return double.TryParse(CellText(cell).Replace(",", string.Empty), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryObjectNumber(object raw, out double value)
        {
            if (raw == null)
            {
                value = 0d;
                return false;
            }

            if (raw is double)
            {
                value = (double)raw;
                return true;
            }

            if (raw is float)
            {
                value = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return true;
            }

            if (raw is int || raw is long || raw is short || raw is decimal)
            {
                value = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return true;
            }

            return double.TryParse(TextUtil.S(raw).Replace(",", string.Empty), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        private static object ColumnValue(IList<object> values, int index)
        {
            return index >= values.Count ? null : values[index];
        }

        private static string ColumnString(IList<string> values, int index)
        {
            return index >= values.Count ? string.Empty : TextUtil.S(values[index]);
        }

        private static string CellText(IXLCell cell)
        {
            return TextUtil.S(cell.GetFormattedString());
        }

        private static Encoding DetectCsvEncoding(string path)
        {
            var bytes = File.ReadAllBytes(path);
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                return new UTF8Encoding(true);
            }

            return Encoding.GetEncoding("GB18030");
        }

        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            var quoted = false;
            for (var i = 0; i < line.Length; i++)
            {
                var ch = line[i];
                if (ch == '"')
                {
                    if (quoted && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        quoted = !quoted;
                    }
                }
                else if (ch == ',' && !quoted)
                {
                    result.Add(sb.ToString());
                    sb.Length = 0;
                }
                else
                {
                    sb.Append(ch);
                }
            }

            result.Add(sb.ToString());
            return result;
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

        private sealed class HeaderMap
        {
            public int Row { get; set; }
            public Dictionary<string, int> Columns { get; set; }
        }

        private class SourceReadResult
        {
            public string SheetName { get; set; }
            public string TitleText { get; set; }
            public List<ChongqingPowerRawRow> Rows { get; set; }
            public List<string> Warnings { get; set; }
        }

        private sealed class XlsSheetReadResult : SourceReadResult
        {
            public List<string> InvalidRows { get; set; }
        }

        private sealed class ChongqingPowerRawRow
        {
            public int SourceRow { get; set; }
            public string CustomerName { get; set; }
            public string AccountNumber { get; set; }
            public string PeriodText { get; set; }
            public string PowerText { get; set; }
            public bool HasPower { get; set; }
            public double Power { get; set; }
            public string NormalizedPeriod { get; set; }
        }

        internal sealed class ChongqingPowerDataSet
        {
            public int Month { get; set; }
            public string Unit { get; set; }
            public string SourceSheetName { get; set; }
            public int RawRows { get; set; }
            public List<ChongqingPowerAggregateRow> CustomerRows { get; set; }
            public List<ChongqingPowerAggregateRow> AccountRows { get; set; }
            public List<string> Warnings { get; set; }
        }

        internal sealed class ChongqingPowerAggregateRow
        {
            public string CustomerName { get; set; }
            public string AccountNumber { get; set; }
            public double Total { get; private set; }
            public double Sharp { get; set; }
            public double Peak { get; set; }
            public double Flat { get; set; }
            public double Valley { get; set; }

            public void Round()
            {
                Sharp = Math.Round(Sharp, 4);
                Peak = Math.Round(Peak, 4);
                Flat = Math.Round(Flat, 4);
                Valley = Math.Round(Valley, 4);
                Total = Math.Round(Sharp + Peak + Flat + Valley, 4);
            }
        }
    }
}
