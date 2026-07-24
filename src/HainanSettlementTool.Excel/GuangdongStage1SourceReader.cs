using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class GuangdongStage1SourceReader
    {
        internal const string SourceSheetName = "零售结算明细";
        internal const string Unit = "兆瓦时";
        internal const decimal PowerTolerance = 0.00001m;

        private static readonly string[] RequiredHeaders =
        {
            "用户编号",
            "用户名称",
            "总实际用电量",
            "峰电量",
            "平电量",
            "谷电量",
            "峰_平",
            "谷_平"
        };

        public GuangdongStage1DataSet Read(string path, int selectedMonth)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheets
                    .FirstOrDefault(item => string.Equals(item.Name, SourceSheetName, StringComparison.Ordinal));
                if (worksheet == null)
                {
                    throw new InvalidOperationException("广东零售结算明细 workbook 缺少官方 sheet“" + SourceSheetName + "”。");
                }

                var header = FindHeader(worksheet);
                var rawRows = ReadRows(worksheet, header);
                var warnings = new List<string>();
                var customerRows = AggregateRows(rawRows, warnings);
                AddSameNameDifferentCodeWarnings(customerRows, warnings);

                return new GuangdongStage1DataSet
                {
                    Month = ResolveMonth(path, selectedMonth),
                    Unit = Unit,
                    SourceSheetName = worksheet.Name,
                    RawRows = rawRows.Count,
                    CustomerRows = customerRows,
                    Warnings = warnings
                };
            }
        }

        private static HeaderMap FindHeader(IXLWorksheet worksheet)
        {
            var lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 1, 30);
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var row = 1; row <= lastRow; row++)
            {
                var matches = RequiredHeaders.ToDictionary(
                    header => header,
                    header => new List<int>(),
                    StringComparer.Ordinal);
                for (var column = 1; column <= lastColumn; column++)
                {
                    var text = CellText(worksheet.Cell(row, column));
                    List<int> columns;
                    if (matches.TryGetValue(text, out columns))
                    {
                        columns.Add(column);
                    }
                }

                if (!RequiredHeaders.All(header => matches[header].Count > 0))
                {
                    continue;
                }

                var duplicate = RequiredHeaders.FirstOrDefault(header => matches[header].Count > 1);
                if (duplicate != null)
                {
                    throw new InvalidOperationException(
                        "广东零售结算明细第"
                        + row.ToString(CultureInfo.InvariantCulture)
                        + "行存在重复必要表头“"
                        + duplicate
                        + "”，无法安全读取。");
                }

                return new HeaderMap
                {
                    Row = row,
                    Columns = RequiredHeaders.ToDictionary(
                        header => header,
                        header => matches[header][0],
                        StringComparer.Ordinal)
                };
            }

            throw new InvalidOperationException(
                "广东零售结算明细缺少必要表头："
                + string.Join("、", RequiredHeaders)
                + "。");
        }

        private static List<GuangdongPowerRawRow> ReadRows(IXLWorksheet worksheet, HeaderMap header)
        {
            var rows = new List<GuangdongPowerRawRow>();
            var invalidRows = new List<string>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? header.Row;
            for (var rowNumber = header.Row + 1; rowNumber <= lastRow; rowNumber++)
            {
                if (RequiredHeaders.All(headerText => worksheet.Cell(rowNumber, header.Columns[headerText]).IsEmpty()))
                {
                    continue;
                }

                var errors = new List<string>();
                var rawCode = CellText(worksheet.Cell(rowNumber, header.Columns["用户编号"]));
                var rawName = CellText(worksheet.Cell(rowNumber, header.Columns["用户名称"]));
                var code = TextUtil.CustomerKey(rawCode);
                var customerName = NormalizeDisplayName(rawName);
                var customerNameKey = TextUtil.CustomerKey(rawName);
                if (code.Length == 0)
                {
                    errors.Add("用户编号为空");
                }

                if (customerNameKey.Length == 0)
                {
                    errors.Add("用户名称为空");
                }

                decimal total;
                decimal peak;
                decimal flat;
                decimal valley;
                ReadRequiredPower(worksheet.Cell(rowNumber, header.Columns["总实际用电量"]), "总实际用电量", errors, out total);
                ReadRequiredPower(worksheet.Cell(rowNumber, header.Columns["峰电量"]), "峰电量", errors, out peak);
                ReadRequiredPower(worksheet.Cell(rowNumber, header.Columns["平电量"]), "平电量", errors, out flat);
                ReadRequiredPower(worksheet.Cell(rowNumber, header.Columns["谷电量"]), "谷电量", errors, out valley);

                if (errors.Count == 0 && Math.Abs(total - peak - flat - valley) > PowerTolerance)
                {
                    errors.Add(
                        "总实际用电量与峰/平/谷合计不一致，差额为"
                        + (total - peak - flat - valley).ToString("0.#####", CultureInfo.InvariantCulture)
                        + "兆瓦时");
                }

                decimal peakFlatCoefficient;
                decimal valleyFlatCoefficient;
                var peakFlatState = ReadOptionalNumber(
                    worksheet.Cell(rowNumber, header.Columns["峰_平"]),
                    out peakFlatCoefficient);
                var valleyFlatState = ReadOptionalNumber(
                    worksheet.Cell(rowNumber, header.Columns["谷_平"]),
                    out valleyFlatCoefficient);
                var hasCompleteCoefficientPair = peakFlatState == OptionalNumberState.Valid
                    && valleyFlatState == OptionalNumberState.Valid
                    && peakFlatCoefficient >= 0m
                    && valleyFlatCoefficient >= 0m;
                var hasInvalidCoefficientPair = !hasCompleteCoefficientPair
                    && (peakFlatState != OptionalNumberState.Blank || valleyFlatState != OptionalNumberState.Blank);

                if (errors.Count > 0)
                {
                    invalidRows.Add(
                        "第"
                        + rowNumber.ToString(CultureInfo.InvariantCulture)
                        + "行："
                        + string.Join("；", errors));
                    continue;
                }

                rows.Add(new GuangdongPowerRawRow
                {
                    SourceRow = rowNumber,
                    Code = code,
                    CustomerName = customerName,
                    CustomerNameKey = customerNameKey,
                    Total = total,
                    Peak = peak,
                    Flat = flat,
                    Valley = valley,
                    HasCoefficientPair = hasCompleteCoefficientPair,
                    HasInvalidCoefficientPair = hasInvalidCoefficientPair,
                    PeakFlatCoefficient = peakFlatCoefficient,
                    ValleyFlatCoefficient = valleyFlatCoefficient
                });
            }

            if (invalidRows.Count > 0)
            {
                throw new InvalidOperationException(
                    "广东零售结算明细存在严重数据错误，已停止生成。"
                    + Environment.NewLine
                    + string.Join(Environment.NewLine, invalidRows.Take(10)));
            }

            if (rows.Count == 0)
            {
                throw new InvalidOperationException("广东零售结算明细没有可读取的客户电量行。");
            }

            return rows;
        }

        private static List<GuangdongPowerAggregateRow> AggregateRows(
            IEnumerable<GuangdongPowerRawRow> rows,
            IList<string> warnings)
        {
            var byCode = new Dictionary<string, GuangdongPowerAggregateRow>(StringComparer.Ordinal);
            var ordered = new List<GuangdongPowerAggregateRow>();
            var invalidRows = new List<string>();
            foreach (var row in rows)
            {
                GuangdongPowerAggregateRow aggregate;
                if (!byCode.TryGetValue(row.Code, out aggregate))
                {
                    aggregate = new GuangdongPowerAggregateRow
                    {
                        Code = row.Code,
                        CustomerName = row.CustomerName,
                        CustomerNameKey = row.CustomerNameKey,
                        FirstSourceRow = row.SourceRow
                    };
                    byCode.Add(row.Code, aggregate);
                    ordered.Add(aggregate);
                }
                else if (!string.Equals(aggregate.CustomerNameKey, row.CustomerNameKey, StringComparison.Ordinal))
                {
                    invalidRows.Add(
                        "用户编号“"
                        + row.Code
                        + "”在第"
                        + aggregate.FirstSourceRow.ToString(CultureInfo.InvariantCulture)
                        + "行和第"
                        + row.SourceRow.ToString(CultureInfo.InvariantCulture)
                        + "行对应不同用户名称：“"
                        + aggregate.CustomerName
                        + "”/“"
                        + row.CustomerName
                        + "”");
                    continue;
                }

                aggregate.SourceRows++;
                aggregate.Total += row.Total;
                aggregate.Peak += row.Peak;
                aggregate.Flat += row.Flat;
                aggregate.Valley += row.Valley;

                if (row.HasInvalidCoefficientPair)
                {
                    warnings.Add(
                        "第"
                        + row.SourceRow.ToString(CultureInfo.InvariantCulture)
                        + "行用户“"
                        + row.CustomerName
                        + "”的峰平谷系数不完整、非数字或为负数，已忽略该行系数。");
                }

                if (!row.HasCoefficientPair)
                {
                    continue;
                }

                if (!aggregate.HasCoefficientPair)
                {
                    aggregate.HasCoefficientPair = true;
                    aggregate.PeakFlatCoefficient = row.PeakFlatCoefficient;
                    aggregate.ValleyFlatCoefficient = row.ValleyFlatCoefficient;
                    aggregate.CoefficientSourceRow = row.SourceRow;
                    continue;
                }

                if (aggregate.PeakFlatCoefficient != row.PeakFlatCoefficient
                    || aggregate.ValleyFlatCoefficient != row.ValleyFlatCoefficient)
                {
                    aggregate.CoefficientConflictRows++;
                }
            }

            if (invalidRows.Count > 0)
            {
                throw new InvalidOperationException(
                    "广东零售结算明细存在严重数据错误，已停止生成。"
                    + Environment.NewLine
                    + string.Join(Environment.NewLine, invalidRows.Take(10)));
            }

            foreach (var aggregate in ordered)
            {
                var difference = aggregate.Total - aggregate.Peak - aggregate.Flat - aggregate.Valley;
                if (Math.Abs(difference) > PowerTolerance)
                {
                    throw new InvalidOperationException(
                        "广东零售结算明细聚合后用户“"
                        + aggregate.CustomerName
                        + "”（"
                        + aggregate.Code
                        + "）总实际用电量与峰/平/谷合计不一致，差额为"
                        + difference.ToString("0.#####", CultureInfo.InvariantCulture)
                        + "兆瓦时。");
                }

                if (aggregate.CoefficientConflictRows > 0)
                {
                    warnings.Add(
                        "用户“"
                        + aggregate.CustomerName
                        + "”（"
                        + aggregate.Code
                        + "）存在"
                        + aggregate.CoefficientConflictRows.ToString(CultureInfo.InvariantCulture)
                        + "条后续系数与首个有效系数不一致；已按原始行顺序采用第"
                        + aggregate.CoefficientSourceRow.ToString(CultureInfo.InvariantCulture)
                        + "行的峰_平="
                        + aggregate.PeakFlatCoefficient.ToString(CultureInfo.InvariantCulture)
                        + "、谷_平="
                        + aggregate.ValleyFlatCoefficient.ToString(CultureInfo.InvariantCulture)
                        + "。");
                }
            }

            return ordered;
        }

        private static void AddSameNameDifferentCodeWarnings(
            IEnumerable<GuangdongPowerAggregateRow> rows,
            IList<string> warnings)
        {
            foreach (var group in rows.GroupBy(row => row.CustomerNameKey, StringComparer.Ordinal).Where(group => group.Count() > 1))
            {
                warnings.Add(
                    "多个不同用户编号对应同一归一化用户名称“"
                    + group.First().CustomerName
                    + "”："
                    + string.Join("、", group.Select(row => row.Code))
                    + "；程序不会自动合并这些编码。");
            }
        }

        private static void ReadRequiredPower(
            IXLCell cell,
            string header,
            IList<string> errors,
            out decimal value)
        {
            if (!TryReadNumber(cell, out value))
            {
                errors.Add(header + "不是数字");
                return;
            }

            if (value < 0m)
            {
                errors.Add(header + "为负数");
            }
        }

        private static OptionalNumberState ReadOptionalNumber(IXLCell cell, out decimal value)
        {
            if (cell.IsEmpty() || string.IsNullOrWhiteSpace(CellText(cell)))
            {
                value = 0m;
                return OptionalNumberState.Blank;
            }

            return TryReadNumber(cell, out value)
                ? OptionalNumberState.Valid
                : OptionalNumberState.Invalid;
        }

        private static bool TryReadNumber(IXLCell cell, out decimal value)
        {
            if (cell.IsEmpty())
            {
                value = 0m;
                return false;
            }

            if (cell.TryGetValue(out value))
            {
                return true;
            }

            var text = CellText(cell).Replace(",", string.Empty);
            return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value)
                || decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value);
        }

        private static string NormalizeDisplayName(string value)
        {
            var normalized = TextUtil.S(value).Normalize(NormalizationForm.FormKC);
            var result = new StringBuilder(normalized.Length);
            var pendingSpace = false;
            foreach (var character in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(character) == UnicodeCategory.Format)
                {
                    continue;
                }

                if (char.IsWhiteSpace(character))
                {
                    pendingSpace = result.Length > 0;
                    continue;
                }

                if (pendingSpace)
                {
                    result.Append(' ');
                    pendingSpace = false;
                }

                result.Append(character);
            }

            return result.ToString().Trim();
        }

        private static int ResolveMonth(string path, int selectedMonth)
        {
            var fileName = Path.GetFileNameWithoutExtension(path);
            var match = Regex.Match(
                fileName ?? string.Empty,
                @"(?:^|[^\d])20\d{2}[-_.年](0?[1-9]|1[0-2])(?=月|[^\d]|$)",
                RegexOptions.CultureInvariant);
            int month;
            if (match.Success
                && int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out month))
            {
                return month;
            }

            return selectedMonth;
        }

        private static string CellText(IXLCell cell)
        {
            return TextUtil.S(cell.GetFormattedString());
        }

        private sealed class HeaderMap
        {
            public int Row { get; set; }
            public Dictionary<string, int> Columns { get; set; }
        }

        private enum OptionalNumberState
        {
            Blank,
            Valid,
            Invalid
        }

        private sealed class GuangdongPowerRawRow
        {
            public int SourceRow { get; set; }
            public string Code { get; set; }
            public string CustomerName { get; set; }
            public string CustomerNameKey { get; set; }
            public decimal Total { get; set; }
            public decimal Peak { get; set; }
            public decimal Flat { get; set; }
            public decimal Valley { get; set; }
            public bool HasCoefficientPair { get; set; }
            public bool HasInvalidCoefficientPair { get; set; }
            public decimal PeakFlatCoefficient { get; set; }
            public decimal ValleyFlatCoefficient { get; set; }
        }

        internal sealed class GuangdongStage1DataSet
        {
            public int Month { get; set; }
            public string Unit { get; set; }
            public string SourceSheetName { get; set; }
            public int RawRows { get; set; }
            public List<GuangdongPowerAggregateRow> CustomerRows { get; set; }
            public List<string> Warnings { get; set; }
        }

        internal sealed class GuangdongPowerAggregateRow
        {
            public string Code { get; set; }
            public string CustomerName { get; set; }
            public string CustomerNameKey { get; set; }
            public int FirstSourceRow { get; set; }
            public int SourceRows { get; set; }
            public decimal Total { get; set; }
            public decimal Peak { get; set; }
            public decimal Flat { get; set; }
            public decimal Valley { get; set; }
            public bool HasCoefficientPair { get; set; }
            public decimal PeakFlatCoefficient { get; set; }
            public decimal ValleyFlatCoefficient { get; set; }
            public int CoefficientSourceRow { get; set; }
            public int CoefficientConflictRows { get; set; }
        }
    }
}
