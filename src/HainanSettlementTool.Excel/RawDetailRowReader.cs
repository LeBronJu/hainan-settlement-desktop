using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ExcelDataReader;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal enum RawDetailSheetSelection
    {
        FirstSheet,
        CustomerCodeSheets
    }

    internal sealed class RawDetailRowReader
    {
        private static readonly string[] CustomerCodeSheetNames = { "零售主体电量", "零售户号电量" };

        public List<RawDetailRow> Read(string rawDetailPath, RawDetailSheetSelection sheetSelection)
        {
            var extension = Path.GetExtension(rawDetailPath).ToLowerInvariant();
            if (extension == ".xlsx")
            {
                return ReadXlsx(rawDetailPath, sheetSelection);
            }

            if (extension == ".xls")
            {
                return ReadXls(rawDetailPath, sheetSelection);
            }

            if (extension == ".csv")
            {
                return ReadCsv(rawDetailPath);
            }

            throw new NotSupportedException("原始零售侧明细只支持 .xlsx、.xls 或 .csv。");
        }

        public static bool IsSupported(string rawDetailPath)
        {
            var extension = Path.GetExtension(rawDetailPath);
            return string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase);
        }

        private static List<RawDetailRow> ReadXlsx(string path, RawDetailSheetSelection sheetSelection)
        {
            using (var workbook = new XLWorkbook(path))
            {
                return SelectWorksheets(workbook, sheetSelection)
                    .SelectMany(ReadXlsxSheet)
                    .ToList();
            }
        }

        private static IEnumerable<IXLWorksheet> SelectWorksheets(XLWorkbook workbook, RawDetailSheetSelection sheetSelection)
        {
            if (sheetSelection == RawDetailSheetSelection.FirstSheet)
            {
                yield return workbook.Worksheets.First();
                yield break;
            }

            var namedSheets = CustomerCodeSheetNames
                .Where(name => workbook.Worksheets.Any(ws => ws.Name == name))
                .Select(name => workbook.Worksheet(name))
                .ToList();
            if (namedSheets.Count == 0)
            {
                yield return workbook.Worksheets.First();
                yield break;
            }

            foreach (var worksheet in namedSheets)
            {
                yield return worksheet;
            }
        }

        private static IEnumerable<RawDetailRow> ReadXlsxSheet(IXLWorksheet worksheet)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            for (var row = 4; row <= lastRow; row++)
            {
                var name = TextUtil.S(worksheet.Cell(row, 4).GetFormattedString());
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                yield return new RawDetailRow
                {
                    SourceRow = row,
                    CustomerCode = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString()),
                    Name = name,
                    Key = TextUtil.CustomerKey(name),
                    Total = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 9)),
                    Sharp = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 12)),
                    Peak = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 16)),
                    Flat = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 20)),
                    Valley = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 24)),
                    HasPowerColumns = true
                };
            }
        }

        private static List<RawDetailRow> ReadXls(string path, RawDetailSheetSelection sheetSelection)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                if (sheetSelection == RawDetailSheetSelection.FirstSheet)
                {
                    return ReadXlsSheet(reader);
                }

                var namedSheetRows = new List<RawDetailRow>();
                var firstSheetRows = new List<RawDetailRow>();
                var foundNamedSheet = false;
                var isFirstSheet = true;

                do
                {
                    var isNamedSheet = CustomerCodeSheetNames.Contains(reader.Name);
                    if (isNamedSheet)
                    {
                        namedSheetRows.AddRange(ReadXlsSheet(reader));
                        foundNamedSheet = true;
                    }
                    else if (isFirstSheet)
                    {
                        firstSheetRows.AddRange(ReadXlsSheet(reader));
                    }
                    else
                    {
                        DrainSheet(reader);
                    }

                    isFirstSheet = false;
                }
                while (reader.NextResult());

                return foundNamedSheet ? namedSheetRows : firstSheetRows;
            }
        }

        private static List<RawDetailRow> ReadXlsSheet(IExcelDataReader reader)
        {
            var rows = new List<RawDetailRow>();
            var row = 0;
            while (reader.Read())
            {
                row++;
                if (row < 4)
                {
                    continue;
                }

                var name = ReaderString(reader, 3);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                rows.Add(new RawDetailRow
                {
                    SourceRow = row,
                    CustomerCode = ReaderString(reader, 2),
                    Name = name,
                    Key = TextUtil.CustomerKey(name),
                    Total = ReaderNumber(reader, 8),
                    Sharp = ReaderNumber(reader, 11),
                    Peak = ReaderNumber(reader, 15),
                    Flat = ReaderNumber(reader, 19),
                    Valley = ReaderNumber(reader, 23),
                    HasPowerColumns = true
                });
            }

            return rows;
        }

        private static void DrainSheet(IExcelDataReader reader)
        {
            while (reader.Read())
            {
            }
        }

        private static List<RawDetailRow> ReadCsv(string path)
        {
            var rows = new List<RawDetailRow>();
            var lines = File.ReadAllLines(path, DetectCsvEncoding(path));
            for (var index = 3; index < lines.Length; index++)
            {
                var cols = SplitCsvLine(lines[index]);
                if (cols.Count <= 3 || string.IsNullOrWhiteSpace(cols[3]))
                {
                    continue;
                }

                var name = TextUtil.S(cols[3]);
                rows.Add(new RawDetailRow
                {
                    SourceRow = index + 1,
                    CustomerCode = ColumnString(cols, 2),
                    Name = name,
                    Key = TextUtil.CustomerKey(name),
                    Total = ColumnNumber(cols, 8),
                    Sharp = ColumnNumber(cols, 11),
                    Peak = ColumnNumber(cols, 15),
                    Flat = ColumnNumber(cols, 19),
                    Valley = ColumnNumber(cols, 23),
                    HasPowerColumns = cols.Count > 23
                });
            }

            return rows;
        }

        private static string ColumnString(IList<string> columns, int index)
        {
            return index >= columns.Count ? string.Empty : TextUtil.S(columns[index]);
        }

        private static double ColumnNumber(IList<string> columns, int index)
        {
            return index >= columns.Count ? 0d : TextUtil.N(columns[index]);
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

        private static string ReaderString(IExcelDataReader reader, int index)
        {
            if (index >= reader.FieldCount)
            {
                return string.Empty;
            }

            return TextUtil.S(Convert.ToString(reader.GetValue(index), CultureInfo.CurrentCulture));
        }

        private static double ReaderNumber(IExcelDataReader reader, int index)
        {
            if (index >= reader.FieldCount)
            {
                return 0d;
            }

            var value = reader.GetValue(index);
            if (value == null)
            {
                return 0d;
            }

            if (value is double d)
            {
                return d;
            }

            if (value is float f)
            {
                return f;
            }

            if (value is int i)
            {
                return i;
            }

            if (value is decimal m)
            {
                return (double)m;
            }

            return TextUtil.N(Convert.ToString(value, CultureInfo.CurrentCulture));
        }
    }
}
