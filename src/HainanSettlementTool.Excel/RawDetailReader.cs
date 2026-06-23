using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ExcelDataReader;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class RawDetailReader
    {
        public List<PowerRow> Read(string rawDetailPath)
        {
            var extension = Path.GetExtension(rawDetailPath).ToLowerInvariant();
            if (extension == ".xlsx")
            {
                return ReadXlsx(rawDetailPath);
            }

            if (extension == ".xls")
            {
                return ReadXls(rawDetailPath);
            }

            if (extension == ".csv")
            {
                return ReadCsv(rawDetailPath);
            }

            throw new NotSupportedException("原始零售侧明细只支持 .xlsx、.xls 或 .csv。");
        }

        private static List<PowerRow> ReadXlsx(string path)
        {
            using (var workbook = new XLWorkbook(path))
            {
                var worksheet = workbook.Worksheets.First();
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                var rows = new List<PowerRow>();
                for (var row = 4; row <= lastRow; row++)
                {
                    var name = TextUtil.S(worksheet.Cell(row, 4).GetFormattedString());
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    rows.Add(new PowerRow
                    {
                        SourceRow = row,
                        Name = name,
                        Key = TextUtil.CustomerKey(name),
                        Total = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 9)),
                        Sharp = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 12)),
                        Peak = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 16)),
                        Flat = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 20)),
                        Valley = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 24))
                    });
                }

                return rows;
            }
        }

        private static List<PowerRow> ReadXls(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var rows = new List<PowerRow>();
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

                    rows.Add(new PowerRow
                    {
                        SourceRow = row,
                        Name = name,
                        Key = TextUtil.CustomerKey(name),
                        Total = ReaderNumber(reader, 8),
                        Sharp = ReaderNumber(reader, 11),
                        Peak = ReaderNumber(reader, 15),
                        Flat = ReaderNumber(reader, 19),
                        Valley = ReaderNumber(reader, 23)
                    });
                }

                return rows;
            }
        }

        private static List<PowerRow> ReadCsv(string path)
        {
            var rows = new List<PowerRow>();
            var lines = File.ReadAllLines(path, DetectCsvEncoding(path));
            for (var index = 3; index < lines.Length; index++)
            {
                var cols = SplitCsvLine(lines[index]);
                if (cols.Count <= 23 || string.IsNullOrWhiteSpace(cols[3]))
                {
                    continue;
                }

                var name = TextUtil.S(cols[3]);
                rows.Add(new PowerRow
                {
                    SourceRow = index + 1,
                    Name = name,
                    Key = TextUtil.CustomerKey(name),
                    Total = TextUtil.N(cols[8]),
                    Sharp = TextUtil.N(cols[11]),
                    Peak = TextUtil.N(cols[15]),
                    Flat = TextUtil.N(cols[19]),
                    Valley = TextUtil.N(cols[23])
                });
            }

            return rows;
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
