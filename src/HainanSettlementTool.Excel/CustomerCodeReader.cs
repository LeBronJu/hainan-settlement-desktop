using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelDataReader;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class CustomerCodeReader
    {
        public Dictionary<string, string> Read(string rawDetailPath)
        {
            var result = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(rawDetailPath) || !File.Exists(rawDetailPath))
            {
                return result;
            }

            var extension = Path.GetExtension(rawDetailPath);
            if (string.Equals(extension, ".xls", System.StringComparison.OrdinalIgnoreCase))
            {
                return ReadXls(rawDetailPath);
            }

            if (!string.Equals(extension, ".xlsx", System.StringComparison.OrdinalIgnoreCase))
            {
                return result;
            }

            using (var workbook = new XLWorkbook(rawDetailPath))
            {
                var sheetNames = new[] { "零售主体电量", "零售户号电量" }
                    .Where(name => workbook.Worksheets.Any(ws => ws.Name == name))
                    .ToList();
                if (sheetNames.Count == 0)
                {
                    sheetNames.Add(workbook.Worksheets.First().Name);
                }

                foreach (var sheetName in sheetNames)
                {
                    var worksheet = workbook.Worksheet(sheetName);
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    for (var row = 4; row <= lastRow; row++)
                    {
                        var code = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                        var name = TextUtil.S(worksheet.Cell(row, 4).GetFormattedString());
                        var key = TextUtil.CustomerKey(name);
                        if (key.Length > 0 && code.Length > 0 && !result.ContainsKey(key))
                        {
                            result[key] = code;
                        }
                    }
                }
            }

            return result;
        }

        private static Dictionary<string, string> ReadXls(string rawDetailPath)
        {
            var namedSheetResult = new Dictionary<string, string>();
            var firstSheetResult = new Dictionary<string, string>();
            var foundNamedSheet = false;
            var isFirstSheet = true;

            using (var stream = File.Open(rawDetailPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    var isNamedSheet = reader.Name == "零售主体电量" || reader.Name == "零售户号电量";
                    Dictionary<string, string> target = null;
                    if (isNamedSheet)
                    {
                        target = namedSheetResult;
                        foundNamedSheet = true;
                    }
                    else if (isFirstSheet)
                    {
                        target = firstSheetResult;
                    }

                    if (target != null)
                    {
                        ReadCustomerCodesFromSheet(reader, target);
                    }
                    else
                    {
                        while (reader.Read())
                        {
                        }
                    }

                    isFirstSheet = false;
                }
                while (reader.NextResult());
            }

            return foundNamedSheet ? namedSheetResult : firstSheetResult;
        }

        private static void ReadCustomerCodesFromSheet(IExcelDataReader reader, Dictionary<string, string> result)
        {
            var row = 0;
            while (reader.Read())
            {
                row++;
                if (row < 4)
                {
                    continue;
                }

                var code = ReaderString(reader, 2);
                var name = ReaderString(reader, 3);
                var key = TextUtil.CustomerKey(name);
                if (key.Length > 0 && code.Length > 0 && !result.ContainsKey(key))
                {
                    result[key] = code;
                }
            }
        }

        private static string ReaderString(IExcelDataReader reader, int index)
        {
            if (index >= reader.FieldCount)
            {
                return string.Empty;
            }

            return TextUtil.S(System.Convert.ToString(reader.GetValue(index), CultureInfo.CurrentCulture));
        }
    }
}
