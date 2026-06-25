using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class RawDetailReaderTests
    {
        [TestMethod]
        public void RawCsvDetailProvidesPowerRowsAndCustomerCodes()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var path = Path.Combine(root, "raw.csv");

            try
            {
                Directory.CreateDirectory(root);
                File.WriteAllLines(
                    path,
                    new[] { "header1", "header2", "header3", RawDetailLine("户号001", "测试客户", 1.1, 2.2, 3.3, 4.4, 5.5) },
                    Encoding.GetEncoding("GB18030"));

                var gateway = new ClosedXmlStage1ExcelGateway();
                var powerRows = gateway.ReadRawPowerRows(path);
                var customerCodes = gateway.ReadCustomerCodes(path);

                Assert.AreEqual(1, powerRows.Count);
                Assert.AreEqual(4, powerRows[0].SourceRow);
                Assert.AreEqual("测试客户", powerRows[0].Name);
                Assert.AreEqual(1.1, powerRows[0].Total, 0.00001);
                Assert.AreEqual(2.2, powerRows[0].Sharp, 0.00001);
                Assert.AreEqual(3.3, powerRows[0].Peak, 0.00001);
                Assert.AreEqual(4.4, powerRows[0].Flat, 0.00001);
                Assert.AreEqual(5.5, powerRows[0].Valley, 0.00001);
                Assert.AreEqual("户号001", customerCodes[powerRows[0].Key]);
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        [TestMethod]
        public void RawXlsxDetailKeepsPowerRowsOnFirstSheetAndCustomerCodesOnNamedSheets()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var path = Path.Combine(root, "raw.xlsx");

            try
            {
                Directory.CreateDirectory(root);
                using (var workbook = new XLWorkbook())
                {
                    WriteRawDetailRow(workbook.AddWorksheet("其它明细"), 4, "忽略户号", "第一张表客户", 10, 11, 12, 13, 14);
                    WriteRawDetailRow(workbook.AddWorksheet("零售主体电量"), 4, "主体户号", "主体客户", 20, 21, 22, 23, 24);
                    WriteRawDetailRow(workbook.AddWorksheet("零售户号电量"), 4, "户号表户号", "户号表客户", 30, 31, 32, 33, 34);
                    workbook.SaveAs(path);
                }

                var gateway = new ClosedXmlStage1ExcelGateway();
                var powerRows = gateway.ReadRawPowerRows(path);
                var customerCodes = gateway.ReadCustomerCodes(path);

                Assert.AreEqual(1, powerRows.Count);
                Assert.AreEqual("第一张表客户", powerRows[0].Name);
                Assert.AreEqual(10, powerRows[0].Total, 0.00001);
                CollectionAssert.AreEquivalent(
                    new[] { "主体户号", "户号表户号" },
                    customerCodes.Values.ToArray());
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        private static string RawDetailLine(
            string code,
            string name,
            double total,
            double sharp,
            double peak,
            double flat,
            double valley)
        {
            var columns = Enumerable.Repeat(string.Empty, 24).ToArray();
            columns[2] = code;
            columns[3] = name;
            columns[8] = total.ToString("0.####");
            columns[11] = sharp.ToString("0.####");
            columns[15] = peak.ToString("0.####");
            columns[19] = flat.ToString("0.####");
            columns[23] = valley.ToString("0.####");
            return string.Join(",", columns);
        }

        private static void WriteRawDetailRow(
            IXLWorksheet worksheet,
            int row,
            string code,
            string name,
            double total,
            double sharp,
            double peak,
            double flat,
            double valley)
        {
            worksheet.Cell(row, 3).Value = code;
            worksheet.Cell(row, 4).Value = name;
            worksheet.Cell(row, 9).Value = total;
            worksheet.Cell(row, 12).Value = sharp;
            worksheet.Cell(row, 16).Value = peak;
            worksheet.Cell(row, 20).Value = flat;
            worksheet.Cell(row, 24).Value = valley;
        }
    }
}
