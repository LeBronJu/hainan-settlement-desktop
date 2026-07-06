using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class ChongqingPowerCleanGeneratorTests
    {
        [TestMethod]
        public void CleanPowerDataAggregatesChongqingTransactionRowsByCustomerAndAccount()
        {
            var root = CreateTempRoot();
            try
            {
                var input = Path.Combine(root, "2026年05月售电公司电量确认结算单.xlsx");
                var outputDirectory = Path.Combine(root, "out");
                WriteChongqingWorkbook(input);

                var result = new ClosedXmlStage1ExcelGateway().CleanPowerData(new ProvinceStage1CleanOptions
                {
                    Province = ProvinceCode.Chongqing,
                    RawDetailPath = input,
                    OutputDirectory = outputDirectory
                });

                Assert.AreEqual(ProvinceCode.Chongqing, result.Province);
                Assert.AreEqual(5, result.Month);
                Assert.AreEqual("兆瓦时", result.Unit);
                Assert.AreEqual("sheet1", result.SourceSheetName);
                Assert.AreEqual(12, result.RawRows);
                Assert.AreEqual(2, result.CustomerRows);
                Assert.AreEqual(3, result.AccountRows);
                Assert.AreEqual(43.5, result.TotalPower, 0.00001);
                Assert.IsTrue(File.Exists(result.OutputWorkbookPath));
                Assert.IsTrue(File.Exists(result.ReportPath));
                StringAssert.EndsWith(result.OutputWorkbookPath, "5月重庆零售侧用户电量数据处理表.xlsx");

                using (var workbook = new XLWorkbook(result.OutputWorkbookPath))
                {
                    var summary = workbook.Worksheet("用户电量汇总");
                    Assert.AreEqual("用户名称", summary.Cell("A1").GetString());
                    Assert.AreEqual("总实际电量（兆瓦时）", summary.Cell("B1").GetString());
                    Assert.AreEqual("实际电量（兆瓦时）", summary.Cell("C1").GetString());
                    Assert.AreEqual("尖", summary.Cell("C2").GetString());
                    Assert.AreEqual("峰", summary.Cell("D2").GetString());
                    Assert.AreEqual("平", summary.Cell("E2").GetString());
                    Assert.AreEqual("谷", summary.Cell("F2").GetString());

                    var customerA = FindRow(summary, "测试客户A");
                    Assert.AreEqual(31.5, summary.Cell(customerA, 2).GetDouble(), 0.00001);
                    Assert.AreEqual(3.5, summary.Cell(customerA, 3).GetDouble(), 0.00001);
                    Assert.AreEqual(17, summary.Cell(customerA, 4).GetDouble(), 0.00001);
                    Assert.AreEqual(5, summary.Cell(customerA, 5).GetDouble(), 0.00001);
                    Assert.AreEqual(6, summary.Cell(customerA, 6).GetDouble(), 0.00001);

                    var detail = workbook.Worksheet("户号明细");
                    Assert.AreEqual(3, CountDataRows(detail));
                    var accountA1 = FindAccountRow(detail, "测试客户A", "A-001");
                    Assert.AreEqual(22.5, detail.Cell(accountA1, 3).GetDouble(), 0.00001);
                    Assert.AreEqual(12, detail.Cell(accountA1, 5).GetDouble(), 0.00001);
                }

                var report = JObject.Parse(File.ReadAllText(result.ReportPath));
                Assert.AreEqual("重庆", (string)report["province"]);
                Assert.AreEqual(5, (int)report["month"]);
                Assert.AreEqual(2, (int)report["customerRows"]);
                Assert.AreEqual(3, (int)report["accountRows"]);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void CleanPowerDataStopsOnNegativeChongqingPower()
        {
            var root = CreateTempRoot();
            try
            {
                var input = Path.Combine(root, "2026年05月售电公司电量确认结算单.xlsx");
                WriteChongqingWorkbook(input, includeNegative: true);

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                    new ClosedXmlStage1ExcelGateway().CleanPowerData(new ProvinceStage1CleanOptions
                    {
                        Province = ProvinceCode.Chongqing,
                        RawDetailPath = input,
                        OutputDirectory = Path.Combine(root, "out")
                    }));

                StringAssert.Contains(ex.Message, "严重数据错误");
                StringAssert.Contains(ex.Message, "用电量为负数");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static void WriteChongqingWorkbook(string path, bool includeNegative = false)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("已处理");
                var ws = workbook.AddWorksheet("sheet1");
                ws.Range("A1:I1").Merge();
                ws.Cell("A1").Value = "重庆电力交易中心有限公司2026年05月售电公司电量确认单";
                ws.Cell("A2").Value = "售电公司：测试售电公司";
                ws.Cell("A3").Value = "单位：兆瓦时、元/兆瓦时、元";
                var headers = new[] { "用户名称", "户号", "计量点ID", "是否分时", "是否绿电", "时段", "用电量", "电价", "电费" };
                for (var column = 1; column <= headers.Length; column++)
                {
                    ws.Cell(4, column).Value = headers[column - 1];
                }

                var row = 5;
                WriteRow(ws, row++, "测试客户A", "A-001", "M1", "尖峰", 1.5);
                WriteRow(ws, row++, "测试客户A", "A-001", "M2", "尖峰", 2.0);
                WriteRow(ws, row++, "测试客户A", "A-001", "M1", "高峰", 10.0);
                WriteRow(ws, row++, "测试客户A", "A-001", "M2", "高峰", 2.0);
                WriteRow(ws, row++, "测试客户A", "A-001", "M1", "平段", 4.0);
                WriteRow(ws, row++, "测试客户A", "A-001", "M1", "低谷", 3.0);
                WriteRow(ws, row++, "测试客户A", "A-002", "M3", "尖峰", 0.0);
                WriteRow(ws, row++, "测试客户A", "A-002", "M3", "高峰", 5.0);
                WriteRow(ws, row++, "测试客户A", "A-002", "M3", "平段", 1.0);
                WriteRow(ws, row++, "测试客户A", "A-002", "M3", "低谷", 3.0);
                WriteRow(ws, row++, "测试客户B", "B-001", "M4", "尖峰", includeNegative ? -1.0 : 4.0);
                WriteRow(ws, row++, "测试客户B", "B-001", "M4", "高峰", 8.0);
                ws.Cell(row, 1).Value = "确认人：";

                workbook.SaveAs(path);
            }
        }

        private static void WriteRow(IXLWorksheet ws, int row, string customer, string account, string meter, string period, double power)
        {
            ws.Cell(row, 1).Value = customer;
            ws.Cell(row, 2).Value = account;
            ws.Cell(row, 3).Value = meter;
            ws.Cell(row, 4).Value = "是";
            ws.Cell(row, 5).Value = "否";
            ws.Cell(row, 6).Value = period;
            ws.Cell(row, 7).Value = power;
            ws.Cell(row, 8).Value = 100;
            ws.Cell(row, 9).Value = power * 100;
        }

        private static int FindRow(IXLWorksheet worksheet, string customerName)
        {
            var lastRow = worksheet.LastRowUsed().RowNumber();
            for (var row = 3; row <= lastRow; row++)
            {
                if (worksheet.Cell(row, 1).GetString() == customerName)
                {
                    return row;
                }
            }

            Assert.Fail("未找到客户：" + customerName);
            return 0;
        }

        private static int FindAccountRow(IXLWorksheet worksheet, string customerName, string account)
        {
            var lastRow = worksheet.LastRowUsed().RowNumber();
            for (var row = 3; row <= lastRow; row++)
            {
                if (worksheet.Cell(row, 1).GetString() == customerName && worksheet.Cell(row, 2).GetString() == account)
                {
                    return row;
                }
            }

            Assert.Fail("未找到户号明细：" + customerName + " / " + account);
            return 0;
        }

        private static int CountDataRows(IXLWorksheet worksheet)
        {
            return worksheet.RowsUsed()
                .Count(row => row.RowNumber() >= 3 && !string.IsNullOrWhiteSpace(row.Cell(1).GetString()));
        }

        private static string CreateTempRoot()
        {
            return Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
        }

        private static void DeleteTempRoot(string root)
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }
}
