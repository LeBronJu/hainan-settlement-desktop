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

        [TestMethod]
        public void UpdateLedgerWritesChongqingPowerAndAccountCodesToCopiedLedger()
        {
            var root = CreateTempRoot();
            try
            {
                var raw = Path.Combine(root, "2026年05月售电公司电量确认结算单.xlsx");
                var ledger = Path.Combine(root, "重庆2026年售电结算台账.xlsx");
                var outputDirectory = Path.Combine(root, "out");
                WriteChongqingWorkbook(raw);
                WriteChongqingLedger(ledger);
                var gateway = new ClosedXmlStage1ExcelGateway();
                var options = new ProvinceStage1LedgerUpdateOptions
                {
                    Province = ProvinceCode.Chongqing,
                    Month = 5,
                    LedgerPath = ledger,
                    RawDetailPath = raw,
                    OutputDirectory = outputDirectory
                };

                var plan = gateway.PlanLedgerUpdate(options);
                var result = gateway.UpdateLedger(options);

                Assert.IsTrue(plan.RequiresConfirmation);
                Assert.AreEqual(3, plan.LedgerCustomerRows);
                Assert.AreEqual(2, plan.PowerCustomerRows);
                Assert.AreEqual(2, plan.MatchedRows);
                Assert.AreEqual(2, plan.CodeFillRows);
                Assert.AreEqual(1, plan.MultiAccountRows);
                Assert.AreEqual(1, plan.MissingInPowerRows);
                Assert.IsTrue(File.Exists(result.OutputLedgerPath));
                Assert.IsTrue(File.Exists(result.ReportPath));
                Assert.AreEqual(2, result.UpdatedPowerRows);
                Assert.AreEqual(2, result.CodeFillRows);

                using (var original = new XLWorkbook(ledger))
                using (var updated = new XLWorkbook(result.OutputLedgerPath))
                {
                    var originalSheet = original.Worksheet("Sheet1");
                    var updatedSheet = updated.Worksheet("Sheet1");
                    var customerA = FindLedgerRow(updatedSheet, "测试客户A");
                    var customerB = FindLedgerRow(updatedSheet, "测试客户B");
                    var customerC = FindLedgerRow(updatedSheet, "台账独有客户");

                    Assert.AreEqual(string.Empty, originalSheet.Cell(customerA, 2).GetString());
                    Assert.AreEqual("A-001、A-002", updatedSheet.Cell(customerA, 2).GetString());
                    Assert.AreEqual("B-001", updatedSheet.Cell(customerB, 2).GetString());
                    Assert.AreEqual(31.5, updatedSheet.Cell(customerA, 8).GetDouble(), 0.00001);
                    Assert.AreEqual(3.5, updatedSheet.Cell(customerA, 9).GetDouble(), 0.00001);
                    Assert.AreEqual(17, updatedSheet.Cell(customerA, 10).GetDouble(), 0.00001);
                    Assert.AreEqual(5, updatedSheet.Cell(customerA, 11).GetDouble(), 0.00001);
                    Assert.AreEqual(6, updatedSheet.Cell(customerA, 12).GetDouble(), 0.00001);
                    Assert.AreEqual(12, updatedSheet.Cell(customerB, 8).GetDouble(), 0.00001);
                    Assert.IsTrue(updatedSheet.Cell(customerC, 8).IsEmpty());
                }

                var report = JObject.Parse(File.ReadAllText(result.ReportPath));
                Assert.AreEqual("重庆", (string)report["province"]);
                Assert.AreEqual(5, (int)report["month"]);
                Assert.AreEqual(2, (int)report["updatedPowerRows"]);
                Assert.AreEqual(2, (int)report["codeFillRows"]);
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

        private static void WriteChongqingLedger(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet("Sheet1");
                ws.Range("A1:G1").Merge();
                ws.Cell("A1").Value = "重庆2026年售电结算台账";
                ws.Cell("A2").Value = "序号";
                ws.Cell("B2").Value = "电力用户编码";
                ws.Cell("C2").Value = "电力用户名称";
                ws.Cell("D2").Value = "合同年用电量（兆瓦时）";
                ws.Cell("E2").Value = "履约开始月份";
                ws.Cell("F2").Value = "履约结束月份";
                ws.Cell("G2").Value = "负责人";
                ws.Range("H1:N1").Merge();
                ws.Cell("H1").Value = "5月";
                ws.Cell("H2").Value = "总实际电量（兆瓦时）";
                ws.Cell("I2").Value = "实际电量（兆瓦时）";
                ws.Cell("I3").Value = "尖";
                ws.Cell("J3").Value = "峰";
                ws.Cell("K3").Value = "平";
                ws.Cell("L3").Value = "谷";
                ws.Cell("M2").Value = "峰平谷系数";
                ws.Cell("M3").Value = "峰_平";
                ws.Cell("N3").Value = "谷_平";

                WriteLedgerRow(ws, 4, 1, "测试客户A");
                WriteLedgerRow(ws, 5, 2, "测试客户B");
                WriteLedgerRow(ws, 6, 3, "台账独有客户");
                workbook.SaveAs(path);
            }
        }

        private static void WriteLedgerRow(IXLWorksheet ws, int row, int index, string customerName)
        {
            ws.Cell(row, 1).Value = index;
            ws.Cell(row, 3).Value = customerName;
            ws.Cell(row, 4).Value = 100;
            ws.Cell(row, 5).Value = 202601;
            ws.Cell(row, 6).Value = 202612;
            ws.Cell(row, 7).Value = "测试负责人";
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

        private static int FindLedgerRow(IXLWorksheet worksheet, string customerName)
        {
            var lastRow = worksheet.LastRowUsed().RowNumber();
            for (var row = 4; row <= lastRow; row++)
            {
                if (worksheet.Cell(row, 3).GetString() == customerName)
                {
                    return row;
                }
            }

            Assert.Fail("未找到台账客户：" + customerName);
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
