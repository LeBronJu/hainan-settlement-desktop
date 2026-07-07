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

                var result = new ClosedXmlSettlementExcelGateway().CleanPowerData(new ProvinceStage1CleanOptions
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
                    new ClosedXmlSettlementExcelGateway().CleanPowerData(new ProvinceStage1CleanOptions
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
        public void UpdateLedgerWritesChongqingPowerOnlyToCopiedLedger()
        {
            var root = CreateTempRoot();
            try
            {
                var raw = Path.Combine(root, "2026年05月售电公司电量确认结算单.xlsx");
                var ledger = Path.Combine(root, "重庆2026年售电结算台账.xlsx");
                var outputDirectory = Path.Combine(root, "out");
                WriteChongqingWorkbook(raw);
                WriteChongqingLedger(ledger);
                var gateway = new ClosedXmlSettlementExcelGateway();
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
                Assert.AreEqual(1, plan.MultiAccountRows);
                Assert.AreEqual(1, plan.MissingInPowerRows);
                Assert.IsTrue(plan.Issues.All(issue => !string.IsNullOrWhiteSpace(issue.Kind)));
                CollectionAssert.Contains(
                    plan.Issues.Select(issue => issue.Kind).ToList(),
                    ProvinceStage1LedgerUpdateIssueKinds.MultiAccountCustomer);
                CollectionAssert.Contains(
                    plan.Issues.Select(issue => issue.Kind).ToList(),
                    ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower);
                Assert.IsTrue(File.Exists(result.OutputLedgerPath));
                Assert.IsTrue(File.Exists(result.ReportPath));
                Assert.AreEqual(2, result.UpdatedPowerRows);
                Assert.AreEqual(1, result.MultiAccountRows);

                using (var original = new XLWorkbook(ledger))
                using (var updated = new XLWorkbook(result.OutputLedgerPath))
                {
                    var originalSheet = original.Worksheet("Sheet1");
                    var updatedSheet = updated.Worksheet("Sheet1");
                    var customerA = FindLedgerRow(updatedSheet, "测试客户A");
                    var customerB = FindLedgerRow(updatedSheet, "测试客户B");
                    var customerC = FindLedgerRow(updatedSheet, "台账独有客户");

                    Assert.AreEqual(string.Empty, originalSheet.Cell(customerA, 2).GetString());
                    Assert.AreEqual(string.Empty, updatedSheet.Cell(customerA, 2).GetString());
                    Assert.AreEqual(string.Empty, updatedSheet.Cell(customerB, 2).GetString());
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
                Assert.AreEqual(1, (int)report["multiAccountRows"]);
                Assert.AreEqual(
                    ProvinceStage1LedgerUpdateIssueKinds.MultiAccountCustomer,
                    (string)report["issues"][0]["Kind"]);
                Assert.AreEqual("多户号客户", (string)report["issues"][0]["Category"]);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void UpdateLedgerUsesManualCustomerMatchForChongqingAlias()
        {
            var root = CreateTempRoot();
            try
            {
                var raw = Path.Combine(root, "2026年05月售电公司电量确认结算单.xlsx");
                var ledger = Path.Combine(root, "重庆2026年售电结算台账.xlsx");
                var outputDirectory = Path.Combine(root, "out");
                WriteChongqingWorkbook(raw, customerBName: "测试客户B旧名");
                WriteChongqingLedger(ledger, customerBName: "测试客户B新名");
                var gateway = new ClosedXmlSettlementExcelGateway();
                var options = new ProvinceStage1LedgerUpdateOptions
                {
                    Province = ProvinceCode.Chongqing,
                    Month = 5,
                    LedgerPath = ledger,
                    RawDetailPath = raw,
                    OutputDirectory = outputDirectory
                };

                var plan = gateway.PlanLedgerUpdate(options);
                CollectionAssert.Contains(plan.PowerOnlyCustomers, "测试客户B旧名");
                CollectionAssert.Contains(plan.LedgerOnlyCustomers, "测试客户B新名");

                options.ManualCustomerMatches.Add(new ProvinceStage1CustomerMatch
                {
                    SourceCustomerName = "测试客户B旧名",
                    TargetCustomerName = "测试客户B新名"
                });

                var result = gateway.UpdateLedger(options);

                Assert.AreEqual(2, result.MatchedRows);
                Assert.AreEqual(1, result.ManualMatchedRows);
                Assert.AreEqual(2, result.UpdatedPowerRows);
                Assert.AreEqual(1, result.ManualCustomerMatches.Count);
                Assert.AreEqual("测试客户B旧名", result.ManualCustomerMatches[0].SourceCustomerName);
                Assert.AreEqual("测试客户B新名", result.ManualCustomerMatches[0].TargetCustomerName);

                using (var updated = new XLWorkbook(result.OutputLedgerPath))
                {
                    var updatedSheet = updated.Worksheet("Sheet1");
                    var customerA = FindLedgerRow(updatedSheet, "测试客户A");
                    var customerB = FindLedgerRow(updatedSheet, "测试客户B新名");
                    Assert.AreEqual(31.5, updatedSheet.Cell(customerA, 8).GetDouble(), 0.00001);
                    Assert.AreEqual(12, updatedSheet.Cell(customerB, 8).GetDouble(), 0.00001);
                    Assert.AreEqual(4, updatedSheet.Cell(customerB, 9).GetDouble(), 0.00001);
                    Assert.AreEqual(8, updatedSheet.Cell(customerB, 10).GetDouble(), 0.00001);
                }

                var report = JObject.Parse(File.ReadAllText(result.ReportPath));
                Assert.AreEqual(1, (int)report["manualMatchedRows"]);
                Assert.AreEqual("测试客户B旧名", (string)report["manualCustomerMatches"][0]["sourceCustomerName"]);
                Assert.AreEqual("测试客户B新名", (string)report["manualCustomerMatches"][0]["targetCustomerName"]);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void UpdateLedgerCreatesTargetMonthBlockFromPreviousChongqingLedger()
        {
            var root = CreateTempRoot();
            try
            {
                var raw = Path.Combine(root, "2026年05月售电公司电量确认结算单.xlsx");
                var ledger = Path.Combine(root, "重庆2026年售电结算台账-4月.xlsx");
                var outputDirectory = Path.Combine(root, "out");
                WriteChongqingWorkbook(raw);
                WriteChongqingLedgerWithPreviousMonthOnly(ledger);
                var gateway = new ClosedXmlSettlementExcelGateway();
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

                Assert.IsTrue(plan.Warnings.Any(item => item.Contains("已基于4月电量区块创建5月电量区块")));
                Assert.IsTrue(result.Warnings.Any(item => item.Contains("已基于4月电量区块创建5月电量区块")));
                Assert.AreEqual(2, result.UpdatedPowerRows);

                using (var updated = new XLWorkbook(result.OutputLedgerPath))
                {
                    var ws = updated.Worksheet("Sheet1");
                    var aprStart = FindMonthStartColumn(ws, "4月");
                    var mayStart = FindMonthStartColumn(ws, "5月");
                    Assert.AreEqual(aprStart + 30, mayStart);
                    Assert.AreEqual("总实际电量（兆瓦时）", ws.Cell(2, mayStart).GetString());
                    Assert.AreEqual("实际电量（兆瓦时）", ws.Cell(2, mayStart + 1).GetString());
                    Assert.AreEqual("尖", ws.Cell(3, mayStart + 1).GetString());
                    Assert.AreEqual("峰", ws.Cell(3, mayStart + 2).GetString());
                    Assert.AreEqual("平", ws.Cell(3, mayStart + 3).GetString());
                    Assert.AreEqual("谷", ws.Cell(3, mayStart + 4).GetString());

                    var customerA = FindLedgerRow(updated.Worksheet("Sheet1"), "测试客户A");
                    var customerB = FindLedgerRow(updated.Worksheet("Sheet1"), "测试客户B");
                    var customerC = FindLedgerRow(updated.Worksheet("Sheet1"), "台账独有客户");
                    Assert.AreEqual(31.5, ws.Cell(customerA, mayStart).GetDouble(), 0.00001);
                    Assert.AreEqual(3.5, ws.Cell(customerA, mayStart + 1).GetDouble(), 0.00001);
                    Assert.AreEqual(17, ws.Cell(customerA, mayStart + 2).GetDouble(), 0.00001);
                    Assert.AreEqual(5, ws.Cell(customerA, mayStart + 3).GetDouble(), 0.00001);
                    Assert.AreEqual(6, ws.Cell(customerA, mayStart + 4).GetDouble(), 0.00001);
                    Assert.AreEqual(12, ws.Cell(customerB, mayStart).GetDouble(), 0.00001);
                    Assert.IsTrue(ws.Cell(customerC, mayStart).IsEmpty());
                    Assert.IsTrue(ws.Cell(customerC, mayStart + 1).IsEmpty());
                    Assert.IsTrue(ws.Cell(customerC, mayStart + 2).IsEmpty());
                    Assert.IsTrue(ws.Cell(customerC, mayStart + 3).IsEmpty());
                    Assert.IsTrue(ws.Cell(customerC, mayStart + 4).IsEmpty());
                    Assert.AreEqual(1.2, ws.Cell(customerA, mayStart + 5).GetDouble(), 0.00001);
                    Assert.AreEqual(0.8, ws.Cell(customerA, mayStart + 6).GetDouble(), 0.00001);
                    Assert.IsTrue(ws.Cell(customerB, mayStart + 22).IsEmpty());
                    Assert.IsTrue(ws.Cell(customerB, mayStart + 23).IsEmpty());
                    Assert.IsTrue(ws.Cell(customerB, mayStart + 24).IsEmpty());
                    StringAssert.Contains(ws.Cell(customerA, mayStart + 24).FormulaA1, ws.Cell(customerA, mayStart).Address.ToStringRelative());
                    StringAssert.Contains(ws.Cell(customerA, mayStart + 24).FormulaA1, ws.Cell(customerA, mayStart + 22).Address.ToStringRelative());
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static void WriteChongqingWorkbook(string path, bool includeNegative = false, string customerBName = "测试客户B")
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
                WriteRow(ws, row++, customerBName, "B-001", "M4", "尖峰", includeNegative ? -1.0 : 4.0);
                WriteRow(ws, row++, customerBName, "B-001", "M4", "高峰", 8.0);
                ws.Cell(row, 1).Value = "确认人：";

                workbook.SaveAs(path);
            }
        }

        private static void WriteChongqingLedgerWithPreviousMonthOnly(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet("Sheet1");
                ws.Range("A1:J1").Merge();
                ws.Cell("A1").Value = "重庆2026年售电结算台账";
                ws.Cell("A2").Value = "序号";
                ws.Cell("B2").Value = "电力用户编码";
                ws.Cell("C2").Value = "电力用户名称";
                ws.Cell("D2").Value = "合同年用电量（兆瓦时）";
                ws.Cell("E2").Value = "履约开始月份";
                ws.Cell("F2").Value = "履约结束月份";
                ws.Cell("G2").Value = "项目开发人";
                ws.Cell("H2").Value = "代理或自营";
                ws.Cell("I2").Value = "负责人";
                ws.Cell("J2").Value = "员工离职后负责人";
                WriteChongqingMonthBlock(ws, 11, "4月");
                WriteDetailedLedgerRow(ws, 4, 1, "测试客户A", "代理");
                WriteDetailedLedgerRow(ws, 5, 2, "测试客户B", "自营");
                WriteDetailedLedgerRow(ws, 6, 3, "台账独有客户", "代理");
                FillPreviousMonthData(ws, 4, 11, includeSettlementTemplate: true);
                FillPreviousMonthData(ws, 5, 11, includeSettlementTemplate: false);
                FillPreviousMonthData(ws, 6, 11, includeSettlementTemplate: true);
                workbook.SaveAs(path);
            }
        }

        private static void WriteChongqingMonthBlock(IXLWorksheet ws, int startColumn, string monthLabel)
        {
            var headers2 = new[]
            {
                "总实际电量（兆瓦时）",
                "实际电量（兆瓦时）",
                "",
                "",
                "",
                "峰平谷系数",
                "",
                "电量占比(%)",
                "单价（元）",
                "居间收益  （万元）",
                "税点",
                "税费 （万元）",
                "",
                "电量占比(%)",
                "尖峰单价（元）",
                "峰段单价（元）",
                "平段单价（元）",
                "谷段单价（元）",
                "退补收益  （万元）",
                "税点",
                "税费 （万元）",
                "",
                "电量占比(%)",
                "单价（元）",
                "代理收益（万元）",
                "税点",
                "税费  （万元）",
                "",
                "",
                ""
            };
            var headers3 = new[]
            {
                "",
                "尖",
                "峰",
                "平",
                "谷",
                "峰_平",
                "谷_平",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "少回收电能量电费（万元）",
                ""
            };

            ws.Cell(1, startColumn).Value = monthLabel;
            ws.Cell(1, startColumn + 7).Value = "居间";
            ws.Cell(1, startColumn + 10).Value = "扣除费用";
            ws.Cell(1, startColumn + 12).Value = "居间实际收益     （万元）";
            ws.Cell(1, startColumn + 13).Value = "退补电费";
            ws.Cell(1, startColumn + 19).Value = "扣除费用";
            ws.Cell(1, startColumn + 21).Value = "退补电费实际收益     （万元）";
            ws.Cell(1, startColumn + 22).Value = "代理";
            ws.Cell(1, startColumn + 25).Value = "扣除费用";
            ws.Cell(1, startColumn + 27).Value = "代理实际收益      （万元）";
            ws.Cell(1, startColumn + 28).Value = "触发超额电费";
            ws.Cell(1, startColumn + 29).Value = "备注";
            for (var offset = 0; offset < 30; offset++)
            {
                ws.Cell(2, startColumn + offset).Value = headers2[offset];
                ws.Cell(3, startColumn + offset).Value = headers3[offset];
            }
        }

        private static void FillPreviousMonthData(IXLWorksheet ws, int row, int startColumn, bool includeSettlementTemplate)
        {
            ws.Cell(row, startColumn).Value = 99;
            ws.Cell(row, startColumn + 1).Value = 11;
            ws.Cell(row, startColumn + 2).Value = 22;
            ws.Cell(row, startColumn + 3).Value = 33;
            ws.Cell(row, startColumn + 4).Value = 44;
            ws.Cell(row, startColumn + 5).Value = 1.2;
            ws.Cell(row, startColumn + 6).Value = 0.8;
            if (!includeSettlementTemplate)
            {
                return;
            }

            ws.Cell(row, startColumn + 22).Value = 1;
            ws.Cell(row, startColumn + 23).Value = 30;
            ws.Cell(row, startColumn + 24).FormulaA1 = "=ROUND(" + ws.Cell(row, startColumn).Address.ToStringRelative() + "*" + ws.Cell(row, startColumn + 22).Address.ToStringRelative() + "*" + ws.Cell(row, startColumn + 23).Address.ToStringRelative() + "/10,4)";
            ws.Cell(row, startColumn + 25).Value = 0.06;
            ws.Cell(row, startColumn + 26).FormulaA1 = "=ROUND(" + ws.Cell(row, startColumn + 24).Address.ToStringRelative() + "/1.13*" + ws.Cell(row, startColumn + 25).Address.ToStringRelative() + ",4)";
            ws.Cell(row, startColumn + 27).FormulaA1 = "=" + ws.Cell(row, startColumn + 24).Address.ToStringRelative() + "-" + ws.Cell(row, startColumn + 26).Address.ToStringRelative();
        }

        private static void WriteChongqingLedger(string path, string customerBName = "测试客户B")
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
                WriteLedgerRow(ws, 5, 2, customerBName);
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

        private static void WriteDetailedLedgerRow(IXLWorksheet ws, int row, int index, string customerName, string agentOrSelf)
        {
            ws.Cell(row, 1).Value = index;
            ws.Cell(row, 3).Value = customerName;
            ws.Cell(row, 4).Value = 100;
            ws.Cell(row, 5).Value = 202601;
            ws.Cell(row, 6).Value = 202612;
            ws.Cell(row, 7).Value = "测试开发人";
            ws.Cell(row, 8).Value = agentOrSelf;
            ws.Cell(row, 9).Value = "测试负责人";
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

        private static int FindMonthStartColumn(IXLWorksheet worksheet, string monthLabel)
        {
            var lastColumn = worksheet.LastColumnUsed().ColumnNumber();
            for (var column = 1; column <= lastColumn; column++)
            {
                if (worksheet.Cell(1, column).GetString() == monthLabel)
                {
                    return column;
                }
            }

            Assert.Fail("未找到月份区块：" + monthLabel);
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
