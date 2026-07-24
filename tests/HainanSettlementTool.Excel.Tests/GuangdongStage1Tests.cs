using System;
using System.Collections.Generic;
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
    public sealed class GuangdongStage1Tests
    {
        [TestMethod]
        public void CleanPowerDataAggregatesByCodeUsesFirstCoefficientPairAndIgnoresExtraSheets()
        {
            var root = CreateTempRoot();
            try
            {
                var input = Path.Combine(root, "2026-06零售结算明细.xlsx");
                var outputDirectory = Path.Combine(root, "out");
                WriteGuangdongSource(
                    input,
                    new SourceRow("0001", "广东\u200b测试\n客户Ａ", 10m, 4m, 3m, 3m, 1.2m, 0.8m),
                    new SourceRow("0001", "广东测试客户A", 5m, 2m, 1m, 2m, 1.3m, 0.7m),
                    new SourceRow("0002", "广东客户B", 6.00001m, 2m, 2m, 2.00001m, 1.1m, 0.9m));
                var originalBytes = File.ReadAllBytes(input);

                var result = new ClosedXmlSettlementExcelGateway().CleanPowerData(
                    new ProvinceStage1CleanOptions
                    {
                        Province = ProvinceCode.Guangdong,
                        Month = 6,
                        RawDetailPath = input,
                        OutputDirectory = outputDirectory
                    });

                Assert.AreEqual(ProvinceCode.Guangdong, result.Province);
                Assert.AreEqual(6, result.Month);
                Assert.AreEqual("兆瓦时", result.Unit);
                Assert.AreEqual("零售结算明细", result.SourceSheetName);
                Assert.AreEqual(3, result.RawRows);
                Assert.AreEqual(2, result.CustomerRows);
                Assert.AreEqual(21.00001d, result.TotalPower, 0.000001d);
                Assert.IsTrue(result.Warnings.Any(item => item.Contains("首个有效系数")));
                Assert.IsTrue(File.Exists(result.HtmlReportPath));
                var cleanHtml = File.ReadAllText(result.HtmlReportPath);
                StringAssert.Contains(cleanHtml, "下一步");
                StringAssert.Contains(cleanHtml, "提醒事项");
                StringAssert.Contains(cleanHtml, "程序记录（一般不用打开）");
                Assert.IsFalse(cleanHtml.Contains("电量守恒"));
                CollectionAssert.AreEqual(originalBytes, File.ReadAllBytes(input));

                using (var workbook = new XLWorkbook(result.OutputWorkbookPath))
                {
                    var worksheet = workbook.Worksheet("用户电量汇总");
                    CollectionAssert.AreEqual(
                        new[]
                        {
                            "用户名称",
                            "用户编号",
                            "总实际用电量",
                            "峰电量",
                            "平电量",
                            "谷电量",
                            "峰_平",
                            "谷_平"
                        },
                        Enumerable.Range(1, 8).Select(column => worksheet.Cell(1, column).GetString()).ToArray());
                    Assert.AreEqual("0001", worksheet.Cell(2, 2).GetString());
                    Assert.AreEqual(15m, worksheet.Cell(2, 3).GetValue<decimal>());
                    Assert.AreEqual(6m, worksheet.Cell(2, 4).GetValue<decimal>());
                    Assert.AreEqual(4m, worksheet.Cell(2, 5).GetValue<decimal>());
                    Assert.AreEqual(5m, worksheet.Cell(2, 6).GetValue<decimal>());
                    Assert.AreEqual(1.2m, worksheet.Cell(2, 7).GetValue<decimal>());
                    Assert.AreEqual(0.8m, worksheet.Cell(2, 8).GetValue<decimal>());
                }

                var report = JObject.Parse(File.ReadAllText(result.ReportPath));
                Assert.AreEqual("广东", (string)report["province"]);
                Assert.AreEqual(1, (int)report["coefficientConflictCustomers"]);
                Assert.AreEqual(21.00001m, (decimal)report["totalPower"]);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void CleanPowerDataStopsWhenPowerReconciliationExceedsTolerance()
        {
            var root = CreateTempRoot();
            try
            {
                var input = Path.Combine(root, "2026-06零售结算明细.xlsx");
                WriteGuangdongSource(
                    input,
                    new SourceRow("G001", "误差客户", 10.00002m, 4m, 3m, 3m, 1.2m, 0.8m));

                var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                    new ClosedXmlSettlementExcelGateway().CleanPowerData(
                        new ProvinceStage1CleanOptions
                        {
                            Province = ProvinceCode.Guangdong,
                            Month = 6,
                            RawDetailPath = input,
                            OutputDirectory = Path.Combine(root, "out")
                        }));

                StringAssert.Contains(exception.Message, "严重数据错误");
                StringAssert.Contains(exception.Message, "总实际用电量与峰/平/谷合计不一致");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void CleanPowerDataDoesNotRequireSettlementStaffAuxiliarySheets()
        {
            var root = CreateTempRoot();
            try
            {
                var input = Path.Combine(root, "2026-06零售结算明细.xlsx");
                WriteGuangdongSourceCore(
                    input,
                    false,
                    new SourceRow("G001", "正常下载客户", 10m, 4m, 3m, 3m, 1.2m, 0.8m));

                var result = new ClosedXmlSettlementExcelGateway().CleanPowerData(
                    new ProvinceStage1CleanOptions
                    {
                        Province = ProvinceCode.Guangdong,
                        Month = 6,
                        RawDetailPath = input,
                        OutputDirectory = Path.Combine(root, "out")
                    });

                Assert.AreEqual(1, result.CustomerRows);
                using (var source = new XLWorkbook(input))
                {
                    CollectionAssert.AreEquivalent(
                        new[] { "零售结算明细", "市场联动价格", "零售合同模式" },
                        source.Worksheets.Select(worksheet => worksheet.Name).ToArray());
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void CleanPowerDataStopsWhenOneCodeHasSubstantivelyDifferentNames()
        {
            var root = CreateTempRoot();
            try
            {
                var input = Path.Combine(root, "2026-06零售结算明细.xlsx");
                WriteGuangdongSource(
                    input,
                    new SourceRow("G001", "客户甲", 10m, 4m, 3m, 3m, 1.2m, 0.8m),
                    new SourceRow("G001", "客户乙", 5m, 2m, 1m, 2m, 1.2m, 0.8m));

                var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                    new ClosedXmlSettlementExcelGateway().CleanPowerData(
                        new ProvinceStage1CleanOptions
                        {
                            Province = ProvinceCode.Guangdong,
                            Month = 6,
                            RawDetailPath = input,
                            OutputDirectory = Path.Combine(root, "out")
                        }));

                StringAssert.Contains(exception.Message, "对应不同用户名称");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void UpdateExistingMonthClearsAllPowerPreservesLedgerOnlyCoefficientsAndAppendsSafeCustomer()
        {
            var root = CreateTempRoot();
            try
            {
                var raw = Path.Combine(root, "2026-06零售结算明细.xlsx");
                var ledger = Path.Combine(root, "广东2026年售电结算台账.xlsx");
                var outputDirectory = Path.Combine(root, "out");
                WriteGuangdongSource(
                    raw,
                    new SourceRow("G001", "来源客户一", 10m, 4m, 3m, 3m, 1.2m, 0.8m),
                    new SourceRow("G001", "来源客户一", 5m, 2m, 1m, 2m, 1.3m, 0.7m),
                    new SourceRow("G002", "新增客户二", 6m, 2m, 2m, 2m, 1.1m, 0.9m));
                WriteGuangdongLedger(ledger, 6);
                var rawBytes = File.ReadAllBytes(raw);
                var ledgerBytes = File.ReadAllBytes(ledger);
                var gateway = new ClosedXmlSettlementExcelGateway();
                var options = new ProvinceStage1LedgerUpdateOptions
                {
                    Province = ProvinceCode.Guangdong,
                    Month = 6,
                    LedgerPath = ledger,
                    RawDetailPath = raw,
                    OutputDirectory = outputDirectory
                };

                var plan = gateway.PlanLedgerUpdate(options);

                Assert.IsTrue(plan.RequiresConfirmation);
                Assert.AreEqual(2, plan.LedgerCustomerRows);
                Assert.AreEqual(2, plan.PowerCustomerRows);
                Assert.AreEqual(1, plan.MatchedRows);
                Assert.AreEqual(1, plan.CreatedCustomerRows);
                Assert.AreEqual(1, plan.MissingInPowerRows);
                Assert.AreEqual(0, plan.PowerOnlyCustomers.Count);
                Assert.IsTrue(plan.Issues.Any(issue => issue.Kind == ProvinceStage1LedgerUpdateIssueKinds.CustomerNameMismatch));
                Assert.IsTrue(plan.Issues.Any(issue => issue.Kind == ProvinceStage1LedgerUpdateIssueKinds.CreatedCustomer));
                Assert.IsTrue(plan.Issues.Any(issue => issue.Kind == ProvinceStage1LedgerUpdateIssueKinds.CoefficientConflict));

                var result = gateway.UpdateLedger(options);

                Assert.AreEqual(2, result.UpdatedPowerRows);
                Assert.AreEqual(1, result.CreatedCustomerRows);
                Assert.IsTrue(File.Exists(result.OutputPowerWorkbookPath));
                Assert.IsTrue(File.Exists(result.HtmlReportPath));
                CollectionAssert.AreEqual(rawBytes, File.ReadAllBytes(raw));
                CollectionAssert.AreEqual(ledgerBytes, File.ReadAllBytes(ledger));
                using (var workbook = new XLWorkbook(result.OutputLedgerPath))
                {
                    var worksheet = workbook.Worksheet("广东2026年结算台账");
                    var start = FindMonthStartColumn(worksheet, "6月");
                    var existingRow = FindLedgerRowByCode(worksheet, "G001");
                    var ledgerOnlyRow = FindLedgerRowByCode(worksheet, "G003");
                    var newRow = FindLedgerRowByCode(worksheet, "G002");

                    Assert.AreEqual("台账客户一", worksheet.Cell(existingRow, 3).GetString());
                    Assert.AreEqual(15m, worksheet.Cell(existingRow, start).GetValue<decimal>());
                    Assert.AreEqual(6m, worksheet.Cell(existingRow, start + 1).GetValue<decimal>());
                    Assert.AreEqual(4m, worksheet.Cell(existingRow, start + 2).GetValue<decimal>());
                    Assert.AreEqual(5m, worksheet.Cell(existingRow, start + 3).GetValue<decimal>());
                    Assert.AreEqual(1.2m, worksheet.Cell(existingRow, start + 4).GetValue<decimal>());
                    Assert.AreEqual(0.8m, worksheet.Cell(existingRow, start + 5).GetValue<decimal>());

                    for (var offset = 0; offset < 4; offset++)
                    {
                        Assert.IsFalse(worksheet.Cell(ledgerOnlyRow, start + offset).IsEmpty());
                        Assert.AreEqual(0m, worksheet.Cell(ledgerOnlyRow, start + offset).GetValue<decimal>());
                    }

                    Assert.AreEqual(1.4m, worksheet.Cell(ledgerOnlyRow, start + 4).GetValue<decimal>());
                    Assert.AreEqual(0.6m, worksheet.Cell(ledgerOnlyRow, start + 5).GetValue<decimal>());
                    Assert.AreEqual(3, worksheet.Cell(newRow, 1).GetValue<int>());
                    Assert.AreEqual("新增客户二", worksheet.Cell(newRow, 3).GetString());
                    Assert.AreEqual(202606, worksheet.Cell(newRow, 5).GetValue<int>());
                    Assert.IsTrue(worksheet.Cell(newRow, 4).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(newRow, 7).IsEmpty());
                    Assert.AreEqual(6m, worksheet.Cell(newRow, start).GetValue<decimal>());
                    Assert.AreEqual(1.1m, worksheet.Cell(newRow, start + 4).GetValue<decimal>());
                    Assert.AreEqual(0.9m, worksheet.Cell(newRow, start + 5).GetValue<decimal>());
                    Assert.IsTrue(worksheet.Cell(newRow, start + 20).HasFormula);
                }

                using (var cleanWorkbook = new XLWorkbook(result.OutputPowerWorkbookPath))
                {
                    var worksheet = cleanWorkbook.Worksheet("用户电量汇总");
                    Assert.AreEqual("用户名称", worksheet.Cell(1, 1).GetString());
                    Assert.AreEqual("用户编号", worksheet.Cell(1, 2).GetString());
                    Assert.AreEqual(2, worksheet.LastRowUsed().RowNumber() - 1);
                }

                var html = File.ReadAllText(result.HtmlReportPath);
                StringAssert.Contains(html, "下一步");
                StringAssert.Contains(html, "重点检查");
                StringAssert.Contains(html, "写入 0");
                StringAssert.Contains(html, "程序记录（一般不用打开）");
                StringAssert.Contains(html, "多计量点峰平谷系数不同客户");
                StringAssert.Contains(html, "不影响代理费结算，但建议检查台账。");

                var json = JObject.Parse(File.ReadAllText(result.ReportPath));
                Assert.AreEqual(result.OutputPowerWorkbookPath, (string)json["outputPowerWorkbookPath"]);
                Assert.AreEqual(result.HtmlReportPath, (string)json["htmlReportPath"]);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void UpdateMissingMonthCopiesActualThirtyTwoColumnBlockAndIsRepeatable()
        {
            var root = CreateTempRoot();
            try
            {
                var raw = Path.Combine(root, "2026-06零售结算明细.xlsx");
                var ledger = Path.Combine(root, "广东2026年售电结算台账-5月.xlsx");
                var outputDirectory = Path.Combine(root, "out");
                WriteGuangdongSource(
                    raw,
                    new SourceRow("G001", "台账客户一", 10m, 4m, 3m, 3m, 1.2m, 0.8m));
                WriteGuangdongLedger(ledger, 5);
                var gateway = new ClosedXmlSettlementExcelGateway();
                var options = new ProvinceStage1LedgerUpdateOptions
                {
                    Province = ProvinceCode.Guangdong,
                    Month = 6,
                    LedgerPath = ledger,
                    RawDetailPath = raw,
                    OutputDirectory = outputDirectory
                };

                var plan = gateway.PlanLedgerUpdate(options);
                Assert.IsTrue(plan.Warnings.Any(item => item.Contains("将基于5月的32列结构创建")));
                var first = gateway.UpdateLedger(options);

                using (var workbook = new XLWorkbook(first.OutputLedgerPath))
                {
                    var worksheet = workbook.Worksheet("广东2026年结算台账");
                    var mayStart = FindMonthStartColumn(worksheet, "5月");
                    var juneStart = FindMonthStartColumn(worksheet, "6月");
                    var customer = FindLedgerRowByCode(worksheet, "G001");
                    var ledgerOnly = FindLedgerRowByCode(worksheet, "G003");
                    Assert.AreEqual(mayStart + 32, juneStart);
                    Assert.AreEqual(10m, worksheet.Cell(customer, juneStart).GetValue<decimal>());
                    Assert.AreEqual(0m, worksheet.Cell(ledgerOnly, juneStart).GetValue<decimal>());
                    Assert.AreEqual(0m, worksheet.Cell(ledgerOnly, juneStart + 1).GetValue<decimal>());
                    Assert.AreEqual(0m, worksheet.Cell(ledgerOnly, juneStart + 2).GetValue<decimal>());
                    Assert.AreEqual(0m, worksheet.Cell(ledgerOnly, juneStart + 3).GetValue<decimal>());
                    Assert.IsFalse(worksheet.Cell(ledgerOnly, juneStart).IsEmpty());
                    Assert.IsFalse(worksheet.Cell(ledgerOnly, juneStart + 1).IsEmpty());
                    Assert.IsFalse(worksheet.Cell(ledgerOnly, juneStart + 2).IsEmpty());
                    Assert.IsFalse(worksheet.Cell(ledgerOnly, juneStart + 3).IsEmpty());
                    Assert.AreEqual(1.4m, worksheet.Cell(ledgerOnly, juneStart + 4).GetValue<decimal>());
                    Assert.AreEqual(0.6m, worksheet.Cell(ledgerOnly, juneStart + 5).GetValue<decimal>());
                    Assert.IsTrue(worksheet.Column(juneStart + 31).IsHidden);
                    Assert.IsTrue(worksheet.Cell(customer, juneStart + 20).HasFormula);
                    Assert.IsTrue(IsExactMergedRange(worksheet, 1, juneStart, 1, juneStart + 5));
                }

                options.LedgerPath = first.OutputLedgerPath;
                var second = gateway.UpdateLedger(options);
                using (var workbook = new XLWorkbook(second.OutputLedgerPath))
                {
                    var worksheet = workbook.Worksheet("广东2026年结算台账");
                    Assert.AreEqual(
                        1,
                        Enumerable.Range(1, worksheet.LastColumnUsed().ColumnNumber())
                            .Count(column => worksheet.Cell(1, column).GetString() == "6月"));
                    Assert.AreEqual(1, CountLedgerRowsByCode(worksheet, "G001"));
                    Assert.AreEqual(1, CountLedgerRowsByCode(worksheet, "G003"));
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static void WriteGuangdongSource(string path, params SourceRow[] rows)
        {
            WriteGuangdongSourceCore(path, true, rows);
        }

        private static void WriteGuangdongSourceCore(
            string path,
            bool includeManualSheets,
            params SourceRow[] rows)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("零售结算明细");
                var headers = new[]
                {
                    "无关字段",
                    "用户名称",
                    "谷_平",
                    "用户编号",
                    "峰电量",
                    "总实际用电量",
                    "平电量",
                    "谷电量",
                    "峰_平",
                    "可选费用字段"
                };
                for (var column = 1; column <= headers.Length; column++)
                {
                    worksheet.Cell(1, column).Value = headers[column - 1];
                }

                for (var index = 0; index < rows.Length; index++)
                {
                    var row = rows[index];
                    var excelRow = index + 2;
                    worksheet.Cell(excelRow, 1).Value = "忽略";
                    worksheet.Cell(excelRow, 2).Value = row.CustomerName;
                    worksheet.Cell(excelRow, 3).SetValue(row.ValleyFlatCoefficient);
                    worksheet.Cell(excelRow, 4).Value = row.Code;
                    worksheet.Cell(excelRow, 4).Style.NumberFormat.Format = "@";
                    worksheet.Cell(excelRow, 5).SetValue(row.Peak);
                    worksheet.Cell(excelRow, 6).SetValue(row.Total);
                    worksheet.Cell(excelRow, 7).SetValue(row.Flat);
                    worksheet.Cell(excelRow, 8).SetValue(row.Valley);
                    worksheet.Cell(excelRow, 9).SetValue(row.PeakFlatCoefficient);
                    worksheet.Cell(excelRow, 10).Value = 999;
                }

                workbook.AddWorksheet("市场联动价格").Cell("A1").Value = "阶段一不读取";
                workbook.AddWorksheet("零售合同模式").Cell("A1").Value = "阶段一不读取";
                if (includeManualSheets)
                {
                    workbook.AddWorksheet("人工辅助-任意名称").Cell("A1").Value = "错误数据也必须忽略";
                    workbook.AddWorksheet("6月电量 (2)").Cell("A1").Value = "不是输入合同";
                }

                workbook.SaveAs(path);
            }
        }

        private static void WriteGuangdongLedger(string path, int month)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("广东2026年结算台账");
                worksheet.Cell(1, 1).Value = "广东2026年售电结算台账";
                worksheet.Cell(2, 1).Value = "序号";
                worksheet.Cell(2, 2).Value = "电力用户编码";
                worksheet.Cell(2, 3).Value = "电力用户名称";
                worksheet.Cell(2, 4).Value = "合同年用电量（兆瓦时）";
                worksheet.Cell(2, 5).Value = "履约开始月份";
                worksheet.Cell(2, 6).Value = "履约结束月份";
                worksheet.Cell(2, 7).Value = "手工业务资料";
                worksheet.Cell(2, 8).Value = "代理或自营";
                worksheet.Cell(2, 9).Value = "负责人";

                const int monthStart = 11;
                WriteGuangdongMonthBlock(worksheet, monthStart, month + "月");
                WriteLedgerCustomer(worksheet, 4, 1, "G001", "台账客户一");
                WriteLedgerCustomer(worksheet, 5, 2, "G003", "台账独有客户");
                worksheet.Cell(5, 7).Value = "不得继承的业务值";
                worksheet.Cell(5, 8).Value = "代理";
                worksheet.Cell(5, 9).Value = "负责人";
                worksheet.Cell(4, monthStart).Value = 88;
                worksheet.Cell(4, monthStart + 1).Value = 8;
                worksheet.Cell(4, monthStart + 2).Value = 30;
                worksheet.Cell(4, monthStart + 3).Value = 50;
                worksheet.Cell(4, monthStart + 4).Value = 9.9;
                worksheet.Cell(4, monthStart + 5).Value = 8.8;
                worksheet.Cell(5, monthStart).Value = 77;
                worksheet.Cell(5, monthStart + 1).Value = 7;
                worksheet.Cell(5, monthStart + 2).Value = 20;
                worksheet.Cell(5, monthStart + 3).Value = 50;
                worksheet.Cell(5, monthStart + 4).Value = 1.4;
                worksheet.Cell(5, monthStart + 5).Value = 0.6;
                worksheet.Cell(4, monthStart + 20).FormulaA1 = "=ROUND(" + worksheet.Cell(4, monthStart).Address.ToStringRelative() + "*2,4)";
                worksheet.Cell(5, monthStart + 20).FormulaA1 = "=ROUND(" + worksheet.Cell(5, monthStart).Address.ToStringRelative() + "*2,4)";
                worksheet.Column(monthStart + 31).Hide();
                worksheet.Cell(6, 1).Value = "合计";
                worksheet.Cell(6, 3).Value = "footer";
                workbook.SaveAs(path);
            }
        }

        private static void WriteGuangdongMonthBlock(IXLWorksheet worksheet, int start, string label)
        {
            worksheet.Range(1, start, 1, start + 5).Merge();
            worksheet.Cell(1, start).Value = label;
            worksheet.Cell(2, start).Value = "总实际用电量";
            worksheet.Cell(2, start + 1).Value = "实际用电量（兆瓦时）";
            worksheet.Cell(3, start + 1).Value = "峰";
            worksheet.Cell(3, start + 2).Value = "平";
            worksheet.Cell(3, start + 3).Value = "谷";
            worksheet.Cell(2, start + 4).Value = "峰平谷系数";
            worksheet.Cell(3, start + 4).Value = "峰_平";
            worksheet.Cell(3, start + 5).Value = "谷_平";
            for (var offset = 6; offset < 32; offset++)
            {
                worksheet.Cell(2, start + offset).Value = "月度字段" + offset;
            }
        }

        private static void WriteLedgerCustomer(
            IXLWorksheet worksheet,
            int row,
            int sequence,
            string code,
            string customerName)
        {
            worksheet.Cell(row, 1).Value = sequence;
            worksheet.Cell(row, 2).Value = code;
            worksheet.Cell(row, 2).Style.NumberFormat.Format = "@";
            worksheet.Cell(row, 3).Value = customerName;
            worksheet.Cell(row, 4).Value = 100;
            worksheet.Cell(row, 5).Value = 202601;
            worksheet.Cell(row, 6).Value = 202612;
        }

        private static int FindMonthStartColumn(IXLWorksheet worksheet, string label)
        {
            var lastColumn = worksheet.LastColumnUsed().ColumnNumber();
            for (var column = 1; column <= lastColumn; column++)
            {
                if (worksheet.Cell(1, column).GetString() == label)
                {
                    return column;
                }
            }

            Assert.Fail("未找到月份区块：" + label);
            return 0;
        }

        private static int FindLedgerRowByCode(IXLWorksheet worksheet, string code)
        {
            var lastRow = worksheet.LastRowUsed().RowNumber();
            for (var row = 4; row <= lastRow; row++)
            {
                if (worksheet.Cell(row, 2).GetString() == code)
                {
                    return row;
                }
            }

            Assert.Fail("未找到台账客户编码：" + code);
            return 0;
        }

        private static int CountLedgerRowsByCode(IXLWorksheet worksheet, string code)
        {
            return Enumerable.Range(4, Math.Max(0, worksheet.LastRowUsed().RowNumber() - 3))
                .Count(row => worksheet.Cell(row, 2).GetString() == code);
        }

        private static bool IsExactMergedRange(
            IXLWorksheet worksheet,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn)
        {
            return worksheet.MergedRanges.Any(range =>
                range.RangeAddress.FirstAddress.RowNumber == firstRow
                && range.RangeAddress.FirstAddress.ColumnNumber == firstColumn
                && range.RangeAddress.LastAddress.RowNumber == lastRow
                && range.RangeAddress.LastAddress.ColumnNumber == lastColumn);
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

        private sealed class SourceRow
        {
            public SourceRow(
                string code,
                string customerName,
                decimal total,
                decimal peak,
                decimal flat,
                decimal valley,
                decimal peakFlatCoefficient,
                decimal valleyFlatCoefficient)
            {
                Code = code;
                CustomerName = customerName;
                Total = total;
                Peak = peak;
                Flat = flat;
                Valley = valley;
                PeakFlatCoefficient = peakFlatCoefficient;
                ValleyFlatCoefficient = valleyFlatCoefficient;
            }

            public string Code { get; }
            public string CustomerName { get; }
            public decimal Total { get; }
            public decimal Peak { get; }
            public decimal Flat { get; }
            public decimal Valley { get; }
            public decimal PeakFlatCoefficient { get; }
            public decimal ValleyFlatCoefficient { get; }
        }
    }
}
