using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class ChongqingStage2SettlementGeneratorTests
    {
        [TestMethod]
        public void AnalyzeSettlementReportsPaymentPartyForNewProxyIntermediaryAndRefundSubjects()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath);
                WriteSummaryTemplate(options.SummaryTemplatePath);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.RequiresPaymentPartySelection);
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Proxy, "新增代理");
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Intermediary, "新增居间");
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Refund, "新增退补");
                Assert.IsFalse(report.Issues.Any(issue => issue.Entity == "忽略自营代理"));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementDoesNotRequirePaymentPartyForExistingSummarySubject()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath);
                WriteSummaryTemplate(options.SummaryTemplatePath, "新增代理", ChongqingStage2SettlementKinds.Proxy);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsFalse(report.Issues.Any(issue =>
                    issue.RequiresPaymentPartySelection
                    && issue.SettlementKind == ChongqingStage2SettlementKinds.Proxy
                    && issue.Entity == "新增代理"));
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Intermediary, "新增居间");
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Refund, "新增退补");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementDoesNotRequirePaymentPartyWhenDecisionAlreadyProvided()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                options.SummarySubjectDecisions.Add(new ChongqingStage2SummarySubjectDecision
                {
                    SettlementKind = ChongqingStage2SettlementKinds.Proxy,
                    Entity = "新增代理",
                    PaymentParty = ChongqingStage2PaymentParties.Qingneng
                });
                WriteChongqingStage2Ledger(options.LedgerPath);
                WriteSummaryTemplate(options.SummaryTemplatePath);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsFalse(report.Issues.Any(issue =>
                    issue.RequiresPaymentPartySelection
                    && issue.SettlementKind == ChongqingStage2SettlementKinds.Proxy
                    && issue.Entity == "新增代理"));
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Intermediary, "新增居间");
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Refund, "新增退补");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementWritesChongqingSplitAndSummaryWorkbooks()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                var proxyTemplate = Path.Combine(options.ProxyTemplateDirectory, "测试负责人", "新增代理.xlsx");
                var refundTemplate = Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx");
                WriteProxyTemplate(proxyTemplate);
                WriteRefundTemplate(refundTemplate);
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);

                ChongqingStage2Report report;
                using (File.Open(proxyTemplate, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (File.Open(refundTemplate, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (File.Open(options.SummaryTemplatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null);
                }

                Assert.AreEqual(1, report.ProxyGroups);
                Assert.AreEqual(1, report.RefundGroups);
                Assert.AreEqual(1, report.ProxyRows);
                Assert.AreEqual(1, report.RefundRows);
                Assert.IsTrue(File.Exists(report.Summary));
                Assert.IsTrue(File.Exists(report.ReportPath));
                Assert.IsTrue(Directory.Exists(report.ProxyOutputDirectory));
                Assert.IsTrue(Directory.Exists(report.RefundOutputDirectory));

                var proxyGroup = report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Proxy);
                var refundGroup = report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Refund);
                Assert.IsTrue(File.Exists(proxyGroup.OutputFile));
                Assert.IsTrue(File.Exists(refundGroup.OutputFile));

                using (var workbook = new XLWorkbook(proxyGroup.OutputFile))
                {
                    var ws = workbook.Worksheet("5");
                    Assert.AreEqual("代理名称:新增代理", ws.Cell("A2").GetFormattedString());
                    Assert.AreEqual("代理客户", ws.Cell(5, 2).GetFormattedString());
                    Assert.AreEqual("ROUND(C5*H5*I5/10,4)", ws.Cell(5, 10).FormulaA1);
                    Assert.AreEqual("J5-L5", ws.Cell(5, 11).FormulaA1);
                    Assert.AreEqual(0.3, CellNumber(ws.Cell(5, 12)), 0.0001);
                    Assert.AreEqual("ROUND(K5/1.13*O5,4)", ws.Cell(5, 13).FormulaA1);
                    Assert.AreEqual("K5-M5", ws.Cell(5, 14).FormulaA1);
                }

                using (var workbook = new XLWorkbook(refundGroup.OutputFile))
                {
                    var ws = workbook.Worksheet("5");
                    Assert.AreEqual("名称:新增退补", ws.Cell("A2").GetFormattedString());
                    Assert.AreEqual("退补客户", ws.Cell(5, 2).GetFormattedString());
                    Assert.AreEqual("SUM(D5:G5)", ws.Cell(5, 3).FormulaA1);
                    Assert.AreEqual("ROUND((D5*H5*I5+E5*H5*J5+F5*H5*K5+G5*H5*L5)/10,4)", ws.Cell(5, 13).FormulaA1);
                    Assert.AreEqual("ROUND(M5/1.13*P5,4)", ws.Cell(5, 14).FormulaA1);
                    Assert.AreEqual("M5-N5", ws.Cell(5, 15).FormulaA1);
                }

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var ws = workbook.Worksheet("汇总表");
                    Assert.IsTrue(ws.Row(2).CellsUsed().Any(cell => cell.GetFormattedString().Contains("2026") && cell.GetFormattedString().Contains("5")));
                    Assert.IsTrue(workbook.Worksheets.Any(sheet => sheet.Name == "清能5月"));
                    Assert.IsTrue(workbook.Worksheets.Any(sheet => sheet.Name == "清辉5月"));
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementRefreshesPaymentPartySheetColumnVisibility()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteProxyTemplate(Path.Combine(options.ProxyTemplateDirectory, "测试负责人", "新增代理.xlsx"));
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);
                HidePaymentPartyTemplateColumns(options.SummaryTemplatePath, "清能4月");

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null);

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var ws = workbook.Worksheet("清能5月");
                    AssertColumnsHidden(ws, 12, 17, true);
                    AssertColumnsHidden(ws, 18, 23, false);
                    AssertColumnsHidden(ws, 24, 30, false);
                    Assert.IsTrue(ws.Column(31).IsHidden);
                    Assert.IsFalse(ws.Column(32).IsHidden);
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementPreservesExistingTargetMonthSheets()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                var proxyTemplate = Path.Combine(options.ProxyTemplateDirectory, "测试负责人", "新增代理.xlsx");
                var refundTemplate = Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx");
                WriteProxyTemplate(proxyTemplate);
                WriteRefundTemplate(refundTemplate);
                AddExistingSplitMonthSheet(proxyTemplate, "5", "目标月代理格式");
                AddExistingSplitMonthSheet(refundTemplate, "5", "目标月退补格式");
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);
                AddExistingPaymentPartyMonthSheet(options.SummaryTemplatePath, "清能4月", "清能5月", "目标月清能格式");

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null);

                using (var workbook = new XLWorkbook(report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Proxy).OutputFile))
                {
                    Assert.AreEqual("目标月代理格式", workbook.Worksheet("5").Cell("A3").GetFormattedString());
                }

                using (var workbook = new XLWorkbook(report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Refund).OutputFile))
                {
                    Assert.AreEqual("目标月退补格式", workbook.Worksheet("5").Cell("A3").GetFormattedString());
                }

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    Assert.AreEqual("目标月清能格式", workbook.Worksheet("清能5月").Cell("D1").GetFormattedString());
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementSyncsRefundExtraPowerRowsAndWarns()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteProxyTemplate(Path.Combine(options.ProxyTemplateDirectory, "测试负责人", "新增代理.xlsx"));
                var refundTemplate = Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx");
                WriteRefundTemplate(refundTemplate);
                AddRefundExtraPowerBlock(refundTemplate);
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null);

                Assert.IsTrue(report.Warnings.Any(item => item.Contains("额外扣减块") && item.Contains("已同步 C-G")));
                var refundGroup = report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Refund);
                using (var workbook = new XLWorkbook(refundGroup.OutputFile))
                {
                    var ws = workbook.Worksheet("5");
                    Assert.AreEqual("SUM(D14:G14)", ws.Cell(14, 3).FormulaA1);
                    Assert.AreEqual(10, CellNumber(ws.Cell(14, 4)), 0.0001);
                    Assert.AreEqual(20, CellNumber(ws.Cell(14, 5)), 0.0001);
                    Assert.AreEqual(30, CellNumber(ws.Cell(14, 6)), 0.0001);
                    Assert.AreEqual(40, CellNumber(ws.Cell(14, 7)), 0.0001);
                    Assert.AreEqual(0.9, CellNumber(ws.Cell(14, 8)), 0.0001);
                    Assert.AreEqual("D14+E14", ws.Cell(14, 13).FormulaA1);
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static void AssertHasPaymentIssue(ChongqingStage2PreflightReport report, string kind, string entity)
        {
            Assert.IsTrue(report.Issues.Any(issue =>
                issue.RequiresPaymentPartySelection
                && issue.Kind == ChongqingStage2IssueKinds.NewSummarySubjectPaymentPartyRequired
                && issue.SettlementKind == kind
                && issue.Entity == entity
                && issue.AvailablePaymentParties.Contains(ChongqingStage2PaymentParties.Qingneng)
                && issue.AvailablePaymentParties.Contains(ChongqingStage2PaymentParties.Qinghui)));
        }

        private static ChongqingStage2Options CreateOptions(string root)
        {
            var proxyDir = Path.Combine(root, "proxy");
            var refundDir = Path.Combine(root, "refund");
            Directory.CreateDirectory(proxyDir);
            Directory.CreateDirectory(refundDir);

            return new ChongqingStage2Options
            {
                Month = 5,
                LedgerPath = Path.Combine(root, "ledger.xlsx"),
                ProxyTemplateDirectory = proxyDir,
                RefundTemplateDirectory = refundDir,
                SummaryTemplatePath = Path.Combine(root, "summary.xlsx"),
                OutputDirectory = Path.Combine(root, "out")
            };
        }

        private static void WriteChongqingStage2Ledger(string path, bool includeIntermediary = true)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet("重庆台账");
                ws.Cell("A1").Value = "重庆2026年售电结算台账";
                ws.Cell("A2").Value = "序号";
                ws.Cell("B2").Value = "电力用户编码";
                ws.Cell("C2").Value = "电力用户名称";
                ws.Cell("G2").Value = "项目开发人";
                ws.Cell("H2").Value = "代理或自营";
                ws.Cell("I2").Value = "负责人";
                ws.Cell("J2").Value = "居间人";
                ws.Cell("K2").Value = "收款人";
                WriteChongqingMonthBlock(ws, 12, "5月");
                WriteLedgerRow(ws, 4, "代理客户", "新增代理", "代理", "测试负责人", null, null, 10.1234, 0, 0);
                if (includeIntermediary)
                {
                    WriteLedgerRow(ws, 5, "居间客户", null, "代理", "测试负责人", "新增居间", null, 0, 2.2222, 0);
                }
                WriteLedgerRow(ws, 6, "退补客户", null, "自营", "测试负责人", null, "新增退补", 0, 0, 3.3333);
                WriteLedgerRow(ws, 7, "自营客户", "忽略自营代理", "自营", "测试负责人", null, null, 9.9999, 0, 0);
                workbook.SaveAs(path);
            }
        }

        private static void WriteLedgerRow(
            IXLWorksheet ws,
            int row,
            string customer,
            string proxyEntity,
            string agentOrSelf,
            string owner,
            string intermediaryEntity,
            string refundEntity,
            double proxyNet,
            double intermediaryNet,
            double refundNet)
        {
            var start = 12;
            ws.Cell(row, 1).Value = row - 3;
            ws.Cell(row, 3).Value = customer;
            ws.Cell(row, 7).Value = proxyEntity;
            ws.Cell(row, 8).Value = agentOrSelf;
            ws.Cell(row, 9).Value = owner;
            ws.Cell(row, 10).Value = intermediaryEntity;
            ws.Cell(row, 11).Value = refundEntity;
            ws.Cell(row, start).Value = 100;
            ws.Cell(row, start + 1).Value = 10;
            ws.Cell(row, start + 2).Value = 20;
            ws.Cell(row, start + 3).Value = 30;
            ws.Cell(row, start + 4).Value = 40;
            ws.Cell(row, start + 7).Value = 0.5;
            ws.Cell(row, start + 8).Value = 0.7;
            ws.Cell(row, start + 10).Value = 0.13;
            ws.Cell(row, start + 12).Value = intermediaryNet;
            ws.Cell(row, start + 13).Value = 0.5;
            ws.Cell(row, start + 14).Value = 0.1;
            ws.Cell(row, start + 15).Value = 0.2;
            ws.Cell(row, start + 16).Value = 0.3;
            ws.Cell(row, start + 17).Value = 0.4;
            ws.Cell(row, start + 19).Value = 0.13;
            ws.Cell(row, start + 21).Value = refundNet;
            ws.Cell(row, start + 22).Value = 0.5;
            ws.Cell(row, start + 23).Value = 0.8;
            ws.Cell(row, start + 25).Value = 0.13;
            ws.Cell(row, start + 27).Value = proxyNet;
            ws.Cell(row, start + 28).Value = customer == "代理客户" ? 0.3 : 0;
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
            for (var offset = 0; offset < 30; offset++)
            {
                ws.Cell(2, startColumn + offset).Value = headers2[offset];
                ws.Cell(3, startColumn + offset).Value = headers3[offset];
            }
        }

        private static void WriteProxyTemplate(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet("4");
                ws.Cell("A1").Value = "代理费用结算单";
                ws.Cell("A2").Value = "代理名称:新增代理";
                ws.Cell("E2").Value = "所属期：2026 年 04 月";
                ws.Cell("K2").Value = "结算日期：2026 年 05 月 15 日";
                WriteSplitHeader(ws, 15);
                ws.Cell(5, 1).Value = 1;
                ws.Cell(5, 2).Value = "上月客户";
                ws.Cell(6, 1).Value = "合计";
                for (var column = 3; column <= 14; column++)
                {
                    ws.Cell(6, column).FormulaA1 = "SUM(" + ColumnLetter(column) + "5:" + ColumnLetter(column) + "5)";
                }

                workbook.SaveAs(path);
            }
        }

        private static void WriteRefundTemplate(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet("4");
                ws.Cell("A1").Value = "退补电费结算单";
                ws.Cell("A2").Value = "名称:新增退补";
                ws.Cell("E2").Value = "所属期：2026 年 04 月";
                ws.Cell("L2").Value = "结算日期：2026 年 05 月 15 日";
                WriteSplitHeader(ws, 16);
                ws.Cell(5, 1).Value = 1;
                ws.Cell(5, 2).Value = "上月客户";
                ws.Cell(6, 1).Value = "合计";
                for (var column = 3; column <= 15; column++)
                {
                    ws.Cell(6, column).FormulaA1 = "SUM(" + ColumnLetter(column) + "5:" + ColumnLetter(column) + "5)";
                }

                workbook.SaveAs(path);
            }
        }

        private static void AddExistingSplitMonthSheet(string path, string sheetName, string marker)
        {
            using (var workbook = new XLWorkbook(path))
            {
                workbook.Worksheet("4").CopyTo(sheetName);
                workbook.Worksheet(sheetName).Cell("A3").Value = marker;
                workbook.Save();
            }
        }

        private static void AddRefundExtraPowerBlock(string path)
        {
            using (var workbook = new XLWorkbook(path))
            {
                var ws = workbook.Worksheet("4");
                ws.Cell(14, 2).Value = "当月应扣电表改造费用";
                ws.Cell(14, 3).Value = 1;
                ws.Cell(14, 4).Value = 2;
                ws.Cell(14, 5).Value = 3;
                ws.Cell(14, 6).Value = 4;
                ws.Cell(14, 7).Value = 5;
                ws.Cell(14, 8).Value = 0.9;
                ws.Cell(14, 13).FormulaA1 = "D14+E14";
                workbook.Save();
            }
        }

        private static void WriteSplitHeader(IXLWorksheet ws, int maxColumn)
        {
            for (var column = 1; column <= maxColumn; column++)
            {
                ws.Cell(4, column).Value = "H" + column;
            }
        }

        private static void WriteSummaryTemplate(
            string path,
            string existingEntity = null,
            string existingKind = null,
            string secondEntity = null,
            string secondKind = null)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                WriteSummarySheet(workbook.AddWorksheet("汇总表"), existingEntity, existingKind, secondEntity, secondKind, null);
                WriteSummarySheet(workbook.AddWorksheet("清能4月"), existingEntity, existingKind, null, null, ChongqingStage2PaymentParties.Qingneng);
                WriteSummarySheet(workbook.AddWorksheet("清辉4月"), secondEntity, secondKind, null, null, ChongqingStage2PaymentParties.Qinghui);
                workbook.SaveAs(path);
            }
        }

        private static void AddExistingPaymentPartyMonthSheet(string path, string sourceSheetName, string targetSheetName, string marker)
        {
            using (var workbook = new XLWorkbook(path))
            {
                workbook.Worksheet(sourceSheetName).CopyTo(targetSheetName);
                workbook.Worksheet(targetSheetName).Cell("D1").Value = marker;
                workbook.Save();
            }
        }

        private static void HidePaymentPartyTemplateColumns(string path, string sheetName)
        {
            using (var workbook = new XLWorkbook(path))
            {
                var ws = workbook.Worksheet(sheetName);
                AssertColumnsHidden(ws, 12, 17, false);
                AssertColumnsHidden(ws, 18, 24, false);
                for (var column = 18; column <= 24; column++)
                {
                    ws.Column(column).Hide();
                }

                ws.Column(25).Hide();
                ws.Column(26).Unhide();
                workbook.Save();
            }
        }

        private static void AssertColumnsHidden(IXLWorksheet worksheet, int firstColumn, int lastColumn, bool expectedHidden)
        {
            for (var column = firstColumn; column <= lastColumn; column++)
            {
                Assert.AreEqual(expectedHidden, worksheet.Column(column).IsHidden, "Column " + column);
            }
        }

        private static void WriteSummarySheet(
            IXLWorksheet ws,
            string existingEntity,
            string existingKind,
            string secondEntity,
            string secondKind,
            string paymentPartyFilter)
        {
            ws.Cell("A1").Value = "重庆2026年代理费汇总表（单位：万元）";
            ws.Cell("A2").Value = "序号";
            ws.Cell("B2").Value = "名称";
            ws.Cell("C2").Value = "类目";
            ws.Cell("F2").Value = "收款人";
            ws.Cell("K2").Value = "负责人";
            WriteSummaryMonthBlock(ws, 12, "2026年4月");
            ws.Cell(2, 18).Value = "当年费用总计";
            ws.Cell(2, 19).Value = "借支";
            ws.Cell(2, 20).Value = "已抵扣借支";
            ws.Cell(2, 21).Value = "借支剩余未抵扣";
            ws.Cell(2, 22).Value = "借支开始抵扣月份";
            ws.Cell(2, 23).Value = "借支还完月份";
            ws.Cell(2, 24).Value = "代理/居间/退补电费新增月份";
            ws.Cell(2, 25).Value = "支付方";
            ws.Cell(2, 26).Value = "备注";

            var row = 4;
            if (!string.IsNullOrWhiteSpace(existingEntity)
                && (string.IsNullOrWhiteSpace(paymentPartyFilter) || paymentPartyFilter == ChongqingStage2PaymentParties.Qingneng))
            {
                WriteSummaryDataRow(ws, row++, existingEntity, existingKind, ChongqingStage2PaymentParties.Qingneng);
            }

            if (!string.IsNullOrWhiteSpace(secondEntity)
                && (string.IsNullOrWhiteSpace(paymentPartyFilter) || paymentPartyFilter == ChongqingStage2PaymentParties.Qinghui))
            {
                WriteSummaryDataRow(ws, row++, secondEntity, secondKind, ChongqingStage2PaymentParties.Qinghui);
            }

            ws.Cell(row, 1).Value = "合计";
            for (var column = 12; column <= 21; column++)
            {
                ws.Cell(row, column).FormulaA1 = "SUM(" + ColumnLetter(column) + "4:" + ColumnLetter(column) + (row - 1) + ")";
            }
        }

        private static void WriteSummaryDataRow(IXLWorksheet ws, int row, string entity, string kind, string paymentParty)
        {
            ws.Cell(row, 1).Value = row - 3;
            ws.Cell(row, 2).Value = entity;
            ws.Cell(row, 3).Value = kind;
            ws.Cell(row, 6).Value = entity;
            ws.Cell(row, 11).Value = "测试负责人";
            ws.Cell(row, 15).FormulaA1 = "L" + row + "+M" + row + "+N" + row;
            ws.Cell(row, 17).FormulaA1 = "O" + row + "-P" + row;
            ws.Cell(row, 18).FormulaA1 = "O" + row;
            ws.Cell(row, 20).FormulaA1 = "P" + row;
            ws.Cell(row, 21).FormulaA1 = "S" + row + "-T" + row;
            ws.Cell(row, 25).Value = paymentParty;
        }

        private static void WriteSummaryMonthBlock(IXLWorksheet ws, int startColumn, string label)
        {
            ws.Cell(2, startColumn).Value = label;
            ws.Cell(2, startColumn + 5).Value = "当月实际支付";
            ws.Cell(3, startColumn).Value = "代理费";
            ws.Cell(3, startColumn + 1).Value = "居间费";
            ws.Cell(3, startColumn + 2).Value = "退补电费";
            ws.Cell(3, startColumn + 3).Value = "费用合计";
            ws.Cell(3, startColumn + 4).Value = "当月抵扣";
        }

        private static string CreateTempRoot()
        {
            return Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
        }

        private static double CellNumber(IXLCell cell)
        {
            return cell.DataType == XLDataType.Number ? cell.GetDouble() : Convert.ToDouble(cell.GetFormattedString());
        }

        private static string ColumnLetter(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;
            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
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
