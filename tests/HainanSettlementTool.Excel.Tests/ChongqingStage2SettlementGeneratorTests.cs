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
        public void GenerateSettlementClearlyReportsThatChongqingWorkbookGenerationIsNotImplemented()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath);
                WriteSummaryTemplate(options.SummaryTemplatePath);

                var ex = Assert.ThrowsException<NotImplementedException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null));
                StringAssert.Contains(ex.Message, "当前只实现台账读取和预检");
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

        private static void WriteChongqingStage2Ledger(string path)
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
                WriteLedgerRow(ws, 5, "居间客户", null, "代理", "测试负责人", "新增居间", null, 0, 2.2222, 0);
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
            ws.Cell(row, start + 12).Value = intermediaryNet;
            ws.Cell(row, start + 21).Value = refundNet;
            ws.Cell(row, start + 27).Value = proxyNet;
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

        private static void WriteSummaryTemplate(string path, string existingEntity = null, string existingKind = null)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet("汇总表");
                ws.Cell("A1").Value = "重庆代理费汇总表";
                if (!string.IsNullOrWhiteSpace(existingEntity))
                {
                    ws.Cell(4, 1).Value = 1;
                    ws.Cell(4, 2).Value = existingEntity;
                    ws.Cell(4, 3).Value = existingKind;
                    ws.Cell(5, 1).Value = "合计";
                }
                else
                {
                    ws.Cell(4, 1).Value = "合计";
                }

                workbook.SaveAs(path);
            }
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
