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
    public sealed class HainanStage2SettlementGeneratorTests
    {
        [TestMethod]
        public void AnalyzeSettlementBlocksProxySubjectWithMissingRelationshipParameter()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "测试代理", "存量客户");
                using (var workbook = new XLWorkbook(ledgerPath))
                {
                    var start = HainanLedgerLayout.MonthStartColumn(4);
                    workbook.Worksheet(HainanLedgerLayout.MainSheetName).Cell(4, start + 13).Clear(XLClearOptions.Contents);
                    workbook.Save();
                }

                WriteProxyTemplate(proxyRoot, "测试负责人", "测试代理");
                WriteSummaryTemplate(summaryPath, "测试代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.HasBlockingIssues);
                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.RelationshipParametersInvalid
                    && issue.Kind == "代理费"
                    && issue.SettlementKind == "代理费"
                    && issue.LedgerRow == 4
                    && issue.BlocksGeneration));
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
        public void AnalyzeSettlementBlocksRelationshipParametersWithoutProxySubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "测试代理", "存量客户");
                using (var workbook = new XLWorkbook(ledgerPath))
                {
                    workbook.Worksheet(HainanLedgerLayout.MainSheetName).Cell(4, 8).Clear(XLClearOptions.Contents);
                    workbook.Save();
                }

                WriteSummaryTemplate(summaryPath, "占位主体", "代理费", "清辉");
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.RelationshipParametersWithoutSubject
                    && issue.Kind == "代理费"
                    && issue.LedgerRow == 4
                    && issue.BlocksGeneration));
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
        public void AnalyzeSettlementBlocksNonPositiveIntermediaryRelationshipParameter()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "临时代理", "存量客户");
                using (var workbook = new XLWorkbook(ledgerPath))
                {
                    var worksheet = workbook.Worksheet(HainanLedgerLayout.MainSheetName);
                    var start = HainanLedgerLayout.MonthStartColumn(4);
                    worksheet.Cell(4, 8).Clear(XLClearOptions.Contents);
                    worksheet.Cell(4, start + 13).Clear(XLClearOptions.Contents);
                    worksheet.Cell(4, start + 14).Clear(XLClearOptions.Contents);
                    worksheet.Cell(4, start + 16).Clear(XLClearOptions.Contents);
                    worksheet.Cell(4, 19).Value = "测试居间人";
                    worksheet.Cell(4, start + 7).Value = 0.5;
                    worksheet.Cell(4, start + 8).Value = 1.2;
                    worksheet.Cell(4, start + 10).Value = 0;
                    workbook.Save();
                }

                WriteSummaryTemplate(summaryPath, "占位主体", "居间费", HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);
                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.RelationshipParametersInvalid
                    && issue.Kind == "居间费"
                    && issue.Entity == "测试居间人"
                    && issue.LedgerRow == 4
                    && issue.BlocksGeneration));
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
        public void AnalyzeSettlementBlocksConflictingTaxRatesForSameProxySubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyRows(ledgerPath, "测试负责人", "测试代理", "客户甲", "客户乙");
                using (var workbook = new XLWorkbook(ledgerPath))
                {
                    var start = HainanLedgerLayout.MonthStartColumn(4);
                    workbook.Worksheet(HainanLedgerLayout.MainSheetName).Cell(5, start + 16).Value = 0.060000001;
                    workbook.Save();
                }

                WriteProxyTemplate(proxyRoot, "测试负责人", "测试代理");
                WriteSummaryTemplate(summaryPath, "测试代理", "代理费", "清辉");
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingTaxRates
                    && issue.Kind == "代理费"
                    && issue.Entity == "测试代理"
                    && issue.BlocksGeneration));
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
        public void AnalyzeSettlementIncludesZeroAmountRelationshipInTaxRateConflict()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(Path.GetDirectoryName(ledgerPath));
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                    var start = HainanLedgerLayout.MonthStartColumn(4);
                    worksheet.Cell(1, start).Value = "4月";
                    WriteLedgerRow(worksheet, 4, start, "测试负责人", "税率冲突代理", "零电量客户", 0);
                    worksheet.Cell(4, start + 16).Value = 0.03;
                    WriteLedgerRow(worksheet, 5, start, "测试负责人", "税率冲突代理", "结算客户", 100);
                    workbook.SaveAs(ledgerPath);
                }

                WriteProxyTemplate(proxyRoot, "测试负责人", "税率冲突代理");
                WriteSummaryTemplate(summaryPath, "税率冲突代理", "代理费", HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var preflight = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingTaxRates
                    && issue.Entity == "税率冲突代理"
                    && issue.BlocksGeneration));
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
        public void GenerateSettlementRejectsDirectCallWithoutConfirmedPreflight()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "禁止绕过预检代理");
                var gateway = new ClosedXmlSettlementExcelGateway();

                var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                    gateway.GenerateSettlement(options));

                StringAssert.Contains(exception.Message, "缺少本次预检签名");
                Assert.IsFalse(Directory.Exists(options.OutputDirectory));
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
        public void GenerateSettlementRejectsLedgerChangedAfterConfirmedPreflight()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "输入变化代理");
                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                options.ExpectedPreflightSignature = preflight.PreflightSignature;
                options.ExpectedInputFingerprint = preflight.InputFingerprint;
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    workbook.Worksheet(HainanLedgerLayout.MainSheetName).Cell(4, 2).Value = "预检后变更客户";
                    workbook.Save();
                }

                var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                    service.Run(options, null));

                StringAssert.Contains(exception.Message, "输入已变化");
                Assert.IsFalse(Directory.Exists(options.OutputDirectory));
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
        public void GenerateSettlementRejectsOutputConfigurationChangedAfterConfirmedPreflight()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "配置变化代理");
                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                options.ExpectedPreflightSignature = preflight.PreflightSignature;
                options.ExpectedInputFingerprint = preflight.InputFingerprint;
                options.OutputDirectory = Path.Combine(root, "changed-output");

                var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                    service.Run(options, null));

                StringAssert.Contains(exception.Message, "输入已变化");
                Assert.IsFalse(Directory.Exists(options.OutputDirectory));
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
        public void OutputSummaryNameRejectsAbsoluteAndParentTraversalWithoutTouchingOutsideFiles()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "路径防护代理");
                var absoluteOutside = Path.Combine(root, "absolute-outside.xlsx");
                var traversalOutside = Path.Combine(root, "traversal-outside.xlsx");
                File.WriteAllText(absoluteOutside, "absolute-sentinel");
                File.WriteAllText(traversalOutside, "traversal-sentinel");
                var cases = new[]
                {
                    new { Name = absoluteOutside, Path = absoluteOutside, Sentinel = "absolute-sentinel" },
                    new
                    {
                        Name = ".." + Path.DirectorySeparatorChar + Path.GetFileName(traversalOutside),
                        Path = traversalOutside,
                        Sentinel = "traversal-sentinel"
                    }
                };

                foreach (var item in cases)
                {
                    options.OutputSummaryName = item.Name;
                    var serviceException = Assert.ThrowsException<ArgumentException>(() =>
                        new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options));
                    StringAssert.Contains(serviceException.Message, ".xlsx");
                    var defenseException = Assert.ThrowsException<InvalidOperationException>(() =>
                        HainanStage2SummaryWorkbookWriter.PlanOutputPath(options));
                    StringAssert.Contains(defenseException.Message, "输出根目录");
                    Assert.AreEqual(item.Sentinel, File.ReadAllText(item.Path));
                }

                Assert.IsFalse(Directory.Exists(options.OutputDirectory));
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
        public void GenerateSettlementCombinesMultiOwnerSubjectUnderFirstPhysicalLedgerOwner()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);
                Directory.CreateDirectory(Path.GetDirectoryName(ledgerPath));
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                    var start = HainanLedgerLayout.MonthStartColumn(4);
                    worksheet.Cell(1, start).Value = "4月";
                    WriteLedgerRow(worksheet, 4, start, "负责人甲", "共同代理", "零电量客户", 0);
                    WriteLedgerRow(worksheet, 5, start, "负责人乙", "共同代理", "结算客户", 100);
                    workbook.SaveAs(ledgerPath);
                }

                WriteProxyTemplate(proxyRoot, "负责人乙", "共同代理");
                WriteSummaryTemplate(summaryPath, "共同代理", "代理费", "清辉");
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.MultipleOwners
                    && issue.Entity == "共同代理"
                    && !issue.BlocksGeneration));
                Assert.IsFalse(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingTaxRates
                    && issue.Entity == "共同代理"));

                var report = RunAfterPreflight(service, options);

                Assert.AreEqual(1, report.ProxyGroups);
                Assert.AreEqual("负责人甲", report.Groups.Single().Owner);
                var expected = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "负责人甲 - 海南2026",
                    "共同代理 2026海南.xlsx");
                Assert.IsTrue(File.Exists(expected));
                Assert.AreEqual(
                    1,
                    Directory.GetFiles(Path.Combine(outputRoot, "2026年代理 - 海南"), "*.xlsx", SearchOption.AllDirectories).Length);
                using (var workbook = new XLWorkbook(expected))
                {
                    Assert.AreEqual("结算客户", workbook.Worksheet("4月").Cell(5, 2).GetFormattedString());
                }
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
        public void GenerateSettlementStopsBeforeWritingFilesWhenExactTemplatesAreDuplicated()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "负责人甲", "重复代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "负责人甲", "重复代理");
                WriteProxyTemplate(proxyRoot, "负责人乙", "重复代理");
                WriteSummaryTemplate(summaryPath, "重复代理", "代理费", "清辉");
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.DuplicateExactTemplates
                    && issue.Kind == "代理费"
                    && issue.Entity == "重复代理"
                    && issue.BlocksGeneration));

                var error = Assert.ThrowsException<InvalidOperationException>(() => RunAfterPreflight(service, options));
                StringAssert.Contains(error.Message, "多个同名上月分表");
                Assert.AreEqual(
                    0,
                    Directory.Exists(outputRoot)
                        ? Directory.GetFiles(outputRoot, "*", SearchOption.AllDirectories).Length
                        : 0);
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
        public void AnalyzeSettlementIgnoresAppleDoubleTemplateFilesConsistentlyWithFingerprint()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "AppleDouble代理");
                var ignoredPath = Path.Combine(options.ProxyTemplateDirectory, "._ignored.xlsx");
                File.WriteAllText(ignoredPath, "not-an-excel-workbook");

                var preflight = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsFalse(string.IsNullOrWhiteSpace(preflight.InputFingerprint));
                Assert.IsFalse(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.TemplateUnreadable
                    && string.Equals(issue.TemplateFile, ignoredPath, StringComparison.OrdinalIgnoreCase)));
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        [DataTestMethod]
        [DataRow("")]
        [DataRow("未知支付方")]
        public void ExistingSubjectWithInvalidPaymentPartyRequiresDecisionAndUsesSelection(string initialPaymentParty)
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "支付方待选代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "支付方待选代理");
                WriteSummaryTemplate(summaryPath, "支付方待选代理", "代理费", initialPaymentParty);
                using (var workbook = new XLWorkbook(summaryPath))
                {
                    workbook.Worksheet("清辉汇总表")
                        .Range(4, 1, 4, 30)
                        .Clear(XLClearOptions.Contents);
                    workbook.Save();
                }
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PaymentPartyRequired
                    && issue.Disposition == Stage2PreflightDisposition.RequiredDecision
                    && issue.Kind == "代理费"
                    && issue.Entity == "支付方待选代理"));

                Assert.ThrowsException<InvalidOperationException>(() => RunAfterPreflight(service, options));
                Assert.AreEqual(
                    0,
                    Directory.Exists(outputRoot)
                        ? Directory.GetFiles(outputRoot, "*", SearchOption.AllDirectories).Length
                        : 0);

                options.SummarySubjectDecisions.Add(new HainanStage2SummarySubjectDecision
                {
                    SettlementKind = "代理费",
                    Entity = "支付方待选代理",
                    PaymentParty = HainanStage2PaymentParties.Qingneng
                });
                var generated = RunAfterPreflight(service, options);
                using (var workbook = new XLWorkbook(generated.Summary))
                {
                    var worksheet = workbook.Worksheet("汇总表");
                    Assert.AreEqual("J4-I4", worksheet.Cell(4, 8).FormulaA1);
                    Assert.AreEqual(0.06, worksheet.Cell(4, 9).GetDouble(), 0.0000001);
                    Assert.AreEqual(0.13, worksheet.Cell(4, 10).GetDouble(), 0.0000001);
                    Assert.AreEqual("测试负责人", worksheet.Cell(4, 11).GetFormattedString());
                    Assert.AreEqual(HainanStage2PaymentParties.Qingneng, worksheet.Cell(4, 36).GetFormattedString());
                }
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
        public void SettlementPreservesOpaqueMultiNamePayeeAndIgnoresOnlyBoundaryWhitespace()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");
            const string mainPayee = "  张三、李四、王五\r\n";

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithProxyRows(ledgerPath, "测试负责人", "多人收款代理", "存量客户", "新增客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "多人收款代理");
                WriteSummaryTemplateWithReliableSources(
                    summaryPath,
                    "多人收款代理",
                    "代理费",
                    mainPayee,
                    HainanStage2PaymentParties.Qinghui,
                    "\n张三、李四、王五  ",
                    HainanStage2PaymentParties.Qinghui);
                string storedMainPayee;
                using (var sourceWorkbook = new XLWorkbook(summaryPath))
                {
                    storedMainPayee = sourceWorkbook.Worksheet("汇总表").Cell(4, 6).GetString();
                }

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsFalse(preflight.Issues.Exists(issue => issue.Code == Stage2PreflightIssueKinds.ConflictingPayees));
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.NewCustomer
                    && issue.CurrentValue.Contains("张三、李四、王五")
                    && issue.CurrentValue.Contains(HainanStage2PaymentParties.Qinghui)));

                var report = RunAfterPreflight(service, options);
                using (var workbook = new XLWorkbook(report.Summary))
                {
                    Assert.AreEqual(storedMainPayee, workbook.Worksheet("汇总表").Cell(4, 6).GetString());
                }
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
        public void GenerateSettlementStopsWhenReliablePayeeSourcesConflict()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "收款冲突代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "收款冲突代理");
                WriteSummaryTemplateWithReliableSources(
                    summaryPath,
                    "收款冲突代理",
                    "代理费",
                    "张三",
                    HainanStage2PaymentParties.Qinghui,
                    "李四",
                    HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingPayees
                    && issue.Entity == "收款冲突代理"
                    && issue.BlocksGeneration));

                Assert.ThrowsException<InvalidOperationException>(() => RunAfterPreflight(service, options));
                Assert.IsFalse(Directory.Exists(outputRoot));
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
        public void GenerateSettlementStopsWhenReliablePaymentPartySourcesConflict()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "支付冲突代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "支付冲突代理");
                WriteSummaryTemplateWithReliableSources(
                    summaryPath,
                    "支付冲突代理",
                    "代理费",
                    "张三",
                    HainanStage2PaymentParties.Qingneng,
                    "张三",
                    HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingPaymentParties
                    && issue.Entity == "支付冲突代理"
                    && issue.BlocksGeneration));

                Assert.ThrowsException<InvalidOperationException>(() => RunAfterPreflight(service, options));
                Assert.IsFalse(Directory.Exists(outputRoot));
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        [DataTestMethod]
        [DataRow("清辉")]
        [DataRow("")]
        public void AnalyzeSettlementBlocksSameSubjectInBothPaymentSheetsEvenWhenPartyTextDoesNotConflict(
            string paymentPartyText)
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "双支付方代理");
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    var qingneng = workbook.Worksheet("清能汇总表");
                    var qinghui = workbook.Worksheet("清辉汇总表");
                    qingneng.Row(4).InsertRowsAbove(1);
                    qingneng.Cell(4, 1).Value = 1;
                    qingneng.Cell(4, 2).Value = "双支付方代理";
                    qingneng.Cell(4, 3).Value = "代理费";
                    qingneng.Cell(4, 30).Value = paymentPartyText;
                    main.Cell(4, 30).Value = paymentPartyText;
                    qinghui.Cell(4, 30).Value = paymentPartyText;
                    workbook.Save();
                }

                var preflight = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingPaymentParties
                    && issue.Entity == "双支付方代理"
                    && issue.Category.Contains("同时出现")
                    && issue.Message.Contains("清能和清辉")
                    && issue.BlocksGeneration));
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
        public void SettlementInheritsQingnengFromUniquePaymentSheetMembershipWhenPartyFieldsAreBlank()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "清能所属代理");
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    workbook.Worksheet(HainanLedgerLayout.MainSheetName).Cell(4, 3).Value = "新增客户";
                    workbook.Save();
                }
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    workbook.Worksheet("汇总表").Cell(4, 30).Clear(XLClearOptions.Contents);
                    workbook.Worksheet("清辉汇总表")
                        .Range(4, 1, 4, 30)
                        .Clear(XLClearOptions.Contents);
                    var qingneng = workbook.Worksheet("清能汇总表");
                    qingneng.Row(4).InsertRowsAbove(1);
                    qingneng.Cell(4, 1).Value = 1;
                    qingneng.Cell(4, 2).Value = "清能所属代理";
                    qingneng.Cell(4, 3).Value = "代理费";
                    workbook.Save();
                }

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsFalse(preflight.Issues.Exists(issue =>
                    issue.Entity == "清能所属代理"
                    && (issue.Code == Stage2PreflightIssueKinds.PaymentPartyRequired
                        || issue.Code == Stage2PreflightIssueKinds.ConflictingPaymentParties)));
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.NewCustomer
                    && issue.Entity == "清能所属代理"
                    && issue.CurrentValue.Contains("支付方：清能")
                    && !issue.CurrentValue.Contains("待选择")));

                var report = RunAfterPreflight(service, options);
                using (var workbook = new XLWorkbook(report.Summary))
                {
                    Assert.AreEqual(
                        HainanStage2PaymentParties.Qingneng,
                        workbook.Worksheet("汇总表").Cell(4, 36).GetFormattedString());
                    Assert.IsTrue(HainanStage2SummaryWorkbookWriter
                        .ReadSummaryMeta(workbook.Worksheet("清能汇总表"))
                        .Any(row => row.Entity == "清能所属代理"));
                    Assert.IsFalse(HainanStage2SummaryWorkbookWriter
                        .ReadSummaryMeta(workbook.Worksheet("清辉汇总表"))
                        .Any(row => row.Entity == "清能所属代理"));
                }
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
        public void AnalyzeSettlementBlocksPaymentSheetMembershipConflictingWithMainPartyField()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "归属字段冲突代理");
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    workbook.Worksheet("清辉汇总表")
                        .Range(4, 1, 4, 30)
                        .Clear(XLClearOptions.Contents);
                    var qingneng = workbook.Worksheet("清能汇总表");
                    qingneng.Row(4).InsertRowsAbove(1);
                    qingneng.Cell(4, 1).Value = 1;
                    qingneng.Cell(4, 2).Value = "归属字段冲突代理";
                    qingneng.Cell(4, 3).Value = "代理费";
                    workbook.Save();
                }

                var preflight = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingPaymentParties
                    && issue.Entity == "归属字段冲突代理"
                    && issue.Category.Contains("工作表所属与字段冲突")
                    && issue.CurrentValue.Contains("工作表所属=清能")
                    && issue.CurrentValue.Contains("汇总表=清辉")
                    && issue.BlocksGeneration));
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
        public void SettlementUsesUniqueOpaquePayeeWhenMainSummaryPayeeIsBlank()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");
            const string canonicalPayee = "\r\n张三、李四  ";

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "收款补全代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "收款补全代理");
                WriteSummaryTemplateWithReliableSources(
                    summaryPath,
                    "收款补全代理",
                    "代理费",
                    string.Empty,
                    HainanStage2PaymentParties.Qinghui,
                    canonicalPayee,
                    HainanStage2PaymentParties.Qinghui);
                string storedCanonicalPayee;
                using (var sourceWorkbook = new XLWorkbook(summaryPath))
                {
                    storedCanonicalPayee = sourceWorkbook.Worksheet("清辉汇总表").Cell(4, 6).GetString();
                }

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PayeeSourceMissing
                    && issue.Disposition == Stage2PreflightDisposition.Review
                    && issue.CurrentValue == storedCanonicalPayee));

                var report = RunAfterPreflight(service, options);
                using (var workbook = new XLWorkbook(report.Summary))
                {
                    Assert.AreEqual(storedCanonicalPayee, workbook.Worksheet("汇总表").Cell(4, 6).GetString());
                }
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
        public void SettlementReviewsBlankPaymentSheetPayeeAndUsesUniqueMainPayee()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");
            const string mainPayee = "  张三、李四  ";
            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "反向空白收款代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "反向空白收款代理");
                WriteSummaryTemplateWithReliableSources(
                    summaryPath,
                    "反向空白收款代理",
                    "代理费",
                    mainPayee,
                    HainanStage2PaymentParties.Qinghui,
                    string.Empty,
                    HainanStage2PaymentParties.Qinghui);
                string storedMainPayee;
                using (var workbook = new XLWorkbook(summaryPath))
                {
                    storedMainPayee = workbook.Worksheet("汇总表").Cell(4, 6).GetString();
                }
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };
                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());

                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PayeeSourceMissing
                    && issue.Disposition == Stage2PreflightDisposition.Review
                    && issue.PreviousValue.Contains("清辉汇总表")
                    && issue.CurrentValue == storedMainPayee));

                var report = RunAfterPreflight(service, options);
                using (var workbook = new XLWorkbook(report.Summary))
                {
                    Assert.AreEqual(storedMainPayee, workbook.Worksheet("汇总表").Cell(4, 6).GetString());
                    Assert.AreEqual(storedMainPayee, workbook.Worksheet("清辉汇总表").Cell(4, 6).GetString());
                }
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
        public void SettlementCopiesCanonicalMainLongTermFieldsWhenPaymentSheetLacksSubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "跨表继承代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "跨表继承代理");
                WriteSummaryTemplateWithMissingPaymentSheetSubject(
                    summaryPath,
                    "跨表继承代理",
                    "代理费",
                    "张三、李四");
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = RunAfterPreflight(options);
                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var worksheet = workbook.Worksheet("清能汇总表");
                    Assert.AreEqual("跨表继承代理", worksheet.Cell(4, 2).GetFormattedString());
                    Assert.AreEqual("是", worksheet.Cell(4, 4).GetFormattedString());
                    Assert.AreEqual("张三、李四", worksheet.Cell(4, 6).GetString());
                    Assert.AreEqual("专票", worksheet.Cell(4, 7).GetFormattedString());
                    Assert.AreEqual(HainanStage2PaymentParties.Qingneng, worksheet.Cell(4, 36).GetFormattedString());
                }
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
        public void GenerateSettlementStopsWhenMainSummaryContainsDuplicateSubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "重复汇总代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "重复汇总代理");
                WriteSummaryTemplateWithDuplicateSubject(
                    summaryPath,
                    "重复汇总代理",
                    "代理费",
                    HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.DuplicateSummarySubject
                    && issue.Entity == "重复汇总代理"
                    && issue.BlocksGeneration));

                Assert.ThrowsException<InvalidOperationException>(() => RunAfterPreflight(service, options));
                Assert.IsFalse(Directory.Exists(outputRoot));
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
        public void JingyanProxyPaymentPartyOverrideRemainsQingneng()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "海南精研科技有限公司", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "海南精研科技有限公司");
                WriteSummaryTemplate(
                    summaryPath,
                    "海南精研科技有限公司",
                    "代理费",
                    HainanStage2PaymentParties.Qinghui);
                using (var workbook = new XLWorkbook(summaryPath))
                {
                    var qingneng = workbook.Worksheet("清能汇总表");
                    qingneng.Row(4).InsertRowsAbove(1);
                    qingneng.Cell(4, 1).Value = 1;
                    qingneng.Cell(4, 2).Value = "海南精研科技有限公司";
                    qingneng.Cell(4, 3).Value = "代理费";
                    qingneng.Cell(4, 30).Value = HainanStage2PaymentParties.Qinghui;
                    workbook.Save();
                }
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsFalse(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PaymentPartyRequired
                    && issue.Entity == "海南精研科技有限公司"));
                Assert.IsFalse(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingPaymentParties
                    && issue.Entity == "海南精研科技有限公司"));

                options.SummarySubjectDecisions.Add(new HainanStage2SummarySubjectDecision
                {
                    SettlementKind = "代理费",
                    Entity = "海南精研科技有限公司",
                    PaymentParty = HainanStage2PaymentParties.Qinghui
                });
                Assert.ThrowsException<InvalidOperationException>(() => RunAfterPreflight(service, options));
                options.SummarySubjectDecisions.Clear();

                var report = RunAfterPreflight(service, options);
                using (var workbook = new XLWorkbook(report.Summary))
                {
                    Assert.AreEqual(
                        HainanStage2PaymentParties.Qingneng,
                        workbook.Worksheet("汇总表").Cell(4, 36).GetFormattedString());
                    Assert.IsTrue(HainanStage2SummaryWorkbookWriter
                        .ReadSummaryMeta(workbook.Worksheet("清能汇总表"))
                        .Any(row => row.Entity == "海南精研科技有限公司"));
                    Assert.IsFalse(HainanStage2SummaryWorkbookWriter
                        .ReadSummaryMeta(workbook.Worksheet("清辉汇总表"))
                        .Any(row => row.Entity == "海南精研科技有限公司"));
                }
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
        public void GenerateSettlementStopsWhenFirstPhysicalRelationshipOwnerIsBlank()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(Path.GetDirectoryName(ledgerPath));
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                    var start = HainanLedgerLayout.MonthStartColumn(4);
                    worksheet.Cell(1, start).Value = "4月";
                    WriteLedgerRow(worksheet, 4, start, string.Empty, "负责人缺失代理", "零电量客户", 0);
                    WriteLedgerRow(worksheet, 5, start, "后续负责人", "负责人缺失代理", "结算客户", 100);
                    workbook.SaveAs(ledgerPath);
                }

                WriteProxyTemplate(proxyRoot, "后续负责人", "负责人缺失代理");
                WriteSummaryTemplate(summaryPath, "负责人缺失代理", "代理费", HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot,
                    AllowMissingOwner = true
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.FirstOwnerMissing
                    && issue.LedgerRow == 4
                    && issue.BlocksGeneration));

                Assert.ThrowsException<InvalidOperationException>(() => RunAfterPreflight(service, options));
                Assert.IsFalse(Directory.Exists(outputRoot));
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
        public void AnalyzeSettlementRequiresTemplateDecisionWhenBorrowedTemplateCandidateIsNotUnique()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "新增无模板代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "负责人甲", "模板代理甲");
                WriteProxyTemplateWithExcelDateSignature(proxyRoot, "负责人乙", "模板代理乙");
                WriteSummaryTemplate(summaryPath, "新增无模板代理", "代理费", HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                var templateIssue = preflight.Issues.Single(issue =>
                    issue.Code == Stage2PreflightIssueKinds.AmbiguousBorrowTemplates
                    && issue.Entity == "新增无模板代理");
                Assert.AreEqual(Stage2PreflightDisposition.RequiredDecision, templateIssue.Disposition);
                Assert.IsFalse(templateIssue.BlocksGeneration);
                Assert.IsTrue(templateIssue.RequiresTemplateSelection);
                Assert.AreEqual(2, templateIssue.TemplateOptions.Count);
                StringAssert.Contains(templateIssue.Suggestion, "选择");

                Assert.ThrowsException<InvalidOperationException>(() => RunAfterPreflight(service, options));
                Assert.IsFalse(Directory.Exists(outputRoot));

                var selectedTemplate = templateIssue.TemplateOptions.Single(path =>
                    Path.GetFileName(path).Contains("模板代理乙"));
                options.TemplateDecisions.Add(new HainanStage2TemplateDecision
                {
                    SettlementKind = "代理费",
                    Entity = "新增无模板代理",
                    TemplatePath = selectedTemplate
                });

                var report = RunAfterPreflight(service, options);
                Assert.IsTrue(report.AuditIssues.Exists(issue =>
                    issue.Category == "本次分表模板选择"
                    && issue.Entity == "新增无模板代理"
                    && issue.TemplateFile == selectedTemplate));
                var outputPath = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "测试负责人 - 海南2026",
                    "新增无模板代理 2026海南.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    Assert.AreEqual(1, workbook.Worksheets.Count);
                    var worksheet = workbook.Worksheet("4月");
                    Assert.IsNotNull(FindFormattedCell(worksheet, "2026年6月8日"));
                }
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
        public void GenerateSettlementExtendsDetailTotalsAndRepairsTemplateFormatting()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyRows(ledgerPath, "测试负责人", "测试代理", "存量客户", "新增客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "测试代理");
                var templatePath = Directory.GetFiles(proxyRoot, "*.xlsx", SearchOption.AllDirectories).Single();
                using (var workbook = new XLWorkbook(templatePath))
                {
                    workbook.Worksheet("3月").Cell("P6").CreateComment().AddText("本主体需要保留的人工批注");
                    workbook.Worksheet("3月").Visibility = XLWorksheetVisibility.Hidden;
                    workbook.Save();
                }
                WriteSummaryTemplate(summaryPath, "测试代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                RunAfterPreflight(options);

                var outputPath = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "测试负责人 - 海南2026",
                    "测试代理 2026海南.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = workbook.Worksheet("4月");
                    Assert.AreEqual(XLWorksheetVisibility.Visible, worksheet.Visibility);
                    Assert.IsTrue(worksheet.TabActive);
                    Assert.IsTrue(worksheet.TabSelected);
                    Assert.AreEqual(1, workbook.Worksheets.Count(sheet => sheet.TabActive));
                    Assert.AreEqual(1, workbook.Worksheets.Count(sheet => sheet.TabSelected));
                    Assert.AreEqual("SUM(C5:C6)", worksheet.Cell(7, 3).FormulaA1);
                    Assert.AreEqual("SUM(P5:P6)", worksheet.Cell(7, 16).FormulaA1);
                    Assert.AreEqual("日期：2026年05月08日", FindSignatureDateText(worksheet));
                    AssertStyleMatches(workbook.Worksheet("2月").Cell(6, 3), worksheet.Cell(7, 3));
                    AssertStyleMatches(worksheet.Cell(5, 2), worksheet.Cell(6, 2));
                    Assert.IsTrue(worksheet.CellsUsed(XLCellsUsedOptions.Comments).Any(cell => cell.HasComment));
                }
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
        public void GenerateSettlementShiftsBottomExcelDateCells()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyRows(ledgerPath, "测试负责人", "日期代理", "存量客户", "新增客户");
                WriteProxyTemplateWithExcelDateSignature(proxyRoot, "测试负责人", "日期代理");
                WriteSummaryTemplate(summaryPath, "日期代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                RunAfterPreflight(options);

                var outputPath = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "测试负责人 - 海南2026",
                    "日期代理 2026海南.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = workbook.Worksheet("4月");
                    var dateCell = FindFormattedCell(worksheet, "2026年6月8日");
                    Assert.AreEqual("yyyy\"年\"m\"月\"d\"日\";@", dateCell.Style.DateFormat.Format);
                }
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
        public void GenerateSettlementKeepsSummaryFooterOutOfDataRowsWhenAddingSubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyEntities(ledgerPath, "测试负责人", "存量代理", "新增代理");
                WriteProxyTemplate(proxyRoot, "测试负责人", "存量代理");
                WriteSummaryTemplateWithFooterInDataColumns(summaryPath, "存量代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };
                options.SummarySubjectDecisions.Add(new HainanStage2SummarySubjectDecision
                {
                    SettlementKind = "代理费",
                    Entity = "新增代理",
                    PaymentParty = HainanStage2PaymentParties.Qinghui
                });

                RunAfterPreflight(options);

                var outputPath = Path.Combine(outputRoot, "【2026年海南省代理费汇总表-4月自动化】.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = workbook.Worksheet("汇总表");
                    Assert.AreEqual("存量代理", worksheet.Cell(4, 2).GetFormattedString());
                    Assert.AreEqual("新增代理", worksheet.Cell(5, 2).GetFormattedString());
                    Assert.AreEqual("合计", worksheet.Cell(6, 1).GetFormattedString());
                    Assert.AreEqual("当月实际支付", worksheet.Cell(2, 27).GetFormattedString());
                    Assert.AreEqual("SUM(V4:V5)", worksheet.Cell(6, 22).FormulaA1);
                    Assert.AreEqual("日期：2026年06月08日", worksheet.Cell(9, 2).GetFormattedString());
                    Assert.AreEqual(string.Empty, worksheet.Cell(9, 1).GetFormattedString());
                    AssertStyleMatches(worksheet.Cell(4, 2), worksheet.Cell(5, 2));
                    Assert.AreEqual("清辉", worksheet.Cell(5, 36).GetFormattedString());
                }
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
        public void AnalyzeSettlementRequiresPaymentPartyForNewSummarySubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyEntities(ledgerPath, "测试负责人", "存量代理", "新增代理");
                WriteProxyTemplate(proxyRoot, "测试负责人", "存量代理");
                WriteSummaryTemplate(summaryPath, "存量代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.AreEqual(2, report.SubjectCount);
                Assert.IsTrue(report.RequiresPaymentPartySelection);
                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.RequiresPaymentPartySelection
                    && issue.Category == "新增汇总主体支付方选择"
                    && issue.Kind == "代理费"
                    && issue.Entity == "新增代理"
                    && issue.AvailablePaymentParties.Contains(HainanStage2PaymentParties.Qingneng)
                    && issue.AvailablePaymentParties.Contains(HainanStage2PaymentParties.Qinghui)));
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
        public void GenerateSettlementRejectsNewSummarySubjectWithoutPaymentPartyDecision()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyEntities(ledgerPath, "测试负责人", "存量代理", "新增代理");
                WriteProxyTemplate(proxyRoot, "测试负责人", "存量代理");
                WriteSummaryTemplate(summaryPath, "存量代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                    RunAfterPreflight(options));
                StringAssert.Contains(ex.Message, "新增汇总主体支付方未选择");
                StringAssert.Contains(ex.Message, "新增代理");
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
        public void GenerateSettlementUsesExplicitPaymentPartyForNewSummarySubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyEntities(ledgerPath, "测试负责人", "存量代理", "新增代理");
                WriteProxyTemplate(proxyRoot, "测试负责人", "存量代理");
                WriteSummaryTemplate(summaryPath, "存量代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };
                options.SummarySubjectDecisions.Add(new HainanStage2SummarySubjectDecision
                {
                    SettlementKind = "代理费",
                    Entity = "新增代理",
                    PaymentParty = HainanStage2PaymentParties.Qingneng
                });

                RunAfterPreflight(options);

                var outputPath = Path.Combine(outputRoot, "【2026年海南省代理费汇总表-4月自动化】.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = workbook.Worksheet("汇总表");
                    Assert.AreEqual("新增代理", worksheet.Cell(5, 2).GetFormattedString());
                    Assert.AreEqual("否", worksheet.Cell(5, 4).GetFormattedString());
                    Assert.AreEqual("新增代理", worksheet.Cell(5, 6).GetString());
                    Assert.AreEqual("平台", worksheet.Cell(5, 7).GetFormattedString());
                    Assert.AreEqual("J5-I5", worksheet.Cell(5, 8).FormulaA1);
                    Assert.AreEqual(0.06, worksheet.Cell(5, 9).GetDouble(), 0.0000001);
                    Assert.AreEqual(0.13, worksheet.Cell(5, 10).GetDouble(), 0.0000001);
                    Assert.AreEqual("测试负责人", worksheet.Cell(5, 11).GetFormattedString());
                    Assert.AreEqual(HainanStage2PaymentParties.Qingneng, worksheet.Cell(5, 36).GetFormattedString());
                }
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
        public void GenerateSettlementDoesNotKeepBorrowedTemplateHistoryOrCommentsForNewSubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyEntities(ledgerPath, "测试负责人", "模板代理", "新增代理");
                WriteProxyTemplate(proxyRoot, "测试负责人", "模板代理");
                var templatePath = Directory.GetFiles(proxyRoot, "*.xlsx", SearchOption.AllDirectories).Single();
                using (var workbook = new XLWorkbook(templatePath))
                {
                    workbook.Worksheet("3月").Cell("P6").CreateComment().AddText("模板主体的人工批注");
                    workbook.Save();
                }
                WriteSummaryTemplate(summaryPath, "模板代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };
                options.SummarySubjectDecisions.Add(new HainanStage2SummarySubjectDecision
                {
                    SettlementKind = "代理费",
                    Entity = "新增代理",
                    PaymentParty = HainanStage2PaymentParties.Qinghui
                });

                RunAfterPreflight(options);

                var outputPath = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "测试负责人 - 海南2026",
                    "新增代理 2026海南.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    Assert.AreEqual(1, workbook.Worksheets.Count);
                    Assert.IsTrue(workbook.Worksheets.Contains("4月"));
                    Assert.AreEqual("代理名称:新增代理", workbook.Worksheet("4月").Cell("A2").GetFormattedString());
                    var commentedCells = workbook.Worksheet("4月")
                        .CellsUsed(XLCellsUsedOptions.Comments)
                        .Where(cell => cell.HasComment)
                        .Select(cell => cell.Address.ToString())
                        .ToList();
                    Assert.AreEqual(0, commentedCells.Count, string.Join(", ", commentedCells));
                }
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
        public void AnalyzeSettlementReportsPreviousDetailRowsOutsideCurrentLedger()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "特殊代理", "存量客户");
                WriteProxyTemplateWithSpecialDetailRow(proxyRoot, "测试负责人", "特殊代理", "扣除三月少扣税费");
                WriteSummaryTemplate(summaryPath, "特殊代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.HasIssues);
                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Category == "上月分表存在本月台账外明细行"
                    && issue.Customer == "扣除三月少扣税费"
                    && issue.Entity == "特殊代理"
                    && issue.SheetName == "3月"));
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
        public void GenerateSettlementDoesNotCarryPreviousSpecialDetailRowsIntoCurrentMonth()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "特殊代理", "存量客户");
                WriteProxyTemplateWithSpecialDetailRow(proxyRoot, "测试负责人", "特殊代理", "扣除三月少扣税费");
                WriteSummaryTemplate(summaryPath, "特殊代理", "代理费", "清辉");

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                RunAfterPreflight(options);

                var outputPath = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "测试负责人 - 海南2026",
                    "特殊代理 2026海南.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    Assert.AreEqual("扣除三月少扣税费", workbook.Worksheet("3月").Cell(6, 2).GetFormattedString());

                    var current = workbook.Worksheet("4月");
                    Assert.AreEqual("存量客户", current.Cell(5, 2).GetFormattedString());
                    Assert.AreEqual("合计", current.Cell(6, 1).GetFormattedString());
                    Assert.AreNotEqual("扣除三月少扣税费", current.Cell(6, 2).GetFormattedString());
                    Assert.AreEqual("SUM(P5:P5)", current.Cell(6, 16).FormulaA1);
                }
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
        public void AnalyzeSettlementReportsLedgerDifferenceAndBlankPayeeAsReviewItems()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "预检代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "预检代理");
                WriteSummaryTemplate(summaryPath, "预检代理", "代理费", HainanStage2PaymentParties.Qinghui);

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.LedgerAmountDifference
                    && issue.Disposition == Stage2PreflightDisposition.Review
                    && issue.Entity == "预检代理"));
                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PayeeSourceMissing
                    && issue.Disposition == Stage2PreflightDisposition.Review
                    && issue.Entity == "预检代理"
                    && issue.CurrentValue == "保持空白"));
                Assert.IsFalse(report.HasBlockingIssues);
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
        public void AnalyzeSettlementBlocksMissingSelectedPaymentSheetOrphanAndExistingTargetMonth()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "结构预检代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "结构预检代理");
                WriteSummaryTemplateWithStructuralProblems(
                    summaryPath,
                    "结构预检代理",
                    "代理费",
                    HainanStage2PaymentParties.Qingneng);

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.SummaryPaymentSheetMissing
                    && issue.BlocksGeneration));
                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.SummaryOrphanSubject
                    && issue.Entity == "孤立主体"
                    && issue.BlocksGeneration));
                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.SummaryTargetMonthAlreadyExists
                    && issue.SheetName == "清辉汇总表"
                    && issue.Disposition == Stage2PreflightDisposition.Review));
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
        public void GenerateSettlementSafelyRewritesUniqueExistingTargetMonthBlock()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "重写月份代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "重写月份代理");
                WriteSummaryTemplate(summaryPath, "重写月份代理", "代理费", HainanStage2PaymentParties.Qinghui);
                using (var workbook = new XLWorkbook(summaryPath))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        worksheet.Cell(2, 16).Value = "2026年4月";
                        worksheet.Cell(4, 16).Value = 999;
                    }
                    workbook.Save();
                }

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };
                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());

                var preflight = service.Analyze(options);
                Assert.AreEqual(
                    3,
                    preflight.Issues.Count(issue =>
                        issue.Code == Stage2PreflightIssueKinds.SummaryTargetMonthAlreadyExists
                        && issue.Disposition == Stage2PreflightDisposition.Review));
                Assert.IsFalse(preflight.HasBlockingIssues);

                var report = RunAfterPreflight(service, options);
                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    Assert.AreEqual(1, HainanStage2SummaryWorkbookWriter.FindSummaryMonthBlocks(main, 4).Count);
                    Assert.AreEqual("累计代理费总计", main.Cell(2, 22).GetFormattedString());
                    Assert.AreNotEqual(999d, ClosedXmlUtil.CellNumber(main.Cell(4, 16)), 0.0001);
                }
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
        public void GenerateSettlementClampsFullyDeductedLoanRemainingAtZero()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "借支结清代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "借支结清代理");
                WriteLoanSummaryTemplate(summaryPath, "借支结清代理", "代理费", HainanStage2PaymentParties.Qinghui);
                using (var workbook = new XLWorkbook(summaryPath))
                {
                    var main = workbook.Worksheet("汇总表");
                    main.Cell(4, 15).Value = 0.1;
                    main.Cell(4, 21).Value = 0.1;
                    main.Cell(4, 27).Value = 0.1;
                    main.Cell(4, 31).Value = 1.3;
                    main.Cell(4, 32).Value = 0.3;
                    main.Cell(4, 35).Value = "每月扣除1万";
                    var payment = workbook.Worksheet("清辉汇总表");
                    main.Row(4).CopyTo(payment.Row(4));
                    workbook.Save();
                }

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = RunAfterPreflight(options);

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    var payment = workbook.Worksheet("清辉汇总表");
                    Assert.AreEqual("O4+U4+AA4+AG4", main.Cell(4, 38).FormulaA1);
                    Assert.AreEqual("MAX(0,AK4-AL4)", main.Cell(4, 39).FormulaA1);
                    Assert.AreEqual(main.Cell(4, 39).FormulaA1, payment.Cell(4, 39).FormulaA1);
                    Assert.AreEqual(new DateTime(2026, 4, 1), main.Cell(4, 42).GetDateTime());
                    Assert.AreEqual(main.Cell(4, 42).GetDateTime(), payment.Cell(4, 42).GetDateTime());
                    Assert.AreEqual("yyyy年m月", main.Cell(4, 42).Style.DateFormat.Format);
                }
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
        public void GenerateSettlementPreservesExistingLoanCompletionMonth()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptionsWithLoanSummary(
                    root,
                    "测试负责人",
                    "已有结清月份代理");
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    main.Cell(4, 15).Value = 0.1;
                    main.Cell(4, 21).Value = 0.1;
                    main.Cell(4, 27).Value = 0.1;
                    main.Cell(4, 31).Value = 1.3;
                    main.Cell(4, 32).Value = 0.3;
                    main.Cell(4, 35).Value = "每月扣除1万";
                    main.Cell(4, 36).Value = new DateTime(2026, 3, 1);
                    main.Cell(4, 36).Style.DateFormat.Format = "yyyy年m月";
                    var payment = workbook.Worksheet("清辉汇总表");
                    main.Row(4).CopyTo(payment.Row(4));
                    workbook.Save();
                }

                var report = RunAfterPreflight(options);

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    var payment = workbook.Worksheet("清辉汇总表");
                    Assert.AreEqual(1d, ClosedXmlUtil.CellNumber(main.Cell(4, 33)), 0.0001d);
                    Assert.AreEqual(new DateTime(2026, 3, 1), main.Cell(4, 42).GetDateTime());
                    Assert.AreEqual(main.Cell(4, 42).GetDateTime(), payment.Cell(4, 42).GetDateTime());
                    Assert.AreEqual("yyyy年m月", main.Cell(4, 42).Style.DateFormat.Format);
                }
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
        public void GenerateSettlementRecomputesLoanCompletionWhenTargetMonthAlreadyExists()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var firstOptions = CreateSingleProxyOptionsWithLoanSummary(
                    root,
                    "测试负责人",
                    "同月重算代理");
                using (var workbook = new XLWorkbook(firstOptions.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    main.Cell(4, 15).Value = 0.1;
                    main.Cell(4, 21).Value = 0.1;
                    main.Cell(4, 27).Value = 0.1;
                    main.Cell(4, 31).Value = 1.3;
                    main.Cell(4, 32).Value = 0.3;
                    main.Cell(4, 35).Value = "每月扣除1万";
                    var payment = workbook.Worksheet("清辉汇总表");
                    main.Row(4).CopyTo(payment.Row(4));
                    workbook.Save();
                }

                var firstReport = RunAfterPreflight(firstOptions);
                using (var workbook = new XLWorkbook(firstReport.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    var payment = workbook.Worksheet("清辉汇总表");
                    main.Cell(4, 38).Value = 1.3;
                    payment.Cell(4, 38).Value = 1.3;
                    main.Cell(4, 42).Clear(XLClearOptions.Contents);
                    payment.Cell(4, 42).Clear(XLClearOptions.Contents);
                    workbook.Save();
                }

                var rerunOptions = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = firstOptions.LedgerPath,
                    ProxyTemplateDirectory = firstOptions.ProxyTemplateDirectory,
                    IntermediaryTemplateDirectory = firstOptions.IntermediaryTemplateDirectory,
                    SummaryTemplatePath = firstReport.Summary,
                    OutputDirectory = Path.Combine(root, "rerun-output")
                };

                var rerunReport = RunAfterPreflight(rerunOptions);

                using (var workbook = new XLWorkbook(rerunReport.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    Assert.AreEqual(1d, ClosedXmlUtil.CellNumber(main.Cell(4, 33)), 0.0001d);
                    Assert.AreEqual("O4+U4+AA4+AG4", main.Cell(4, 38).FormulaA1);
                    Assert.AreEqual("MAX(0,AK4-AL4)", main.Cell(4, 39).FormulaA1);
                    Assert.AreEqual(new DateTime(2026, 4, 1), main.Cell(4, 42).GetDateTime());
                    Assert.AreEqual("yyyy年m月", main.Cell(4, 42).Style.DateFormat.Format);
                }
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
        public void GenerateSettlementDoesNotGuessLoanCompletionMonthWhenAlreadySettled()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptionsWithLoanSummary(
                    root,
                    "测试负责人",
                    "历史结清代理");
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    main.Cell(4, 15).Value = 0.5;
                    main.Cell(4, 21).Value = 0.4;
                    main.Cell(4, 27).Value = 0.4;
                    main.Cell(4, 31).Value = 1.3;
                    main.Cell(4, 32).Value = 1.3;
                    main.Cell(4, 35).Value = "每月扣除1万";
                    var payment = workbook.Worksheet("清辉汇总表");
                    main.Row(4).CopyTo(payment.Row(4));
                    workbook.Save();
                }

                var report = RunAfterPreflight(options);

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    Assert.AreEqual("MAX(0,AK4-AL4)", main.Cell(4, 39).FormulaA1);
                    Assert.IsTrue(main.Cell(4, 33).IsEmpty());
                    Assert.IsTrue(main.Cell(4, 42).IsEmpty());
                }
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
        public void GenerateSettlementSynchronizesCanonicalBusinessFieldsToExistingPaymentRow()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "长期字段代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "长期字段代理");
                WriteSummaryTemplate(summaryPath, "长期字段代理", "代理费", HainanStage2PaymentParties.Qinghui);
                using (var workbook = new XLWorkbook(summaryPath))
                {
                    var main = workbook.Worksheet("汇总表");
                    var payment = workbook.Worksheet("清辉汇总表");
                    main.Cell(4, 4).Value = "是";
                    main.Cell(4, 5).Value = "债权转让";
                    main.Cell(4, 6).Value = "张三、李四";
                    main.Cell(4, 7).Value = "专票";
                    main.Cell(4, 23).Value = 12.34;
                    main.Cell(4, 27).Value = "每月扣除1万";
                    main.Cell(4, 28).Value = new DateTime(2026, 3, 1);
                    main.Cell(4, 31).Value = "主表备注";
                    payment.Cell(4, 4).Value = "否";
                    payment.Cell(4, 5).Value = "错误协议";
                    payment.Cell(4, 6).Clear(XLClearOptions.Contents);
                    payment.Cell(4, 7).Value = "平台";
                    payment.Cell(4, 23).Value = 99;
                    payment.Cell(4, 27).Value = "错误扣除";
                    payment.Cell(4, 28).Value = new DateTime(2026, 2, 1);
                    payment.Cell(4, 31).Value = "错误备注";
                    workbook.Save();
                }

                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = RunAfterPreflight(options);

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    var payment = workbook.Worksheet("清辉汇总表");
                    foreach (var column in new[] { 2, 3, 4, 5, 6, 7, 9, 10, 11, 29, 33, 34, 37 })
                    {
                        Assert.AreEqual(
                            main.Cell(4, column).GetFormattedString(),
                            payment.Cell(4, column).GetFormattedString(),
                            "未同步业务列 " + column);
                    }

                    Assert.AreEqual("J4-I4", main.Cell(4, 8).FormulaA1);
                    Assert.AreEqual(main.Cell(4, 8).FormulaA1, payment.Cell(4, 8).FormulaA1);
                }
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
        public void GenerateSettlementPublishesVerifiedBatchWithFinalReportPaths()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "发布代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "发布代理");
                WriteSummaryTemplate(summaryPath, "发布代理", "代理费", HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                HainanStage2PreflightReport preflight;
                var report = RunAfterPreflight(service, options, out preflight);
                var outputPrefix = Path.GetFullPath(outputRoot).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;

                Assert.IsFalse(string.IsNullOrWhiteSpace(preflight.PreflightSignature));
                Assert.IsFalse(string.IsNullOrWhiteSpace(preflight.InputFingerprint));
                Assert.AreEqual(preflight.PreflightSignature, report.PreflightSignature);
                Assert.AreEqual(preflight.InputFingerprint, report.InputFingerprint);
                Assert.AreEqual(Path.GetFullPath(outputRoot), Path.GetFullPath(report.OutputDirectory));
                Assert.IsTrue(File.Exists(report.Summary));
                Assert.IsTrue(File.Exists(report.ReportPath));
                Assert.IsTrue(File.Exists(report.ValidationReportPath));
                Assert.IsTrue(File.Exists(report.HtmlReportPath));
                Assert.IsTrue(report.Groups.All(group =>
                    File.Exists(group.OutputFile)
                    && Path.GetFullPath(group.OutputFile).StartsWith(outputPrefix, StringComparison.OrdinalIgnoreCase)));
                Assert.IsTrue(Path.GetFullPath(report.Summary).StartsWith(outputPrefix, StringComparison.OrdinalIgnoreCase));
                Assert.AreEqual(
                    0,
                    Directory.GetDirectories(outputRoot, Stage2BatchWorkspace.StagingDirectoryPrefix + "*", SearchOption.TopDirectoryOnly).Length);
                Assert.AreEqual(
                    0,
                    Directory.GetDirectories(outputRoot, Stage2BatchWorkspace.FailedDirectoryPrefix + "*", SearchOption.TopDirectoryOnly).Length);
                var reportJson = File.ReadAllText(report.ReportPath);
                StringAssert.Contains(reportJson, preflight.PreflightSignature);
                StringAssert.Contains(reportJson, preflight.InputFingerprint);
                StringAssert.Contains(reportJson, Path.GetFileName(report.HtmlReportPath));
                Assert.IsFalse(reportJson.Contains(Stage2BatchWorkspace.StagingDirectoryPrefix));
                var readableReport = File.ReadAllText(report.HtmlReportPath);
                StringAssert.Contains(readableReport, "海南4月阶段二结算报告");
                StringAssert.Contains(readableReport, "发布代理");
                StringAssert.Contains(readableReport, "生成完成，但需要人工复核");
                Assert.IsFalse(readableReport.Contains(Stage2BatchWorkspace.StagingDirectoryPrefix));
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
        public void GenerateSettlementPublishFailureKeepsFormalOutputAndPreservesFailedBatch()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");
            var finalSummaryPath = Path.Combine(outputRoot, "【2026年海南省代理费汇总表-4月自动化】.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "失败发布代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "失败发布代理");
                WriteSummaryTemplate(summaryPath, "失败发布代理", "代理费", HainanStage2PaymentParties.Qinghui);
                File.WriteAllText(finalSummaryPath, "formal-sentinel");
                var collidingSplitPath = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "测试负责人 - 海南2026",
                    "失败发布代理 2026海南.xlsx");
                Directory.CreateDirectory(collidingSplitPath);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                    RunAfterPreflight(options));

                StringAssert.Contains(exception.Message, "未完成");
                Assert.AreEqual("formal-sentinel", File.ReadAllText(finalSummaryPath));
                Assert.IsTrue(Directory.Exists(collidingSplitPath));
                Assert.IsFalse(File.Exists(Path.Combine(outputRoot, "4月结算生成总报告.json")));
                Assert.IsFalse(File.Exists(Path.Combine(outputRoot, "海南4月阶段二结算报告.html")));
                var failedDirectories = Directory.GetDirectories(
                    outputRoot,
                    Stage2BatchWorkspace.FailedDirectoryPrefix + "*",
                    SearchOption.TopDirectoryOnly);
                Assert.AreEqual(1, failedDirectories.Length);
                Assert.IsTrue(File.Exists(Path.Combine(failedDirectories[0], Stage2BatchWorkspace.FailedMarkerFileName)));
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
        public void StrongIntegrityVerifiersRejectChangedSplitAmountAndPaymentPayee()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                WriteLedgerWithSingleProxyRow(ledgerPath, 4, "测试负责人", "强校验代理", "存量客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "强校验代理");
                WriteSummaryTemplate(summaryPath, "强校验代理", "代理费", HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };
                var report = RunAfterPreflight(options);
                var snapshot = HainanStage2LedgerReader.ReadSnapshot(ledgerPath, 4);

                using (var workbook = new XLWorkbook(report.Groups.Single().OutputFile))
                {
                    workbook.Worksheet("4月").Cell(5, 16).Value = 999;
                    workbook.Save();
                }
                Assert.ThrowsException<InvalidDataException>(() =>
                    HainanStage2SplitWorkbookWriter.VerifyGeneratedSplitWorkbooks(
                        snapshot.SubjectGroups,
                        report.Groups,
                        4));

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    main.Cell(4, 9).Value = snapshot.SubjectGroups.Single().TaxRate + 0.000000001;
                    workbook.Save();
                }
                Assert.ThrowsException<InvalidDataException>(() =>
                    HainanStage2SummaryWorkbookWriter.VerifyGeneratedSummary(
                        options,
                        report.Groups,
                        snapshot.SubjectGroups,
                        report.Summary));

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    workbook.Worksheet("汇总表").Cell(4, 9).Value = snapshot.SubjectGroups.Single().TaxRate;
                    workbook.Worksheet("清辉汇总表").Cell(4, 6).Value = "错误收款人";
                    workbook.Save();
                }
                Assert.ThrowsException<InvalidDataException>(() =>
                    HainanStage2SummaryWorkbookWriter.VerifyGeneratedSummary(
                        options,
                        report.Groups,
                        snapshot.SubjectGroups,
                        report.Summary));
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
        public void AnalyzeSettlementAllowsUnplannedManagedWorkbookWithOnlyOldMonthSheets()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "当前负责人", "受管输出代理");
                WriteManagedWorkbook(
                    Path.Combine(
                        options.OutputDirectory,
                        "2026年代理 - 海南",
                        "历史负责人 - 海南2026",
                        "历史额外主体.xlsx"),
                    "3月");

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsFalse(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.UnexpectedTargetMonthWorkbook
                    || issue.Code == Stage2PreflightIssueKinds.ManagedOutputUnreadable));
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
        public void AnalyzeSettlementBlocksPreviousOwnerWorkbookThatStillContainsTargetMonth()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "新负责人", "负责人迁移代理");
                var oldPath = Path.Combine(
                    options.OutputDirectory,
                    "2026年代理 - 海南",
                    "旧负责人 - 海南2026",
                    "负责人迁移代理 2026海南.xlsx");
                WriteManagedWorkbook(oldPath, "3月", "4月");

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.UnexpectedTargetMonthWorkbook
                    && issue.TemplateFile == Path.GetFullPath(oldPath)
                    && issue.BlocksGeneration));
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
        public void AnalyzeAndGenerateAllowPlannedWorkbookThatAlreadyContainsTargetMonth()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "当前负责人", "计划覆盖代理");
                var plannedPath = Path.Combine(
                    options.OutputDirectory,
                    "2026年代理 - 海南",
                    "当前负责人 - 海南2026",
                    "计划覆盖代理 2026海南.xlsx");
                WriteManagedWorkbook(plannedPath, "4月");
                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());

                var preflight = service.Analyze(options);
                Assert.IsFalse(preflight.Issues.Exists(issue =>
                    (issue.Code == Stage2PreflightIssueKinds.UnexpectedTargetMonthWorkbook
                        || issue.Code == Stage2PreflightIssueKinds.ManagedOutputUnreadable)
                    && issue.BlocksGeneration));
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PlannedTargetMonthWorkbook
                    && issue.Disposition == Stage2PreflightDisposition.Review
                    && issue.TemplateFile == Path.GetFullPath(plannedPath)));

                var report = RunAfterPreflight(service, options);
                Assert.AreEqual(Path.GetFullPath(plannedPath), Path.GetFullPath(report.Groups.Single().OutputFile));
                using (var workbook = new XLWorkbook(plannedPath))
                {
                    Assert.AreEqual(1, workbook.Worksheets.Count(sheet => sheet.Name == "4月"));
                    Assert.AreEqual("存量客户", workbook.Worksheet("4月").Cell(5, 2).GetFormattedString());
                }
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
        public void GenerateSettlementRejectsPlannedWorkbookChangedAfterConfirmedPreflight()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "当前负责人", "预检后变更代理");
                var plannedPath = Path.Combine(
                    options.OutputDirectory,
                    "2026年代理 - 海南",
                    "当前负责人 - 海南2026",
                    "预检后变更代理 2026海南.xlsx");
                WriteManagedWorkbook(plannedPath, "4月");
                var service = new HainanStage2Service(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                options.ExpectedPreflightSignature = preflight.PreflightSignature;
                options.ExpectedInputFingerprint = preflight.InputFingerprint;
                using (var workbook = new XLWorkbook(plannedPath))
                {
                    workbook.Worksheet("4月").Cell("A1").Value = "预检后手工修改";
                    workbook.Save();
                }

                var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                    service.Run(options, null));

                StringAssert.Contains(exception.Message, "输入已变化");
                using (var workbook = new XLWorkbook(plannedPath))
                {
                    Assert.AreEqual("预检后手工修改", workbook.Worksheet("4月").Cell("A1").GetString());
                }
                Assert.IsFalse(File.Exists(Path.Combine(
                    options.OutputDirectory,
                    "【2026年海南省代理费汇总表-4月自动化】.xlsx")));
                Assert.AreEqual(
                    0,
                    Directory.GetDirectories(
                        options.OutputDirectory,
                        Stage2BatchWorkspace.StagingDirectoryPrefix + "*",
                        SearchOption.TopDirectoryOnly).Length);
                Assert.AreEqual(
                    0,
                    Directory.GetDirectories(
                        options.OutputDirectory,
                        Stage2BatchWorkspace.FailedDirectoryPrefix + "*",
                        SearchOption.TopDirectoryOnly).Length);
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
        public void AnalyzeSettlementBlocksMissingFixedPaymentSheetWhenNoActiveSubjectExists()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var summaryPath = Path.Combine(root, "summary.xlsx");
            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                    worksheet.Cell(1, HainanLedgerLayout.MonthStartColumn(4)).Value = "4月";
                    workbook.SaveAs(ledgerPath);
                }
                WriteSummaryTemplate(summaryPath, "历史代理", "代理费", HainanStage2PaymentParties.Qinghui);
                using (var workbook = new XLWorkbook(summaryPath))
                {
                    workbook.Worksheet("清能汇总表").Delete();
                    workbook.Save();
                }
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = Path.Combine(root, "output")
                };

                var preflight = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.AreEqual(0, preflight.SubjectCount);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.SummaryPaymentSheetMissing
                    && issue.CurrentValue == "缺少清能汇总表"
                    && issue.BlocksGeneration));
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
        public void AnalyzeSettlementBlocksAmbiguousPaymentSummarySheetRole()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "工作表歧义代理");
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    PopulateSummaryTemplateSheet(
                        workbook.AddWorksheet("清能结算汇总表"),
                        "工作表歧义代理",
                        "代理费",
                        null,
                        HainanStage2PaymentParties.Qingneng);
                    workbook.Save();
                }

                var preflight = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.SummarySheetAmbiguous
                    && issue.CurrentValue.Contains("清能汇总表")
                    && issue.CurrentValue.Contains("清能结算汇总表")
                    && issue.BlocksGeneration));
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
        public void AnalyzeSettlementBlocksAmbiguousMainSummaryExactAndHistoricalAliases()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "主汇总歧义代理");
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    PopulateSummaryTemplateSheet(
                        workbook.AddWorksheet("历史汇总表"),
                        "主汇总歧义代理",
                        "代理费",
                        null,
                        HainanStage2PaymentParties.Qinghui);
                    workbook.Save();
                }

                var preflight = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.SummarySheetAmbiguous
                    && issue.CurrentValue.Split('、').Contains("汇总表")
                    && issue.CurrentValue.Split('、').Contains("历史汇总表")
                    && issue.BlocksGeneration));
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
        public void AnalyzeSettlementBlocksUnknownSummarySubjectKindAndReaderRejectsIt()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            try
            {
                var options = CreateSingleProxyOptions(root, "测试负责人", "费用类型异常代理");
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    workbook.Worksheet("汇总表").Cell(4, 3).Value = "其他费用";
                    workbook.Save();
                }

                var preflight = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);
                Assert.IsTrue(preflight.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.SummarySubjectKindInvalid
                    && issue.SheetName == "汇总表"
                    && issue.Entity == "费用类型异常代理"
                    && issue.CurrentValue == "其他费用"
                    && issue.BlocksGeneration));

                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var exception = Assert.ThrowsException<InvalidDataException>(() =>
                        HainanStage2SummaryWorkbookWriter.ReadSummaryMeta(workbook.Worksheet("汇总表")));
                    StringAssert.Contains(exception.Message, "只允许代理费或居间费");
                }
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
        public void AnalyzeSettlementBlocksTwoSubjectsPlannedToSameSafeOutputPath()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");
            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(root);
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                    var start = HainanLedgerLayout.MonthStartColumn(4);
                    worksheet.Cell(1, start).Value = "4月";
                    WriteLedgerRow(worksheet, 4, start, "同一负责人", "冲突/代理", "客户甲", 100);
                    WriteLedgerRow(worksheet, 5, start, "同一负责人", "冲突\\代理", "客户乙", 100);
                    workbook.SaveAs(ledgerPath);
                }
                WriteProxyTemplate(proxyRoot, "模板负责人", "唯一借用模板");
                WriteSummaryTemplate(summaryPath, "占位主体", "代理费", HainanStage2PaymentParties.Qinghui);
                var options = new HainanStage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                var report = new HainanStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Exists(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PlannedOutputPathConflict
                    && issue.BlocksGeneration));
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        private static HainanStage2Report RunAfterPreflight(HainanStage2Options options)
        {
            return RunAfterPreflight(
                new HainanStage2Service(new ClosedXmlSettlementExcelGateway()),
                options);
        }

        private static HainanStage2Report RunAfterPreflight(
            HainanStage2Service service,
            HainanStage2Options options)
        {
            HainanStage2PreflightReport preflight;
            return RunAfterPreflight(service, options, out preflight);
        }

        private static HainanStage2Report RunAfterPreflight(
            HainanStage2Service service,
            HainanStage2Options options,
            out HainanStage2PreflightReport preflight)
        {
            preflight = service.Analyze(options);
            options.ExpectedPreflightSignature = preflight.PreflightSignature;
            options.ExpectedInputFingerprint = preflight.InputFingerprint;
            return service.Run(options, null);
        }

        private static void WriteLedgerWithProxyRows(
            string path,
            string owner,
            string proxyEntity,
            string existingCustomer,
            string newCustomer)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                var start = HainanLedgerLayout.MonthStartColumn(4);
                worksheet.Cell(1, start).Value = "4月";
                WriteLedgerRow(worksheet, 4, start, owner, proxyEntity, existingCustomer, 100);
                WriteLedgerRow(worksheet, 5, start, owner, proxyEntity, newCustomer, 200);
                workbook.SaveAs(path);
            }
        }

        private static HainanStage2Options CreateSingleProxyOptions(
            string root,
            string owner,
            string entity)
        {
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");
            Directory.CreateDirectory(proxyRoot);
            Directory.CreateDirectory(interRoot);
            WriteLedgerWithSingleProxyRow(ledgerPath, 4, owner, entity, "存量客户");
            WriteProxyTemplate(proxyRoot, owner, entity);
            WriteSummaryTemplate(summaryPath, entity, "代理费", HainanStage2PaymentParties.Qinghui);
            return new HainanStage2Options
            {
                Month = 4,
                LedgerPath = ledgerPath,
                ProxyTemplateDirectory = proxyRoot,
                IntermediaryTemplateDirectory = interRoot,
                SummaryTemplatePath = summaryPath,
                OutputDirectory = outputRoot
            };
        }

        private static HainanStage2Options CreateSingleProxyOptionsWithLoanSummary(
            string root,
            string owner,
            string entity)
        {
            var options = CreateSingleProxyOptions(root, owner, entity);
            WriteLoanSummaryTemplate(
                options.SummaryTemplatePath,
                entity,
                "代理费",
                HainanStage2PaymentParties.Qinghui);
            return options;
        }

        private static void WriteManagedWorkbook(string path, params string[] sheetNames)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                foreach (var sheetName in sheetNames)
                {
                    workbook.AddWorksheet(sheetName);
                }
                workbook.SaveAs(path);
            }
        }

        private static void WriteLedgerWithSingleProxyRow(
            string path,
            int month,
            string owner,
            string proxyEntity,
            string customer)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                var start = HainanLedgerLayout.MonthStartColumn(month);
                worksheet.Cell(1, start).Value = month + "月";
                WriteLedgerRow(worksheet, 4, start, owner, proxyEntity, customer, 100);
                workbook.SaveAs(path);
            }
        }

        private static void WriteLedgerWithProxyEntities(
            string path,
            string owner,
            string existingEntity,
            string newEntity)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                var start = HainanLedgerLayout.MonthStartColumn(4);
                worksheet.Cell(1, start).Value = "4月";
                WriteLedgerRow(worksheet, 4, start, owner, existingEntity, "存量客户", 100);
                WriteLedgerRow(worksheet, 5, start, owner, newEntity, "新增主体客户", 200);
                workbook.SaveAs(path);
            }
        }

        private static void WriteLedgerRow(
            IXLWorksheet worksheet,
            int row,
            int start,
            string owner,
            string proxyEntity,
            string customer,
            double total)
        {
            worksheet.Cell(row, 3).Value = customer;
            worksheet.Cell(row, 8).Value = proxyEntity;
            worksheet.Cell(row, 10).Value = owner;
            worksheet.Cell(row, start).Value = total;
            worksheet.Cell(row, start + 1).Value = total * 0.1;
            worksheet.Cell(row, start + 2).Value = total * 0.2;
            worksheet.Cell(row, start + 3).Value = total * 0.3;
            worksheet.Cell(row, start + 4).Value = total * 0.4;
            worksheet.Cell(row, start + 5).Value = total * 0.5;
            worksheet.Cell(row, start + 6).Value = total * 0.6;
            worksheet.Cell(row, start + 13).Value = 0.5;
            worksheet.Cell(row, start + 14).Value = 1.2;
            worksheet.Cell(row, start + 16).Value = 0.06;
        }

        private static void WriteProxyTemplate(string root, string owner, string entity)
        {
            var folder = Path.Combine(root, owner + " - 海南2026");
            Directory.CreateDirectory(folder);
            var path = Path.Combine(folder, entity + " 2026海南.xlsx");
            using (var workbook = new XLWorkbook())
            {
                WriteDetailTemplateSheet(workbook.AddWorksheet("2月"), "代理", entity, 2, true);
                WriteDetailTemplateSheet(workbook.AddWorksheet("3月"), "代理", entity, 3, false);
                workbook.SaveAs(path);
            }
        }

        private static void WriteProxyTemplateWithExcelDateSignature(string root, string owner, string entity)
        {
            var folder = Path.Combine(root, owner + " - 海南2026");
            Directory.CreateDirectory(folder);
            var path = Path.Combine(folder, entity + " 2026海南.xlsx");
            using (var workbook = new XLWorkbook())
            {
                WriteDetailTemplateSheet(workbook.AddWorksheet("2月"), "代理", entity, 2, true);
                var sheet = workbook.AddWorksheet("3月");
                WriteDetailTemplateSheet(sheet, "代理", entity, 3, true);
                sheet.Cell("B8").Clear(XLClearOptions.Contents);
                sheet.Cell("N9").Value = new DateTime(2026, 5, 8);
                sheet.Cell("N9").Style.DateFormat.Format = "yyyy\"年\"m\"月\"d\"日\";@";
                sheet.Range("N9:O9").Merge();
                workbook.SaveAs(path);
            }
        }

        private static void WriteProxyTemplateWithSpecialDetailRow(string root, string owner, string entity, string specialText)
        {
            var folder = Path.Combine(root, owner + " - 海南2026");
            Directory.CreateDirectory(folder);
            var path = Path.Combine(folder, entity + " 2026海南.xlsx");
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("3月");
                WriteDetailTemplateSheet(sheet, "代理", entity, 3, true);
                sheet.Row(6).InsertRowsAbove(1);
                sheet.Row(5).CopyTo(sheet.Row(6));
                sheet.Cell(6, 1).Value = 2;
                sheet.Cell(6, 2).Value = specialText;
                sheet.Cell(6, 3).Clear(XLClearOptions.Contents);
                sheet.Cell(6, 10).Clear(XLClearOptions.Contents);
                sheet.Cell(6, 11).Clear(XLClearOptions.Contents);
                sheet.Cell(6, 16).Value = -0.0683;
                sheet.Cell(7, 1).Value = "合计";
                for (var column = 3; column <= 7; column++)
                {
                    var letter = ColumnLetter(column);
                    sheet.Cell(7, column).FormulaA1 = "SUM(" + letter + "5:" + letter + "6)";
                }

                for (var column = 12; column <= 16; column++)
                {
                    var letter = ColumnLetter(column);
                    sheet.Cell(7, column).FormulaA1 = "SUM(" + letter + "5:" + letter + "6)";
                }

                workbook.SaveAs(path);
            }
        }

        private static void WriteDetailTemplateSheet(
            IXLWorksheet worksheet,
            string kind,
            string entity,
            int month,
            bool styleTotalFormulaCells)
        {
            worksheet.Cell("A1").Value = kind + "费用结算单";
            worksheet.Cell("A2").Value = "代理名称:" + entity;
            worksheet.Cell("F2").Value = "所属期：2026年" + month.ToString("00") + "月";
            worksheet.Cell("M2").Value = "结算日期：2026 年 " + (month + 1).ToString("00") + " 月 15 日";

            worksheet.Cell(5, 1).Value = 1;
            worksheet.Cell(5, 2).Value = "存量客户";
            ApplyDetailRowStyle(worksheet.Row(5));

            worksheet.Cell(6, 1).Value = "合计";
            for (var column = 3; column <= 7; column++)
            {
                var letter = ColumnLetter(column);
                worksheet.Cell(6, column).FormulaA1 = "SUM(" + letter + "5:" + letter + "5)";
            }

            for (var column = 12; column <= 16; column++)
            {
                var letter = ColumnLetter(column);
                worksheet.Cell(6, column).FormulaA1 = "SUM(" + letter + "5:" + letter + "5)";
            }

            worksheet.Cell(8, 2).Value = "日期：2026年" + (month + 1).ToString("00") + "月08日";

            if (styleTotalFormulaCells)
            {
                ApplyTotalFormulaStyle(worksheet.Row(6));
            }
        }

        private static void ApplyDetailRowStyle(IXLRow row)
        {
            row.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            row.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            row.Style.NumberFormat.Format = "0.0000";
        }

        private static void ApplyTotalFormulaStyle(IXLRow row)
        {
            row.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            row.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            row.Style.Border.InsideBorder = XLBorderStyleValues.Medium;
            row.Style.NumberFormat.Format = "0.00";
        }

        private static void WriteSummaryTemplate(string path, string entity, string kind, string paymentParty)
        {
            using (var workbook = new XLWorkbook())
            {
                PopulateSummaryTemplateSheet(workbook.AddWorksheet("汇总表"), entity, kind, null, paymentParty);
                var qingneng = workbook.AddWorksheet("清能汇总表");
                var qinghui = workbook.AddWorksheet("清辉汇总表");
                if (paymentParty == HainanStage2PaymentParties.Qingneng)
                {
                    PopulateSummaryTemplateSheet(qingneng, entity, kind, null, paymentParty);
                    PopulateEmptySummaryTemplateSheet(qinghui);
                }
                else
                {
                    PopulateEmptySummaryTemplateSheet(qingneng);
                    PopulateSummaryTemplateSheet(qinghui, entity, kind, null, paymentParty);
                }
                workbook.SaveAs(path);
            }
        }

        private static void WriteLoanSummaryTemplate(
            string path,
            string entity,
            string kind,
            string paymentParty)
        {
            using (var workbook = new XLWorkbook())
            {
                PopulateLoanSummaryTemplateSheet(
                    workbook.AddWorksheet("汇总表"),
                    entity,
                    kind,
                    paymentParty);
                PopulateEmptyLoanSummaryTemplateSheet(workbook.AddWorksheet("清能汇总表"));
                PopulateLoanSummaryTemplateSheet(
                    workbook.AddWorksheet("清辉汇总表"),
                    entity,
                    kind,
                    paymentParty);
                workbook.SaveAs(path);
            }
        }

        private static void PopulateLoanSummaryTemplateSheet(
            IXLWorksheet worksheet,
            string entity,
            string kind,
            string paymentParty)
        {
            WriteLoanSummaryHeaders(worksheet);
            worksheet.Cell(4, 1).Value = 1;
            worksheet.Cell(4, 2).Value = entity;
            worksheet.Cell(4, 3).Value = kind;
            worksheet.Cell(4, 6).Value = entity;
            worksheet.Cell(4, 38).Value = paymentParty;
            worksheet.Cell(5, 1).Value = "合计";
            worksheet.Cell(8, 39).Value = "日期：2026年05月08日";
        }

        private static void PopulateEmptyLoanSummaryTemplateSheet(IXLWorksheet worksheet)
        {
            WriteLoanSummaryHeaders(worksheet);
            worksheet.Cell(4, 1).Value = "合计";
            worksheet.Cell(7, 39).Value = "日期：2026年05月08日";
        }

        private static void WriteLoanSummaryHeaders(IXLWorksheet worksheet)
        {
            for (var month = 1; month <= 3; month++)
            {
                var startColumn = 12 + (month - 1) * 6;
                worksheet.Cell(2, startColumn).Value = "2026年" + month + "月";
                worksheet.Cell(3, startColumn).Value = "代理费";
                worksheet.Cell(3, startColumn + 1).Value = "居间费";
                worksheet.Cell(3, startColumn + 2).Value = "退补电费";
                worksheet.Cell(3, startColumn + 3).Value = "当月抵扣";
                worksheet.Cell(3, startColumn + 4).Value = "费用合计";
                worksheet.Cell(2, startColumn + 5).Value = "当月实际支付";
            }

            worksheet.Cell(2, 30).Value = "累计代理费总计";
        }

        private static void WriteSummaryTemplateWithStructuralProblems(
            string path,
            string entity,
            string kind,
            string paymentParty)
        {
            using (var workbook = new XLWorkbook())
            {
                PopulateSummaryTemplateSheet(workbook.AddWorksheet("汇总表"), entity, kind, entity, paymentParty);
                var paymentSheet = workbook.AddWorksheet("清辉汇总表");
                PopulateSummaryTemplateSheet(
                    paymentSheet,
                    "孤立主体",
                    kind,
                    "孤立收款人",
                    HainanStage2PaymentParties.Qinghui);
                paymentSheet.Cell(2, 16).Value = "2026年4月";
                workbook.SaveAs(path);
            }
        }

        private static void WriteSummaryTemplateWithReliableSources(
            string path,
            string entity,
            string kind,
            string mainPayee,
            string mainPaymentParty,
            string secondaryPayee,
            string secondaryPaymentParty)
        {
            using (var workbook = new XLWorkbook())
            {
                PopulateSummaryTemplateSheet(
                    workbook.AddWorksheet("汇总表"),
                    entity,
                    kind,
                    mainPayee,
                    mainPaymentParty);
                PopulateSummaryTemplateSheet(
                    workbook.AddWorksheet("清辉汇总表"),
                    entity,
                    kind,
                    secondaryPayee,
                    secondaryPaymentParty);
                PopulateEmptySummaryTemplateSheet(workbook.AddWorksheet("清能汇总表"));
                workbook.SaveAs(path);
            }
        }

        private static void WriteSummaryTemplateWithMissingPaymentSheetSubject(
            string path,
            string entity,
            string kind,
            string payee)
        {
            using (var workbook = new XLWorkbook())
            {
                var main = workbook.AddWorksheet("汇总表");
                PopulateSummaryTemplateSheet(main, entity, kind, payee, HainanStage2PaymentParties.Qingneng);
                main.Cell(4, 4).Value = "是";
                main.Cell(4, 7).Value = "专票";

                PopulateEmptySummaryTemplateSheet(workbook.AddWorksheet("清能汇总表"));
                PopulateEmptySummaryTemplateSheet(workbook.AddWorksheet("清辉汇总表"));
                workbook.SaveAs(path);
            }
        }

        private static void PopulateEmptySummaryTemplateSheet(IXLWorksheet worksheet)
        {
            WriteSummaryHeaders(worksheet);
            worksheet.Cell(2, 22).Value = "累计代理费总计";
            worksheet.Cell(4, 1).Value = "合计";
            worksheet.Cell(7, 31).Value = "日期：2026年05月08日";
        }

        private static void WriteSummaryTemplateWithDuplicateSubject(
            string path,
            string entity,
            string kind,
            string paymentParty)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("汇总表");
                PopulateSummaryTemplateSheet(worksheet, entity, kind, entity, paymentParty);
                worksheet.Row(5).InsertRowsAbove(1);
                worksheet.Row(4).CopyTo(worksheet.Row(5));
                PopulateEmptySummaryTemplateSheet(workbook.AddWorksheet("清能汇总表"));
                PopulateEmptySummaryTemplateSheet(workbook.AddWorksheet("清辉汇总表"));
                workbook.SaveAs(path);
            }
        }

        private static void PopulateSummaryTemplateSheet(
            IXLWorksheet worksheet,
            string entity,
            string kind,
            string payee,
            string paymentParty)
        {
            WriteSummaryHeaders(worksheet);
            worksheet.Cell(2, 22).Value = "累计代理费总计";
            worksheet.Cell(4, 1).Value = 1;
            worksheet.Cell(4, 2).Value = entity;
            worksheet.Cell(4, 3).Value = kind;
            if (payee != null)
            {
                worksheet.Cell(4, 6).Value = payee;
            }

            worksheet.Cell(4, 30).Value = paymentParty;
            worksheet.Cell(5, 1).Value = "合计";
            worksheet.Cell(8, 31).Value = "日期：2026年05月08日";
        }

        private static void WriteSummaryTemplateWithFooterInDataColumns(string path, string entity, string kind, string paymentParty)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("汇总表");
                WriteSummaryHeaders(worksheet);
                worksheet.Cell(2, 22).Value = "累计代理费总计";
                worksheet.Cell(4, 1).Value = 1;
                worksheet.Cell(4, 2).Value = entity;
                worksheet.Cell(4, 3).Value = kind;
                worksheet.Cell(4, 30).Value = paymentParty;
                ApplySummaryDataRowStyle(worksheet.Row(4));
                worksheet.Cell(5, 1).Value = "合计";
                worksheet.Cell(8, 2).Value = "日期：2026年05月08日";
                var qingneng = workbook.AddWorksheet("清能汇总表");
                var qinghui = workbook.AddWorksheet("清辉汇总表");
                if (paymentParty == HainanStage2PaymentParties.Qingneng)
                {
                    PopulateSummaryTemplateSheet(qingneng, entity, kind, null, paymentParty);
                    PopulateEmptySummaryTemplateSheet(qinghui);
                }
                else
                {
                    PopulateEmptySummaryTemplateSheet(qingneng);
                    PopulateSummaryTemplateSheet(qinghui, entity, kind, null, paymentParty);
                }
                workbook.SaveAs(path);
            }
        }

        private static void WriteSummaryHeaders(IXLWorksheet worksheet)
        {
            worksheet.Range(2, 16, 2, 20).Merge();
            worksheet.Cell(2, 16).Value = "2026年3月";
            worksheet.Cell(3, 16).Value = "代理费";
            worksheet.Cell(3, 17).Value = "居间费";
            worksheet.Cell(3, 18).Value = "退补电费";
            worksheet.Cell(3, 19).Value = "当月抵扣";
            worksheet.Cell(3, 20).Value = "费用合计";
            worksheet.Range(2, 21, 3, 21).Merge();
            worksheet.Cell(2, 21).Value = "当月实际支付";
        }

        private static void ApplySummaryDataRowStyle(IXLRow row)
        {
            row.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            row.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            row.Style.NumberFormat.Format = "0.00";
        }

        private static string FindSignatureDateText(IXLWorksheet worksheet)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                var value = cell.GetFormattedString();
                if (value.Contains("日期：") && !value.Contains("结算日期"))
                {
                    return value;
                }
            }

            Assert.Fail("未找到签字日期单元格。");
            return null;
        }

        private static IXLCell FindFormattedCell(IXLWorksheet worksheet, string formattedText)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.GetFormattedString() == formattedText)
                {
                    return cell;
                }
            }

            Assert.Fail("未找到显示文本为“" + formattedText + "”的单元格。");
            return null;
        }

        private static void AssertStyleMatches(IXLCell expected, IXLCell actual)
        {
            Assert.AreEqual(expected.Style.Alignment.Horizontal, actual.Style.Alignment.Horizontal);
            Assert.AreEqual(expected.Style.Border.LeftBorder, actual.Style.Border.LeftBorder);
            Assert.AreEqual(expected.Style.Border.RightBorder, actual.Style.Border.RightBorder);
            Assert.AreEqual(expected.Style.Border.TopBorder, actual.Style.Border.TopBorder);
            Assert.AreEqual(expected.Style.Border.BottomBorder, actual.Style.Border.BottomBorder);
            Assert.AreEqual(expected.Style.NumberFormat.Format, actual.Style.NumberFormat.Format);
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
    }
}
