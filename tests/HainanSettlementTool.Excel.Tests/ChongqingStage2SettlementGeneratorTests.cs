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
        public void AnalyzeSettlementKeepsPaymentRequirementWhenDecisionAlreadyProvided()
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

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.RequiresPaymentPartySelection
                    && issue.SettlementKind == ChongqingStage2SettlementKinds.Proxy
                    && issue.Entity == "新增代理"));
                var evaluation = Stage2PreflightPolicy.Evaluate(report.Issues, options.SummarySubjectDecisions);
                Assert.IsTrue(evaluation.DecisionResolutions.Any(item =>
                    item.SettlementKind == ChongqingStage2SettlementKinds.Proxy
                    && item.Entity == "新增代理"
                    && item.Status == Stage2PaymentPartyDecisionStatus.Resolved));
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Intermediary, "新增居间");
                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Refund, "新增退补");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementPaymentDecisionDoesNotChangePreflightFactsOrSignature()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteProxyTemplate(Path.Combine(options.ProxyTemplateDirectory, "测试负责人", "新增代理.xlsx"));
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(options.SummaryTemplatePath);
                var service = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway());

                var before = service.Analyze(options);
                var beforeNewCustomer = before.Issues.Single(issue =>
                    issue.Code == Stage2PreflightIssueKinds.NewCustomer
                    && issue.SettlementKind == ChongqingStage2SettlementKinds.Proxy
                    && issue.Entity == "新增代理");
                AssertHasPaymentIssue(before, ChongqingStage2SettlementKinds.Proxy, "新增代理");
                StringAssert.Contains(beforeNewCustomer.CurrentValue, "支付方：待选择");

                options.SummarySubjectDecisions.Add(new ChongqingStage2SummarySubjectDecision
                {
                    SettlementKind = ChongqingStage2SettlementKinds.Proxy,
                    Entity = "新增代理",
                    PaymentParty = ChongqingStage2PaymentParties.Qingneng
                });
                var after = service.Analyze(options);
                var afterNewCustomer = after.Issues.Single(issue =>
                    issue.Code == Stage2PreflightIssueKinds.NewCustomer
                    && issue.SettlementKind == ChongqingStage2SettlementKinds.Proxy
                    && issue.Entity == "新增代理");

                AssertHasPaymentIssue(after, ChongqingStage2SettlementKinds.Proxy, "新增代理");
                Assert.AreEqual(before.Issues.Count, after.Issues.Count);
                Assert.AreEqual(beforeNewCustomer.CurrentValue, afterNewCustomer.CurrentValue);
                Assert.AreEqual(before.PreflightSignature, after.PreflightSignature);
                Assert.AreEqual(before.InputFingerprint, after.InputFingerprint);
                var evaluation = Stage2PreflightPolicy.Evaluate(after.Issues, options.SummarySubjectDecisions);
                Assert.IsTrue(evaluation.DecisionResolutions.Any(item =>
                    item.SettlementKind == ChongqingStage2SettlementKinds.Proxy
                    && item.Entity == "新增代理"
                    && item.Status == Stage2PaymentPartyDecisionStatus.Resolved));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementRejectsMissingPaymentDecisionsBeforeWritingAnyWorkbook()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteProxyTemplate(Path.Combine(options.ProxyTemplateDirectory, "测试负责人", "新增代理.xlsx"));
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(options.SummaryTemplatePath);
                AuthorizePreflight(options);

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null));

                StringAssert.Contains(ex.Message, "支付方");
                Assert.IsFalse(Directory.Exists(options.OutputDirectory));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementRejectsDirectRunWithoutPreflightSignatures()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null));

                StringAssert.Contains(ex.Message, "预检");
                Assert.IsFalse(Directory.Exists(options.OutputDirectory));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementRejectsAbsoluteOutputSummaryNameWithoutTouchingOutsideFile()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                var outsidePath = Path.Combine(root, "outside-absolute.xlsx");
                File.WriteAllText(outsidePath, "outside-content");
                options.OutputSummaryName = Path.GetFullPath(outsidePath);

                var ex = Assert.ThrowsException<ArgumentException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options));
                var defense = Assert.ThrowsException<InvalidOperationException>(() =>
                    ChongqingStage2SummaryWorkbookWriter.PlanOutputPath(options));

                StringAssert.Contains(ex.Message, "纯 .xlsx 文件名");
                StringAssert.Contains(defense.Message, "严格位于");
                Assert.AreEqual("outside-content", File.ReadAllText(outsidePath));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementRejectsParentTraversalOutputSummaryNameWithoutTouchingOutsideFile()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                var outsidePath = Path.Combine(root, "outside-parent.xlsx");
                File.WriteAllText(outsidePath, "outside-content");
                options.OutputSummaryName = ".." + Path.DirectorySeparatorChar + "outside-parent.xlsx";

                var ex = Assert.ThrowsException<ArgumentException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options));

                StringAssert.Contains(ex.Message, "纯 .xlsx 文件名");
                Assert.AreEqual("outside-content", File.ReadAllText(outsidePath));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementRejectsInputsChangedAfterPreflight()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                AuthorizePreflight(options);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    workbook.Worksheet("重庆台账").Cell(4, 3).Value = "预检后改名";
                    workbook.Save();
                }

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null));

                StringAssert.Contains(ex.Message, "重新预检");
                Assert.IsFalse(Directory.Exists(options.OutputDirectory));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementRejectsPlannedExistingWorkbookChangedAfterPreflight()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                var plannedPath = Path.Combine(
                    options.OutputDirectory,
                    "2026年代理 - 重庆",
                    "测试负责人",
                    "新增代理.xlsx");
                WriteProxyTemplate(plannedPath);
                AddExistingSplitMonthSheet(plannedPath, "5", "预检时内容");
                AuthorizePreflight(options);
                using (var workbook = new XLWorkbook(plannedPath))
                {
                    workbook.Worksheet("5").Cell("A3").Value = "预检后人工修改";
                    workbook.Save();
                }

                var ex = Assert.ThrowsException<InvalidOperationException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null));

                StringAssert.Contains(ex.Message, "重新预检");
                using (var workbook = new XLWorkbook(plannedPath))
                {
                    Assert.AreEqual("预检后人工修改", workbook.Worksheet("5").Cell("A3").GetFormattedString());
                }
                Assert.IsFalse(File.Exists(ChongqingStage2SummaryWorkbookWriter.PlanOutputPath(options)));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementRequiresDecisionForExistingBlankPaymentParty()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);
                SetMainSummaryPaymentParty(options.SummaryTemplatePath, "新增代理", string.Empty);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                AssertHasPaymentIssue(report, ChongqingStage2SettlementKinds.Proxy, "新增代理");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksInvalidProxyRelationshipBeforeAmountFiltering()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    workbook.Worksheet("重庆台账").Cell(4, 12 + 22).Clear(XLClearOptions.Contents);
                    workbook.Save();
                }

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.RelationshipParametersInvalid
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.LedgerRow == 4));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksParametersWithoutIntermediarySubjectIncludingZero()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    workbook.Worksheet("重庆台账").Cell(4, 12 + 7).Value = 0;
                    workbook.Save();
                }

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.RelationshipParametersWithoutSubject
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.LedgerRow == 4));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementCombinesSameProxySubjectAcrossOwnersUsingFirstPhysicalOwner()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    WriteLedgerRow(
                        workbook.Worksheet("重庆台账"),
                        8,
                        "代理客户二",
                        "新增代理",
                        "代理",
                        "后负责人",
                        null,
                        null,
                        1.2345,
                        0,
                        0);
                    workbook.Save();
                }
                WriteProxyTemplate(Path.Combine(options.ProxyTemplateDirectory, "测试负责人", "新增代理.xlsx"));
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);

                var preflight = AuthorizePreflight(options);
                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null);

                Assert.IsTrue(preflight.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.MultipleOwners
                    && issue.Disposition == Stage2PreflightDisposition.Review));
                Assert.AreEqual(1, report.ProxyGroups);
                var proxy = report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Proxy);
                Assert.AreEqual(2, proxy.Rows);
                Assert.AreEqual("测试负责人", proxy.Owner);
                StringAssert.Contains(proxy.OutputFile, "测试负责人");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksConflictingTaxRatesWithinSummarySubject()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    var ws = workbook.Worksheet("重庆台账");
                    WriteLedgerRow(ws, 8, "代理客户二", "新增代理", "代理", "测试负责人", null, null, 1.2345, 0, 0);
                    ws.Cell(8, 12 + 25).Value = 0.06;
                    workbook.Save();
                }
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingTaxRates
                    && issue.Disposition == Stage2PreflightDisposition.Blocker));
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
                HideOnlyTemplateMonth(proxyTemplate);
                HideOnlyTemplateMonth(refundTemplate);
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);
                AuthorizePreflight(options);

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
                Assert.IsTrue(File.Exists(report.HtmlReportPath));
                Assert.IsTrue(Directory.Exists(report.ProxyOutputDirectory));
                Assert.IsTrue(Directory.Exists(report.RefundOutputDirectory));
                Assert.AreEqual(options.ExpectedPreflightSignature, report.PreflightSignature);
                Assert.AreEqual(options.ExpectedInputFingerprint, report.InputFingerprint);
                var reportJson = File.ReadAllText(report.ReportPath);
                StringAssert.Contains(reportJson, options.ExpectedPreflightSignature);
                StringAssert.Contains(reportJson, options.ExpectedInputFingerprint);
                StringAssert.Contains(reportJson, Path.GetFileName(report.HtmlReportPath));
                var readableReport = File.ReadAllText(report.HtmlReportPath);
                StringAssert.Contains(readableReport, "重庆5月阶段二结算报告");
                StringAssert.Contains(readableReport, "新增代理");
                StringAssert.Contains(readableReport, "新增退补");

                var proxyGroup = report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Proxy);
                var refundGroup = report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Refund);
                Assert.IsTrue(File.Exists(proxyGroup.OutputFile));
                Assert.IsTrue(File.Exists(refundGroup.OutputFile));

                using (var workbook = new XLWorkbook(proxyGroup.OutputFile))
                {
                    var ws = workbook.Worksheet("5");
                    Assert.AreEqual(XLWorksheetVisibility.Visible, ws.Visibility);
                    Assert.IsTrue(ws.TabActive);
                    Assert.IsTrue(ws.TabSelected);
                    Assert.AreEqual(1, workbook.Worksheets.Count(sheet => sheet.TabActive));
                    Assert.AreEqual(1, workbook.Worksheets.Count(sheet => sheet.TabSelected));
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
                    Assert.AreEqual(XLWorksheetVisibility.Visible, ws.Visibility);
                    Assert.IsTrue(ws.TabActive);
                    Assert.IsTrue(ws.TabSelected);
                    Assert.AreEqual(1, workbook.Worksheets.Count(sheet => sheet.TabActive));
                    Assert.AreEqual(1, workbook.Worksheets.Count(sheet => sheet.TabSelected));
                    Assert.AreEqual("名称:新增退补", ws.Cell("A2").GetFormattedString());
                    Assert.AreEqual("退补客户", ws.Cell(5, 2).GetFormattedString());
                    Assert.AreEqual("SUM(D5:G5)", ws.Cell(5, 3).FormulaA1);
                    Assert.AreEqual("ROUND((D5*H5*I5+E5*H5*J5+F5*H5*K5+G5*H5*L5)/10,4)", ws.Cell(5, 13).FormulaA1);
                    Assert.AreEqual("ROUND(M5/1.13*P5,4)", ws.Cell(5, 14).FormulaA1);
                    Assert.AreEqual("M5-N5", ws.Cell(5, 15).FormulaA1);
                }

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var sheetNames = workbook.Worksheets.Select(sheet => sheet.Name).ToList();
                    Assert.AreEqual("汇总表", sheetNames[0]);
                    Assert.AreEqual("清能5月", sheetNames[1]);
                    Assert.AreEqual("清辉5月", sheetNames[2]);

                    var ws = workbook.Worksheet("汇总表");
                    Assert.IsTrue(ws.Row(2).CellsUsed().Any(cell => cell.GetFormattedString().Contains("2026") && cell.GetFormattedString().Contains("5")));
                    AssertPaymentPartyTitleMerged(workbook.Worksheet("清能5月"));
                    AssertPaymentPartyTitleMerged(workbook.Worksheet("清辉5月"));
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

                var report = RunAuthorized(options);

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

                var report = RunAuthorized(options);

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
                    Assert.AreEqual("目标月清能格式", workbook.Worksheet("清能5月").Cell("A3").GetFormattedString());
                    AssertPaymentPartyTitleMerged(workbook.Worksheet("清能5月"));
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

                var preflight = AuthorizePreflight(options);
                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Run(options, null);

                Assert.IsTrue(preflight.Issues.Any(issue =>
                    issue.Code == ChongqingStage2IssueKinds.RefundExtraPowerRows
                    && issue.Disposition == Stage2PreflightDisposition.Review
                    && issue.Entity == "新增退补"
                    && issue.Message.Contains("同步到额外块 C-G")
                    && issue.Suggestion.Contains("汇总表的当月抵扣")));
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

        [TestMethod]
        public void AnalyzeSettlementBlocksTinyTaxRateConflictOnZeroAmountRelationshipRow()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    var ws = workbook.Worksheet("重庆台账");
                    WriteLedgerRow(
                        ws,
                        8,
                        "零金额关系客户",
                        "新增代理",
                        "代理",
                        "测试负责人",
                        null,
                        null,
                        0,
                        0,
                        0);
                    for (var column = 12; column <= 16; column++)
                    {
                        ws.Cell(8, column).Value = 0;
                    }

                    ws.Cell(8, 12 + 25).Value = 0.13005d;
                    ws.Cell(8, 12 + 27).Value = 0;
                    workbook.Save();
                }
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingTaxRates
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.Entity == "新增代理"
                    && issue.CurrentValue.Contains("0.13005")));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementUsesCalculatedZeroWhenLedgerNetIsNonzero()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    var worksheet = workbook.Worksheet("重庆台账");
                    for (var column = 12; column <= 16; column++)
                    {
                        worksheet.Cell(4, column).Value = 0;
                    }

                    worksheet.Cell(4, 12 + 28).Value = 0;
                    workbook.Save();
                }

                var preflight = AuthorizePreflight(options);
                Assert.IsTrue(preflight.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.LedgerAmountDifference
                    && issue.Entity == "新增代理"
                    && issue.CurrentValue.Contains("0")));
                var report = new ChongqingStage2Service(
                    new ClosedXmlSettlementExcelGateway()).Run(options, null);

                var proxy = report.Groups.Single(group => group.Kind == ChongqingStage2SettlementKinds.Proxy);
                Assert.AreEqual(0d, proxy.ExpectedNet, 0.0001d);
                using (var workbook = new XLWorkbook(proxy.OutputFile))
                {
                    Assert.AreEqual(0d, CellNumber(workbook.Worksheet("5").Cell(5, 14)), 0.0001d);
                }

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    var row = FindSummaryEntityRow(main, "新增代理");
                    var monthColumn = ChongqingStage2SummaryWorkbookWriter.FindSummaryMonthBlocks(main, 5).Single();
                    Assert.AreEqual(0d, CellNumber(main.Cell(row, monthColumn)), 0.0001d);
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementAllowsSelfOperatedRowOnlyWhenProxyFieldsAreBlank()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    var selfOperatedRow = workbook.Worksheet("重庆台账").Row(7);
                    selfOperatedRow.Cell(7).Value = "内部项目开发人";
                    workbook.Save();
                }
                WriteSummaryTemplate(options.SummaryTemplatePath);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsFalse(report.Issues.Any(issue =>
                    issue.LedgerRow == 7
                    && issue.Code == Stage2PreflightIssueKinds.RelationshipParametersInvalid));
                Assert.AreEqual(2, report.SubjectCount);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksResidualProxyFieldsOnSelfOperatedRowWithoutCreatingSubject()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    workbook.Worksheet("重庆台账").Cell(7, 12 + 22).Value = 0.5;
                    workbook.Save();
                }
                WriteSummaryTemplate(options.SummaryTemplatePath);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.LedgerRow == 7
                    && issue.Code == Stage2PreflightIssueKinds.RelationshipParametersInvalid
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.Message.Contains("自营")));
                Assert.AreEqual(2, report.SubjectCount);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksIntermediaryFieldsOnSelfOperatedRowWithoutCreatingSubject()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    workbook.Worksheet("重庆台账").Cell(7, 10).Value = "残留居间人";
                    workbook.Save();
                }
                WriteSummaryTemplate(options.SummaryTemplatePath);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.LedgerRow == 7
                    && issue.SettlementKind == ChongqingStage2SettlementKinds.Intermediary
                    && issue.Code == Stage2PreflightIssueKinds.RelationshipParametersInvalid
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.Message.Contains("自营")));
                Assert.AreEqual(2, report.SubjectCount);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementAllowsUnplannedManagedWorkbookWithOldMonthsOnly()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                var stalePath = Path.Combine(
                    options.OutputDirectory,
                    "2026年代理 - 重庆",
                    "旧负责人",
                    "历史代理.xlsx");
                WriteProxyTemplate(stalePath);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsFalse(report.Issues.Any(issue =>
                    string.Equals(issue.TemplateFile, stalePath, StringComparison.OrdinalIgnoreCase)
                    && (issue.Code == Stage2PreflightIssueKinds.UnexpectedTargetMonthWorkbook
                        || issue.Code == Stage2PreflightIssueKinds.ManagedOutputUnreadable)));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksUnplannedOldOwnerWorkbookContainingTargetMonth()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options, proxyTemplateOwner: "旧负责人", proxyFileName: "迁移代理.xlsx");
                var oldPath = Path.Combine(
                    options.OutputDirectory,
                    "2026年代理 - 重庆",
                    "旧负责人",
                    "迁移代理.xlsx");
                WriteProxyTemplate(oldPath);
                AddExistingSplitMonthSheet(oldPath, "5", "旧负责人目标月");

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.UnexpectedTargetMonthWorkbook
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && string.Equals(issue.TemplateFile, oldPath, StringComparison.OrdinalIgnoreCase)));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementReviewsPlannedWorkbookContainingTargetMonthAfterOwnerMigration()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options, proxyTemplateOwner: "旧负责人", proxyFileName: "迁移代理.xlsx");
                var plannedPath = Path.Combine(
                    options.OutputDirectory,
                    "2026年代理 - 重庆",
                    "测试负责人",
                    "迁移代理.xlsx");
                WriteProxyTemplate(plannedPath);
                AddExistingSplitMonthSheet(plannedPath, "5", "本批既有目标月");

                var report = AuthorizePreflight(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PlannedTargetMonthWorkbook
                    && issue.Disposition == Stage2PreflightDisposition.Review
                    && string.Equals(issue.TemplateFile, plannedPath, StringComparison.OrdinalIgnoreCase)));
                Assert.IsFalse(report.Issues.Any(issue =>
                    issue.Disposition == Stage2PreflightDisposition.Blocker
                    && string.Equals(issue.TemplateFile, plannedPath, StringComparison.OrdinalIgnoreCase)));

                var generated = new ChongqingStage2Service(
                    new ClosedXmlSettlementExcelGateway()).Run(options, null);
                Assert.IsTrue(generated.Groups.Any(group =>
                    string.Equals(group.OutputFile, plannedPath, StringComparison.OrdinalIgnoreCase)));
                using (var workbook = new XLWorkbook(plannedPath))
                {
                    Assert.AreNotEqual("本批既有目标月", workbook.Worksheet("5").Cell("A3").GetFormattedString());
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksDistinctSubjectsPlannedToSameWorkbookPath()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                using (var workbook = new XLWorkbook(options.LedgerPath))
                {
                    WriteLedgerRow(
                        workbook.Worksheet("重庆台账"),
                        8,
                        "第二代理客户",
                        "第二代理主体",
                        "代理",
                        "测试负责人",
                        null,
                        null,
                        1.2,
                        0,
                        0);
                    workbook.Save();
                }

                var firstTemplate = Path.Combine(options.ProxyTemplateDirectory, "来源一", "共同.xlsx");
                var secondTemplate = Path.Combine(options.ProxyTemplateDirectory, "来源二", "共同.xlsx");
                WriteProxyTemplate(firstTemplate);
                WriteProxyTemplate(secondTemplate);
                SetSplitTemplateEntity(secondTemplate, ChongqingStage2SettlementKinds.Proxy, "第二代理主体");
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(options.SummaryTemplatePath);

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.PlannedOutputPathConflict
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.CurrentValue.Contains("新增代理")
                    && issue.CurrentValue.Contains("第二代理主体")));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksDuplicateTargetMonthBlocksInSummarySheet()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    WriteSummaryMonthBlock(main, 27, "2026年5月");
                    WriteSummaryMonthBlock(main, 33, "2026年5月");
                    workbook.Save();
                }

                var report = new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.SummaryTargetMonthAlreadyExists
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.SheetName == "汇总表"
                    && issue.CurrentValue.Contains("第27列")
                    && issue.CurrentValue.Contains("第33列")));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementRejectsSummaryTemplateWithoutReliableMainSheet()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    main.Name = "数据";
                    main.Cell("A1").Value = "普通数据";
                    workbook.Save();
                }

                var ex = Assert.ThrowsException<InvalidDataException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options));

                StringAssert.Contains(ex.Message, "主汇总工作表");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementRejectsAmbiguousReliableMainSummarySheets()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    main.Name = "主汇总一";
                    main.CopyTo("主汇总二");
                    workbook.Save();
                }

                var ex = Assert.ThrowsException<InvalidDataException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options));

                StringAssert.Contains(ex.Message, "多个可靠的主汇总工作表");
                StringAssert.Contains(ex.Message, "主汇总一");
                StringAssert.Contains(ex.Message, "主汇总二");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementRejectsExactAndAliasMainSummarySheetsAsAmbiguous()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    workbook.Worksheet("汇总表").CopyTo("历史主汇总");
                    workbook.Save();
                }

                var ex = Assert.ThrowsException<InvalidDataException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options));

                StringAssert.Contains(ex.Message, "多个可靠的主汇总工作表");
                StringAssert.Contains(ex.Message, "汇总表");
                StringAssert.Contains(ex.Message, "历史主汇总");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementRejectsUnsupportedSummaryKind()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    var row = FindSummaryEntityRow(main, "新增代理");
                    main.Cell(row, 3).Value = "未知费用";
                    workbook.Save();
                }

                var ex = Assert.ThrowsException<InvalidDataException>(() =>
                    new ChongqingStage2Service(new ClosedXmlSettlementExcelGateway()).Analyze(options));

                StringAssert.Contains(ex.Message, "未知费用");
                StringAssert.Contains(ex.Message, "费用类型");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementTreatsMultiPersonPayeeAsOpaqueTextAndIgnoresEdgeWhitespace()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                var canonicalPayee = "  \r\n张三、李四\r\n  ";
                SetSummaryPayee(options.SummaryTemplatePath, "汇总表", "新增代理", canonicalPayee);
                SetSummaryPayee(options.SummaryTemplatePath, "清能4月", "新增代理", "\n张三、李四\n");
                string storedCanonicalPayee;
                using (var sourceWorkbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var sourceSheet = sourceWorkbook.Worksheet("汇总表");
                    var sourceRow = FindSummaryEntityRow(sourceSheet, "新增代理");
                    storedCanonicalPayee = sourceSheet.Cell(sourceRow, 6).Value.ToString();
                }
                Assert.IsTrue(Stage2OpaqueText.AreEquivalent(canonicalPayee, storedCanonicalPayee));

                var preflight = AuthorizePreflight(options);

                Assert.IsFalse(preflight.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingPayees
                    && issue.Entity == "新增代理"));
                var report = new ChongqingStage2Service(
                    new ClosedXmlSettlementExcelGateway()).Run(options, null);
                using (var workbook = new XLWorkbook(report.Summary))
                {
                    foreach (var sheetName in new[] { "汇总表", "清能5月" })
                    {
                        var worksheet = workbook.Worksheet(sheetName);
                        var matches = worksheet.RowsUsed()
                            .Where(row => row.Cell(2).GetFormattedString() == "新增代理")
                            .ToList();
                        Assert.AreEqual(1, matches.Count);
                        Assert.AreEqual(storedCanonicalPayee, matches[0].Cell(6).Value.ToString());
                    }
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksConflictingCompletePayeeTexts()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                SetSummaryPayee(options.SummaryTemplatePath, "汇总表", "新增代理", "张三");
                SetSummaryPayee(options.SummaryTemplatePath, "清能4月", "新增代理", "李四");

                var report = new ChongqingStage2Service(
                    new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingPayees
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.Entity == "新增代理"
                    && issue.PreviousValue.Contains("张三")
                    && issue.PreviousValue.Contains("李四")));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementCopiesCanonicalLongTermFieldsFromMainToPaymentSheet()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                PrepareManagedOutputPreflightInputs(options);
                using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
                {
                    var main = workbook.Worksheet("汇总表");
                    var payment = workbook.Worksheet("清能4月");
                    var mainRow = FindSummaryEntityRow(main, "新增代理");
                    var paymentRow = FindSummaryEntityRow(payment, "新增代理");
                    main.Cell(mainRow, 4).Value = "主表委托状态";
                    main.Cell(mainRow, 5).Value = "主表协议字段";
                    main.Cell(mainRow, 7).Value = "主表发票种类";
                    main.Cell(mainRow, 19).Value = 12.34d;
                    main.Cell(mainRow, 22).Value = "3月";
                    main.Cell(mainRow, 26).Value = "主表备注";
                    payment.Cell(paymentRow, 4).Value = "旧委托";
                    payment.Cell(paymentRow, 5).Value = "旧协议";
                    payment.Cell(paymentRow, 7).Value = "旧票种";
                    payment.Cell(paymentRow, 19).Value = 99d;
                    payment.Cell(paymentRow, 22).Value = "旧月份";
                    payment.Cell(paymentRow, 26).Value = "旧备注";
                    workbook.Save();
                }

                var report = RunAuthorized(options);

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var target = workbook.Worksheet("清能5月");
                    var row = FindSummaryEntityRow(target, "新增代理");
                    Assert.AreEqual("主表委托状态", target.Cell(row, 4).GetFormattedString());
                    Assert.AreEqual("主表协议字段", target.Cell(row, 5).GetFormattedString());
                    Assert.AreEqual("主表发票种类", target.Cell(row, 7).GetFormattedString());
                    Assert.AreEqual(12.34d, CellNumber(target.Cell(row, FindHeaderColumn(target, "借支"))), 0.0001d);
                    Assert.AreEqual("3月", target.Cell(row, FindHeaderColumn(target, "借支开始抵扣月份")).GetFormattedString());
                    Assert.AreEqual("主表备注", target.Cell(row, FindHeaderColumn(target, "备注")).GetFormattedString());
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateSettlementWritesConfirmedDefaultsForNewSummarySubjects()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteProxyTemplate(Path.Combine(options.ProxyTemplateDirectory, "测试负责人", "新增代理.xlsx"));
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(options.SummaryTemplatePath);
                options.SummarySubjectDecisions.Add(new ChongqingStage2SummarySubjectDecision
                {
                    SettlementKind = ChongqingStage2SettlementKinds.Proxy,
                    Entity = "新增代理",
                    PaymentParty = ChongqingStage2PaymentParties.Qingneng
                });
                options.SummarySubjectDecisions.Add(new ChongqingStage2SummarySubjectDecision
                {
                    SettlementKind = ChongqingStage2SettlementKinds.Refund,
                    Entity = "新增退补",
                    PaymentParty = ChongqingStage2PaymentParties.Qinghui
                });

                var preflight = AuthorizePreflight(options);
                Assert.IsTrue(preflight.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.NewSummarySubject
                    && issue.Entity == "新增代理"));
                var report = new ChongqingStage2Service(
                    new ClosedXmlSettlementExcelGateway()).Run(options, null);

                using (var workbook = new XLWorkbook(report.Summary))
                {
                    var main = workbook.Worksheet("汇总表");
                    var row = FindSummaryEntityRow(main, "新增代理");
                    Assert.AreEqual("否", main.Cell(row, 4).GetFormattedString());
                    Assert.AreEqual("新增代理", main.Cell(row, 6).GetFormattedString());
                    Assert.AreEqual("平台", main.Cell(row, 7).GetFormattedString());
                    Assert.AreEqual("J" + row + "-I" + row, main.Cell(row, 8).FormulaA1);
                    Assert.AreEqual(0.13d, CellNumber(main.Cell(row, 9)), 0.0000000001d);
                    Assert.AreEqual(0.13d, CellNumber(main.Cell(row, 10)), 0.0000000001d);
                    Assert.AreEqual("测试负责人", main.Cell(row, 11).GetFormattedString());
                    Assert.AreEqual(
                        ChongqingStage2PaymentParties.Qingneng,
                        main.Cell(row, FindHeaderColumn(main, "支付方")).GetFormattedString());
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementReviewsUniqueSameKindBorrowTemplate()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                var proxyTemplate = Path.Combine(options.ProxyTemplateDirectory, "历史负责人", "唯一代理模板.xlsx");
                WriteProxyTemplate(proxyTemplate);
                SetSplitTemplateEntity(proxyTemplate, ChongqingStage2SettlementKinds.Proxy, "其它代理主体");
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);

                var report = new ChongqingStage2Service(
                    new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.BorrowedTemplate
                    && issue.Disposition == Stage2PreflightDisposition.Review
                    && issue.Entity == "新增代理"
                    && issue.Message.Contains("唯一同类模板")));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksWhenNoSameKindBorrowTemplateExists()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);

                var report = new ChongqingStage2Service(
                    new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.TemplateMissing
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.Entity == "新增代理"));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeSettlementBlocksAmbiguousSameKindBorrowTemplates()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
                var first = Path.Combine(options.ProxyTemplateDirectory, "来源一", "代理一.xlsx");
                var second = Path.Combine(options.ProxyTemplateDirectory, "来源二", "代理二.xlsx");
                WriteProxyTemplate(first);
                WriteProxyTemplate(second);
                SetSplitTemplateEntity(first, ChongqingStage2SettlementKinds.Proxy, "其它代理一");
                SetSplitTemplateEntity(second, ChongqingStage2SettlementKinds.Proxy, "其它代理二");
                WriteRefundTemplate(Path.Combine(options.RefundTemplateDirectory, "测试负责人", "新增退补.xlsx"));
                WriteSummaryTemplate(
                    options.SummaryTemplatePath,
                    "新增代理",
                    ChongqingStage2SettlementKinds.Proxy,
                    "新增退补",
                    ChongqingStage2SettlementKinds.Refund);

                var report = new ChongqingStage2Service(
                    new ClosedXmlSettlementExcelGateway()).Analyze(options);

                Assert.IsTrue(report.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.AmbiguousBorrowTemplates
                    && issue.Disposition == Stage2PreflightDisposition.Blocker
                    && issue.Entity == "新增代理"));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static ChongqingStage2PreflightReport AuthorizePreflight(
            ChongqingStage2Options options)
        {
            var report = new ChongqingStage2Service(
                new ClosedXmlSettlementExcelGateway()).Analyze(options);
            options.ExpectedPreflightSignature = report.PreflightSignature;
            options.ExpectedInputFingerprint = report.InputFingerprint;
            return report;
        }

        private static ChongqingStage2Report RunAuthorized(ChongqingStage2Options options)
        {
            AuthorizePreflight(options);
            return new ChongqingStage2Service(
                new ClosedXmlSettlementExcelGateway()).Run(options, null);
        }

        private static void PrepareManagedOutputPreflightInputs(
            ChongqingStage2Options options,
            string proxyTemplateOwner = "测试负责人",
            string proxyFileName = "新增代理.xlsx")
        {
            WriteChongqingStage2Ledger(options.LedgerPath, includeIntermediary: false);
            WriteProxyTemplate(Path.Combine(
                options.ProxyTemplateDirectory,
                proxyTemplateOwner,
                proxyFileName));
            WriteRefundTemplate(Path.Combine(
                options.RefundTemplateDirectory,
                "测试负责人",
                "新增退补.xlsx"));
            WriteSummaryTemplate(
                options.SummaryTemplatePath,
                "新增代理",
                ChongqingStage2SettlementKinds.Proxy,
                "新增退补",
                ChongqingStage2SettlementKinds.Refund);
        }

        private static void SetSplitTemplateEntity(string path, string kind, string entity)
        {
            using (var workbook = new XLWorkbook(path))
            {
                workbook.Worksheet("4").Cell("A2").Value = kind == ChongqingStage2SettlementKinds.Refund
                    ? "名称:" + entity
                    : "代理名称:" + entity;
                workbook.Save();
            }
        }

        private static void SetSummaryPayee(
            string path,
            string sheetName,
            string entity,
            string payee)
        {
            using (var workbook = new XLWorkbook(path))
            {
                var worksheet = workbook.Worksheet(sheetName);
                var row = worksheet.RowsUsed()
                    .Single(item => item.Cell(2).GetFormattedString() == entity)
                    .RowNumber();
                worksheet.Cell(row, 6).Value = payee;
                workbook.Save();
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
                WriteLedgerRow(ws, 7, "自营客户", null, "自营", "测试负责人", null, null, 0, 0, 0);
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
            if (!string.IsNullOrWhiteSpace(intermediaryEntity))
            {
                ws.Cell(row, start + 7).Value = 0.5;
                ws.Cell(row, start + 8).Value = 0.7;
                ws.Cell(row, start + 10).Value = 0.13;
                ws.Cell(row, start + 12).Value = intermediaryNet;
            }
            ws.Cell(row, start + 13).Value = 0.5;
            ws.Cell(row, start + 14).Value = 0.1;
            ws.Cell(row, start + 15).Value = 0.2;
            ws.Cell(row, start + 16).Value = 0.3;
            ws.Cell(row, start + 17).Value = 0.4;
            ws.Cell(row, start + 19).Value = 0.13;
            ws.Cell(row, start + 21).Value = refundNet;
            if (!string.IsNullOrWhiteSpace(proxyEntity)
                && !string.Equals(agentOrSelf, "自营", StringComparison.Ordinal))
            {
                ws.Cell(row, start + 22).Value = 0.5;
                ws.Cell(row, start + 23).Value = 0.8;
                ws.Cell(row, start + 25).Value = 0.13;
                ws.Cell(row, start + 27).Value = proxyNet;
                ws.Cell(row, start + 28).Value = customer == "代理客户" ? 0.3 : 0;
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

        private static void HideOnlyTemplateMonth(string path)
        {
            using (var workbook = new XLWorkbook(path))
            {
                workbook.AddWorksheet("说明");
                workbook.Worksheet("4").Visibility = XLWorksheetVisibility.Hidden;
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
                workbook.Worksheet(targetSheetName).Cell("A3").Value = marker;
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

        private static void SetMainSummaryPaymentParty(string path, string entity, string paymentParty)
        {
            using (var workbook = new XLWorkbook(path))
            {
                var ws = workbook.Worksheet("汇总表");
                var row = ws.RowsUsed().First(item => item.Cell(2).GetFormattedString() == entity).RowNumber();
                ws.Cell(row, 25).Value = paymentParty;
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

        private static void AssertPaymentPartyTitleMerged(IXLWorksheet worksheet)
        {
            var remarkColumn = FindHeaderColumn(worksheet, "备注");
            Assert.IsTrue(worksheet.MergedRanges.Any(range => range.RangeAddress.FirstAddress.RowNumber == 1
                && range.RangeAddress.FirstAddress.ColumnNumber == 1
                && range.RangeAddress.LastAddress.RowNumber == 1
                && range.RangeAddress.LastAddress.ColumnNumber == remarkColumn));
            Assert.AreEqual(XLAlignmentHorizontalValues.Center, worksheet.Cell("A1").Style.Alignment.Horizontal);
        }

        private static int FindSummaryEntityRow(IXLWorksheet worksheet, string entity)
        {
            return worksheet.RowsUsed()
                .Single(row => row.Cell(2).GetFormattedString() == entity)
                .RowNumber();
        }

        private static int FindHeaderColumn(IXLWorksheet worksheet, string header)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var row = 1; row <= 3; row++)
            {
                for (var column = 1; column <= lastColumn; column++)
                {
                    if (worksheet.Cell(row, column).GetFormattedString() == header)
                    {
                        return column;
                    }
                }
            }

            Assert.Fail("未找到表头：" + header);
            return 0;
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
