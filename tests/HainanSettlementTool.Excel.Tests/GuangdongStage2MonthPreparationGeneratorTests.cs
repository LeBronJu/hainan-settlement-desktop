using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class GuangdongStage2MonthPreparationGeneratorTests
    {
        [TestMethod]
        public void GenerateCopiesExactPreviousMonthAndPreservesWorkbookStructure()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                var sourcePath = Path.Combine(options.ProxyDirectory, "owner", "proxy.xlsx");
                WriteWorkbook(sourcePath, "代理费用结算单", includeTarget: false, includeStandardSource: true, includeSpecialSheet: true);

                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.AreEqual(1, preflight.CreateCount);
                Assert.AreEqual(0, preflight.SkippedCount);

                var report = service.Run(options, null);
                Assert.AreEqual(1, report.CreatedCount);
                var output = report.Workbooks.Single().OutputPath;
                Assert.IsTrue(File.Exists(output));
                Assert.IsTrue(File.Exists(report.ReportPath));
                Assert.IsTrue(File.Exists(report.ValidationReportPath));

                using (var workbook = new XLWorkbook(output))
                {
                    CollectionAssert.AreEqual(new[] { "4", "4 -2月新增", "5" }, workbook.Worksheets.Select(item => item.Name).ToArray());
                    var source = workbook.Worksheet("4");
                    var special = workbook.Worksheet("4 -2月新增");
                    var target = workbook.Worksheet("5");
                    Assert.AreEqual(100d, source.Cell("C5").GetDouble(), 0.0001);
                    Assert.AreEqual("SPECIAL", special.Cell("R5").GetString());
                    Assert.IsTrue(target.Range("C5:F5").Cells().All(cell => cell.IsEmpty()));
                    Assert.AreEqual("ROUND(C5*I5*J5/10,4)", target.Cell("K5").FormulaA1);
                    StringAssert.Contains(target.Cell("C6").FormulaA1, "C5");
                    Assert.AreEqual("  所属期：2026 年 05 月", target.Cell("F2").GetString());
                    Assert.AreEqual("结算日期：2026 年 06 月 15 日", target.Cell("L2").GetString());
                    Assert.IsTrue(target.MergedRanges.Any(range => range.RangeAddress.ToStringRelative() == "A1:P1"));
                    Assert.IsTrue(target.Column(18).IsHidden);
                    Assert.IsTrue(target.Row(8).IsHidden);
                    Assert.AreEqual(41d, target.Row(5).Height, 0.001);
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateNormalizesExistingTargetWithoutReplacingItsCustomLayout()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                var sourcePath = Path.Combine(options.ProxyDirectory, "proxy.xlsx");
                WriteWorkbook(sourcePath, "代理费用结算单", includeTarget: true, includeStandardSource: false, includeSpecialSheet: false);
                using (var workbook = new XLWorkbook(sourcePath))
                {
                    var target = workbook.Worksheet("5");
                    target.Cell("R5").Value = "TARGET-CUSTOM";
                    target.Cell("C6").Value = 999;
                    target.Cell("D6").Value = 888;
                    target.Cell("E6").Value = 777;
                    target.Cell("F6").Value = 666;
                    workbook.Save();
                }

                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.AreEqual(1, preflight.NormalizeCount);
                Assert.AreEqual(0, preflight.CreateCount);

                var report = service.Run(options, null);
                var item = report.Workbooks.Single();
                Assert.AreEqual(GuangdongStage2PreparationActions.NormalizeExistingTargetMonth, item.Action);
                Assert.IsTrue(item.TotalPowerReset);
                using (var workbook = new XLWorkbook(item.OutputPath))
                {
                    Assert.AreEqual(1, workbook.Worksheets.Count);
                    var target = workbook.Worksheet("5");
                    Assert.AreEqual("TARGET-CUSTOM", target.Cell("R5").GetString());
                    Assert.IsTrue(target.Range("C5:F5").Cells().All(cell => cell.IsEmpty()));
                    Assert.IsTrue(target.Range("C6:F6").Cells().All(cell => cell.GetDouble() == 0));
                    Assert.AreEqual("  所属期：2026 年 05 月", target.Cell("F2").GetString());
                    Assert.AreEqual("结算日期：2026 年 06 月 15 日", target.Cell("L2").GetString());
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateCopiesAlreadyPreparedTargetWithoutRewritingWorkbook()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                var sourcePath = Path.Combine(options.ProxyDirectory, "proxy.xlsx");
                WriteWorkbook(sourcePath, "代理费用结算单", includeTarget: true, includeStandardSource: true, includeSpecialSheet: false);
                using (var workbook = new XLWorkbook(sourcePath))
                {
                    var target = workbook.Worksheet("5");
                    target.Range("C5:F5").Clear(XLClearOptions.Contents);
                    target.Cell("F2").Value = "  所属期：2026 年 05 月";
                    target.Cell("L2").Value = "结算日期：2026 年 06 月 15 日";
                    workbook.Save();
                }

                var sourceHash = FileHash(sourcePath);
                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.AreEqual(1, preflight.AlreadyPreparedCount);

                var report = service.Run(options, null);
                var item = report.Workbooks.Single();
                Assert.AreEqual(GuangdongStage2PreparationActions.AlreadyPrepared, item.Action);
                CollectionAssert.AreEqual(sourceHash, FileHash(item.OutputPath));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void AnalyzeIgnoresEveryNonStandardSheetNameWhenPreviousMonthIsMissing()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                var sourcePath = Path.Combine(options.ProxyDirectory, "proxy.xlsx");
                WriteWorkbook(sourcePath, "代理费用结算单", includeTarget: false, includeStandardSource: false, includeSpecialSheet: true);

                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var report = service.Analyze(options);
                Assert.AreEqual(1, report.SkippedCount);
                Assert.AreEqual(GuangdongStage2IssueKinds.MissingPreviousMonthSheet, report.Workbooks.Single().IssueKind);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GeneratePreservesSkippedWorkbookForManualReviewWithoutModifyingIt()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root, includeAllKinds: true);
                var sourcePath = Path.Combine(
                    options.IntermediaryDirectory,
                    "owner",
                    "【异常】合成居间客户 2026 -广东.xlsx");
                WriteWorkbook(sourcePath, "居间费用结算单", includeTarget: false, includeStandardSource: true, includeSpecialSheet: false);
                using (var workbook = new XLWorkbook(sourcePath))
                {
                    workbook.Worksheet("4").Cell("L2").Value = "结算日期：2026 年 03 月 15 日";
                    workbook.Save();
                }

                var sourceHash = FileHash(sourcePath);
                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var report = service.Run(options, null);

                Assert.AreEqual(1, report.SkippedCount);
                var result = report.Workbooks.Single();
                Assert.AreEqual(GuangdongStage2PreparationActions.Skipped, result.Action);
                Assert.AreEqual(GuangdongStage2IssueKinds.InvalidSettlementDate, result.IssueKind);
                Assert.IsTrue(string.IsNullOrWhiteSpace(result.OutputPath));
                Assert.IsFalse(string.IsNullOrWhiteSpace(result.ReviewCopyPath));
                Assert.IsTrue(File.Exists(result.ReviewCopyPath));
                StringAssert.Contains(result.ReviewCopyPath, Path.Combine("居间", "【未处理-需人工复核】", "owner"));
                CollectionAssert.AreEqual(sourceHash, FileHash(result.ReviewCopyPath));
                using (var preserved = new XLWorkbook(result.ReviewCopyPath))
                {
                    Assert.IsTrue(preserved.TryGetWorksheet("4", out _));
                    Assert.IsFalse(preserved.TryGetWorksheet("5", out _));
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateWritesReadablePartialCompletionReportsForSkippedWorkbook()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root, includeAllKinds: true);
                var fileName = "【异常】合成居间客户 2026 -广东.xlsx";
                var sourcePath = Path.Combine(options.IntermediaryDirectory, "owner", fileName);
                WriteWorkbook(sourcePath, "居间费用结算单", includeTarget: false, includeStandardSource: true, includeSpecialSheet: false);
                using (var workbook = new XLWorkbook(sourcePath))
                {
                    workbook.Worksheet("4").Cell("L2").Value = "结算日期：2026 年 03 月 15 日";
                    workbook.Save();
                }

                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var report = service.Run(options, null);

                Assert.AreEqual(1, report.InputCount);
                Assert.AreEqual(0, report.SuccessfulCount);
                Assert.AreEqual(1, report.PreservedSkippedCount);
                Assert.IsTrue(report.HasReviewItems);
                Assert.IsTrue(File.Exists(report.HtmlReportPath));
                StringAssert.Contains(Path.GetFileName(report.HtmlReportPath), "【必须处理】1个文件未生成新月份sheet");

                var validationText = File.ReadAllText(report.ValidationReportPath);
                StringAssert.Contains(validationText, "本批次状态：部分完成，必须人工复核");
                StringAssert.Contains(validationText, "扫描输入：1");
                StringAssert.Contains(validationText, "正常输出：0");
                StringAssert.Contains(validationText, "未自动处理但原文件已保留：1");
                StringAssert.Contains(validationText, fileName);
                StringAssert.Contains(validationText, "上月结算日期顺延后不是目标结算月份");
                StringAssert.Contains(validationText, report.Workbooks.Single().ReviewCopyPath);

                var html = File.ReadAllText(report.HtmlReportPath);
                StringAssert.Contains(html, "部分完成");
                StringAssert.Contains(html, "必须人工复核");
                StringAssert.Contains(html, fileName);
                StringAssert.Contains(html, "原文件已保留");

                var json = File.ReadAllText(report.ReportPath);
                StringAssert.Contains(json, fileName);
                StringAssert.Contains(json, GuangdongStage2PreparationActions.Skipped);
                StringAssert.Contains(json, "\"ReviewCopyPath\":");
                Assert.IsFalse(json.Contains("\"ReviewCopyPath\": null"));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateMarksBatchCriticalWhenSkippedSourceCannotBePreserved()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                var sourcePath = Path.Combine(options.ProxyDirectory, "locked.xlsx");
                WriteWorkbook(
                    sourcePath,
                    "代理费用结算单",
                    includeTarget: false,
                    includeStandardSource: true,
                    includeSpecialSheet: false);

                GuangdongStage2MonthPreparationReport report;
                using (File.Open(sourcePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                    report = service.Run(options, null);
                }

                Assert.AreEqual(1, report.InputCount);
                Assert.AreEqual(0, report.SuccessfulCount);
                Assert.AreEqual(0, report.SkippedCount);
                Assert.AreEqual(1, report.FailedCount);
                Assert.AreEqual(0, report.PreservedReviewCopyCount);
                Assert.IsTrue(report.IsClassificationComplete);
                Assert.IsFalse(report.HasCompleteOutputSet);
                Assert.IsTrue(report.HasCriticalFailures);
                Assert.IsTrue(report.HasReviewItems);

                var result = report.Workbooks.Single();
                Assert.AreEqual(GuangdongStage2IssueKinds.SkippedWorkbookPreservationFailed, result.IssueKind);
                Assert.IsTrue(string.IsNullOrWhiteSpace(result.OutputPath));
                Assert.IsTrue(string.IsNullOrWhiteSpace(result.ReviewCopyPath));
                StringAssert.Contains(result.Message, "原文件保留失败");

                var validationText = File.ReadAllText(report.ValidationReportPath);
                StringAssert.Contains(validationText, "本批次状态：执行异常，必须处理");
                StringAssert.Contains(validationText, "输入文件保留完整：否");
                var html = File.ReadAllText(report.HtmlReportPath);
                StringAssert.Contains(html, "执行异常，必须处理");
                StringAssert.Contains(html, "原文件未能保留");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateWritesNormalHtmlReportWhenEveryWorkbookCompletes()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteWorkbook(
                    Path.Combine(options.ProxyDirectory, "normal.xlsx"),
                    "代理费用结算单",
                    includeTarget: false,
                    includeStandardSource: true,
                    includeSpecialSheet: false);

                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var report = service.Run(options, null);

                Assert.AreEqual(1, report.SuccessfulCount);
                Assert.AreEqual(0, report.SkippedCount);
                Assert.AreEqual(0, report.FailedCount);
                Assert.IsFalse(report.HasReviewItems);
                Assert.IsTrue(File.Exists(report.HtmlReportPath));
                Assert.IsFalse(Path.GetFileName(report.HtmlReportPath).Contains("【必须处理】"));

                var html = File.ReadAllText(report.HtmlReportPath);
                StringAssert.Contains(html, "全部完成");
                StringAssert.Contains(html, "没有需要人工复核的文件");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateHtmlEncodesDynamicWorkbookNamesAndErrorMessages()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                var fileName = "【异常】合成&客户.xlsx";
                var unsafeTitle = "异常<script>alert(1)</script>&标题";
                WriteWorkbook(
                    Path.Combine(options.ProxyDirectory, fileName),
                    unsafeTitle,
                    includeTarget: false,
                    includeStandardSource: true,
                    includeSpecialSheet: false);

                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var report = service.Run(options, null);

                Assert.AreEqual(1, report.SkippedCount);
                Assert.IsTrue(File.Exists(report.HtmlReportPath));

                var html = File.ReadAllText(report.HtmlReportPath);
                StringAssert.Contains(html, "【异常】合成&amp;客户.xlsx");
                StringAssert.Contains(html, "异常&lt;script&gt;alert(1)&lt;/script&gt;&amp;标题");
                Assert.IsFalse(html.Contains(fileName));
                Assert.IsFalse(html.Contains(unsafeTitle));
                Assert.IsFalse(html.Contains("<script>alert(1)</script>"));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateAccountsForEveryWorkbookInMixedSuccessfulAndSkippedBatch()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root);
                WriteWorkbook(
                    Path.Combine(options.ProxyDirectory, "normal.xlsx"),
                    "代理费用结算单",
                    includeTarget: false,
                    includeStandardSource: true,
                    includeSpecialSheet: false);
                var skippedPath = Path.Combine(options.ProxyDirectory, "skipped.xlsx");
                WriteWorkbook(
                    skippedPath,
                    "代理费用结算单",
                    includeTarget: false,
                    includeStandardSource: true,
                    includeSpecialSheet: false);
                using (var workbook = new XLWorkbook(skippedPath))
                {
                    workbook.Worksheet("4").Cell("L2").Value = "结算日期：2026 年 03 月 15 日";
                    workbook.Save();
                }

                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var report = service.Run(options, null);

                Assert.AreEqual(2, report.InputCount);
                Assert.AreEqual(2, report.ClassifiedCount);
                Assert.AreEqual(2, report.AvailableWorkbookCount);
                Assert.IsTrue(report.IsClassificationComplete);
                Assert.IsTrue(report.HasCompleteOutputSet);
                Assert.AreEqual(1, report.SuccessfulCount);
                Assert.AreEqual(1, report.PreservedSkippedCount);
                Assert.AreEqual(0, report.FailedCount);
                Assert.AreEqual(2, Directory.EnumerateFiles(report.OutputDirectory, "*.xlsx", SearchOption.AllDirectories).Count());
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GenerateRecognizesProxyIntermediaryAndRefundTitles()
        {
            var root = CreateTempRoot();
            try
            {
                var options = CreateOptions(root, includeAllKinds: true);
                WriteWorkbook(Path.Combine(options.ProxyDirectory, "proxy.xlsx"), "代理费用结算单", false, true, false);
                WriteWorkbook(Path.Combine(options.ProxyDirectory, "invalid.xlsx"), "不是广东代理结算单", false, true, false);
                WriteWorkbook(Path.Combine(options.IntermediaryDirectory, "intermediary.xlsx"), "居间费用结算单", false, true, false);
                WriteWorkbook(Path.Combine(options.RefundDirectory, "refund.xlsx"), "退补电费结算单", false, true, false);

                var service = new GuangdongStage2MonthPreparationService(new ClosedXmlSettlementExcelGateway());
                var preflight = service.Analyze(options);
                Assert.AreEqual(3, preflight.CreateCount);
                Assert.AreEqual(1, preflight.SkippedCount);

                var report = service.Run(options, null);
                Assert.AreEqual(3, report.SuccessfulCount);
                Assert.AreEqual(1, report.SkippedCount);
                Assert.AreEqual(1, report.CountFor(GuangdongStage2SettlementKinds.Proxy));
                Assert.AreEqual(1, report.CountFor(GuangdongStage2SettlementKinds.Intermediary));
                Assert.AreEqual(1, report.CountFor(GuangdongStage2SettlementKinds.Refund));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static GuangdongStage2MonthPreparationOptions CreateOptions(string root, bool includeAllKinds = false)
        {
            var proxy = Path.Combine(root, "proxy");
            Directory.CreateDirectory(proxy);
            var options = new GuangdongStage2MonthPreparationOptions
            {
                Year = 2026,
                Month = 5,
                ProxyDirectory = proxy,
                OutputDirectory = Path.Combine(root, "output")
            };
            if (includeAllKinds)
            {
                options.IntermediaryDirectory = Path.Combine(root, "intermediary");
                options.RefundDirectory = Path.Combine(root, "refund");
                Directory.CreateDirectory(options.IntermediaryDirectory);
                Directory.CreateDirectory(options.RefundDirectory);
            }

            return options;
        }

        private static void WriteWorkbook(
            string path,
            string title,
            bool includeTarget,
            bool includeStandardSource,
            bool includeSpecialSheet)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                if (includeStandardSource)
                {
                    WriteSheet(workbook.AddWorksheet("4"), title, "04", "05", false);
                }

                if (includeSpecialSheet)
                {
                    var special = workbook.AddWorksheet("4 -2月新增");
                    WriteSheet(special, title, "04", "05", false);
                    special.Cell("R5").Value = "SPECIAL";
                }

                if (includeTarget)
                {
                    WriteSheet(workbook.AddWorksheet("5"), title, "04", "05", true);
                }

                workbook.SaveAs(path);
            }
        }

        private static void WriteSheet(
            IXLWorksheet worksheet,
            string title,
            string periodMonth,
            string settlementMonth,
            bool customTarget)
        {
            worksheet.Cell("A1").Value = title;
            worksheet.Range("A1:P1").Merge();
            worksheet.Cell("F2").Value = "  所属期：2026 年 " + periodMonth + " 月";
            worksheet.Cell("L2").Value = "结算日期：2026 年 " + settlementMonth + " 月 15 日";
            worksheet.Cell("C4").Value = "总实际电量";
            worksheet.Cell("D4").Value = "峰";
            worksheet.Cell("E4").Value = "平";
            worksheet.Cell("F4").Value = "谷";
            worksheet.Cell("A5").Value = 1;
            worksheet.Cell("B5").Value = "合成客户";
            worksheet.Cell("C5").Value = customTarget ? 200 : 100;
            worksheet.Cell("D5").Value = 30;
            worksheet.Cell("E5").Value = 40;
            worksheet.Cell("F5").Value = 30;
            worksheet.Cell("I5").Value = 0.9;
            worksheet.Cell("J5").Value = 0.012;
            worksheet.Cell("K5").FormulaA1 = "ROUND(C5*I5*J5/10,4)";
            worksheet.Cell("L5").FormulaA1 = "K5-M5-N5";
            worksheet.Cell("O5").FormulaA1 = "L5/1.13*Q5";
            worksheet.Cell("P5").FormulaA1 = "L5-O5";
            worksheet.Cell("Q5").Value = 0.13;
            worksheet.Cell("A6").Value = "合计";
            for (var column = 3; column <= 6; column++)
            {
                var letter = XLHelper.GetColumnLetterFromNumber(column);
                worksheet.Cell(6, column).FormulaA1 = "SUM(" + letter + "5:" + letter + "5)";
            }

            worksheet.Cell("K6").FormulaA1 = "SUM(K5:K5)";
            worksheet.Cell("L6").FormulaA1 = "SUM(L5:L5)";
            worksheet.Cell("O6").FormulaA1 = "SUM(O5:O5)";
            worksheet.Cell("P6").FormulaA1 = "SUM(P5:P5)";
            worksheet.Row(5).Height = 41;
            worksheet.Row(8).Hide();
            worksheet.Column(18).Hide();
        }

        private static byte[] FileHash(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var sha = SHA256.Create())
            {
                return sha.ComputeHash(stream);
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
