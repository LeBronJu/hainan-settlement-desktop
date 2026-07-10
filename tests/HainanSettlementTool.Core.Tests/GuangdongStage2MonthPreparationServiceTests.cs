using System;
using System.Collections.Generic;
using System.IO;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class GuangdongStage2MonthPreparationServiceTests
    {
        [TestMethod]
        public void AnalyzeRequiresAtLeastOneInputDirectory()
        {
            var service = new GuangdongStage2MonthPreparationService(new FakeGateway());
            var options = new GuangdongStage2MonthPreparationOptions
            {
                Month = 5,
                OutputDirectory = Path.GetTempPath()
            };

            var error = Assert.ThrowsException<ArgumentException>(() => service.Analyze(options));
            StringAssert.Contains(error.Message, "至少选择一个");
        }

        [TestMethod]
        public void AnalyzeRejectsOutputInsideInputDirectory()
        {
            var root = CreateTempRoot();
            try
            {
                var input = Path.Combine(root, "proxy");
                Directory.CreateDirectory(input);
                var service = new GuangdongStage2MonthPreparationService(new FakeGateway());
                var options = new GuangdongStage2MonthPreparationOptions
                {
                    Month = 5,
                    ProxyDirectory = input,
                    OutputDirectory = Path.Combine(input, "output")
                };

                var error = Assert.ThrowsException<ArgumentException>(() => service.Analyze(options));
                StringAssert.Contains(error.Message, "不能位于");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunReturnsGatewayReportAndWritesLogMessages()
        {
            var root = CreateTempRoot();
            try
            {
                var input = Path.Combine(root, "proxy");
                var output = Path.Combine(root, "output");
                Directory.CreateDirectory(input);
                var gateway = new FakeGateway();
                var service = new GuangdongStage2MonthPreparationService(gateway);
                var logs = new List<string>();

                var report = service.Run(new GuangdongStage2MonthPreparationOptions
                {
                    Month = 5,
                    ProxyDirectory = input,
                    OutputDirectory = output
                }, logs.Add);

                Assert.AreSame(gateway.Report, report);
                Assert.IsTrue(logs.Exists(item => item.Contains("广东")));
            }
            finally
            {
                DeleteTempRoot(root);
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

        private sealed class FakeGateway : IGuangdongStage2MonthPreparationExcelGateway
        {
            public readonly GuangdongStage2MonthPreparationReport Report = new GuangdongStage2MonthPreparationReport
            {
                OutputDirectory = "output",
                ReportPath = "report.json"
            };

            public GuangdongStage2PreflightReport AnalyzeMonthPreparation(GuangdongStage2MonthPreparationOptions options)
            {
                return new GuangdongStage2PreflightReport { Year = options.Year, Month = options.Month };
            }

            public GuangdongStage2MonthPreparationReport GenerateMonthPreparation(GuangdongStage2MonthPreparationOptions options)
            {
                return Report;
            }
        }
    }
}
