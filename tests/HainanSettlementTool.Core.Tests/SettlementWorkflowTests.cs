using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class SettlementWorkflowTests
    {
        [TestMethod]
        public void RunStage1ReturnsSharedSummaryLines()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway();
                var workflow = CreateWorkflow(gateway);
                var options = CreateStage1Options(root);

                var result = workflow.RunStage1(options, null);

                Assert.AreSame(gateway.Stage1Report, result.Report);
                CollectionAssert.AreEqual(
                    new[]
                    {
                        "阶段1完成。",
                        "输出台账：" + gateway.Stage1Report.Output,
                        "报告：" + gateway.Stage1Report.ReportPath
                    },
                    result.SummaryLines.ToArray());
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void CleanPowerDataReturnsSharedSummaryLines()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway();
                var workflow = CreateWorkflow(gateway);
                var rawDetailPath = CreateFile(root, "raw.csv");
                var outputPath = Path.Combine(root, "out", "clean.xlsx");

                var result = workflow.CleanPowerData(rawDetailPath, outputPath, null);

                Assert.AreEqual(outputPath, result.Report.OutputPath);
                CollectionAssert.AreEqual(
                    new[]
                    {
                        "电量清洗完成。",
                        "电量处理表：" + outputPath,
                        "客户数量：2，合计电量：12.3456"
                    },
                    result.SummaryLines.ToArray());
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunStage2ReturnsSharedSummaryLines()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway();
                var workflow = CreateWorkflow(gateway);
                var options = CreateStage2Options(root);

                var result = workflow.RunStage2(options, null);

                Assert.AreSame(gateway.Stage2Report, result.Report);
                CollectionAssert.AreEqual(
                    new[]
                    {
                        "阶段2完成。",
                        "汇总表：" + gateway.Stage2Report.Summary,
                        "报告：" + gateway.Stage2Report.ReportPath,
                        "代理费合计：123.4567",
                        "居间费合计：8.9"
                    },
                    result.SummaryLines.ToArray());
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static SettlementWorkflow CreateWorkflow(FakeGateway gateway)
        {
            return new SettlementWorkflow(
                new Stage1Service(gateway),
                new Stage2Service(gateway));
        }

        private static Stage1Options CreateStage1Options(string root)
        {
            return new Stage1Options
            {
                Month = 4,
                BaseLedgerPath = CreateFile(root, "base.xlsx"),
                PowerPath = CreateFile(root, "power.xlsx"),
                OutputDirectory = Path.Combine(root, "out")
            };
        }

        private static Stage2Options CreateStage2Options(string root)
        {
            var proxyDir = Path.Combine(root, "proxy");
            var intermediaryDir = Path.Combine(root, "intermediary");
            Directory.CreateDirectory(proxyDir);
            Directory.CreateDirectory(intermediaryDir);

            return new Stage2Options
            {
                Month = 4,
                LedgerPath = CreateFile(root, "ledger.xlsx"),
                ProxyTemplateDirectory = proxyDir,
                IntermediaryTemplateDirectory = intermediaryDir,
                SummaryTemplatePath = CreateFile(root, "summary.xlsx"),
                OutputDirectory = Path.Combine(root, "out")
            };
        }

        private static string CreateFile(string root, string name)
        {
            Directory.CreateDirectory(root);
            var path = Path.Combine(root, name);
            File.WriteAllText(path, "test");
            return path;
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

        private sealed class FakeGateway : IStage1ExcelGateway, IStage2ExcelGateway
        {
            public readonly Stage1Report Stage1Report = new Stage1Report
            {
                Output = "out-ledger.xlsx",
                ReportPath = "stage1-report.json"
            };

            public readonly Stage2Report Stage2Report = new Stage2Report
            {
                Summary = "summary.xlsx",
                ReportPath = "stage2-report.json",
                ProxyTotal = 123.4567,
                IntermediaryTotal = 8.9
            };

            public List<PowerRow> ReadPowerRows(string powerPath)
            {
                return new List<PowerRow>();
            }

            public List<PowerRow> ReadRawPowerRows(string rawDetailPath)
            {
                return new List<PowerRow>
                {
                    new PowerRow { Name = "A", Key = "A", Total = 10 },
                    new PowerRow { Name = "B", Key = "B", Total = 2.3456 }
                };
            }

            public Dictionary<string, string> ReadCustomerCodes(string rawDetailPath)
            {
                return new Dictionary<string, string>();
            }

            public void WritePowerWorkbook(IEnumerable<PowerRow> rows, string outputPath)
            {
            }

            public Stage1Report UpdateLedger(Stage1Options options)
            {
                return Stage1Report;
            }

            public Stage2PreflightReport AnalyzeSettlement(Stage2Options options)
            {
                return new Stage2PreflightReport();
            }

            public Stage2Report GenerateSettlement(Stage2Options options)
            {
                return Stage2Report;
            }
        }
    }
}
