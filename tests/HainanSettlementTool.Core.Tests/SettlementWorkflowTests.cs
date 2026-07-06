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
        public void CleanProvinceStage1PowerDataReturnsSharedSummaryLines()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway();
                var workflow = new SettlementWorkflow(
                    new Stage1Service(gateway),
                    new Stage2Service(gateway),
                    new EmployeeRewardService(gateway),
                    new ProvinceStage1Service(gateway));
                var rawDetailPath = CreateFile(root, "chongqing-raw.xlsx");
                var options = new ProvinceStage1CleanOptions
                {
                    Province = ProvinceCode.Chongqing,
                    RawDetailPath = rawDetailPath,
                    OutputDirectory = Path.Combine(root, "out")
                };

                var result = workflow.CleanProvinceStage1PowerData(options, null);

                Assert.AreSame(gateway.ProvinceStage1CleanResult, result.Report);
                CollectionAssert.AreEqual(
                    new[]
                    {
                        "重庆阶段一电量清洗完成。",
                        "电量处理表：" + gateway.ProvinceStage1CleanResult.OutputWorkbookPath,
                        "报告：" + gateway.ProvinceStage1CleanResult.ReportPath,
                        "客户数量：2，户号数量：3",
                        "合计电量：12.3456 兆瓦时"
                    },
                    result.SummaryLines.ToArray());
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void UpdateProvinceStage1LedgerReturnsSharedSummaryLines()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway();
                var workflow = new SettlementWorkflow(
                    new Stage1Service(gateway),
                    new Stage2Service(gateway),
                    new EmployeeRewardService(gateway),
                    new ProvinceStage1Service(gateway));
                var options = new ProvinceStage1LedgerUpdateOptions
                {
                    Province = ProvinceCode.Chongqing,
                    Month = 5,
                    LedgerPath = CreateFile(root, "ledger.xlsx"),
                    RawDetailPath = CreateFile(root, "raw.xlsx"),
                    OutputDirectory = Path.Combine(root, "out")
                };

                var plan = workflow.PlanProvinceStage1LedgerUpdate(options, null);
                var result = workflow.UpdateProvinceStage1Ledger(options, null);

                Assert.AreSame(gateway.ProvinceStage1LedgerUpdatePlan, plan);
                Assert.AreSame(gateway.ProvinceStage1LedgerUpdateResult, result.Report);
                CollectionAssert.AreEqual(
                    new[]
                    {
                        "重庆阶段一台账更新完成。",
                        "输出台账：" + gateway.ProvinceStage1LedgerUpdateResult.OutputLedgerPath,
                        "报告：" + gateway.ProvinceStage1LedgerUpdateResult.ReportPath,
                        "匹配客户：2，写入电量：2",
                        "多户号提示：1",
                        "合计电量：12.3456 兆瓦时"
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

        [TestMethod]
        public void RunEmployeeRewardReturnsSharedSummaryLines()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway();
                var workflow = CreateWorkflow(gateway);
                var options = CreateEmployeeRewardOptions(root);

                var result = workflow.RunEmployeeReward(options, null);

                Assert.AreEqual(gateway.EmployeeRewardResult.SummaryPath, result.Report.SummaryPath);
                Assert.AreEqual(gateway.EmployeeRewardResult.ReportPath, result.Report.ReportPath);
                CollectionAssert.AreEqual(
                    new[]
                    {
                        "员工电量奖励生成完成。",
                        "奖励总表：" + gateway.EmployeeRewardResult.SummaryPath,
                        "报告：" + gateway.EmployeeRewardResult.ReportPath,
                        "员工确认表：2 个",
                        "电量合计：123.4567 万千瓦时",
                        "奖励金额：123.46 元"
                    },
                    result.SummaryLines.ToArray());
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void CompleteStage2GeneratesWhenPreflightHasNoIssues()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway();
                var workflow = CreateWorkflow(gateway);
                var options = CreateStage2Options(root);

                var plan = workflow.PlanStage2(options);
                var result = workflow.CompleteStage2(plan, confirmed: false, log: null);

                Assert.IsFalse(plan.RequiresConfirmation);
                Assert.IsFalse(result.WasCancelled);
                Assert.AreSame(gateway.Stage2Report, result.Report);
                Assert.AreEqual(1, gateway.GenerateSettlementCalls);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void CompleteStage2GeneratesWhenPreflightHasIssuesAndUserConfirms()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway { AddPreflightIssue = true };
                var workflow = CreateWorkflow(gateway);
                var options = CreateStage2Options(root);

                var plan = workflow.PlanStage2(options);
                var result = workflow.CompleteStage2(plan, confirmed: true, log: null);

                Assert.IsTrue(plan.RequiresConfirmation);
                Assert.IsFalse(result.WasCancelled);
                Assert.AreSame(gateway.Stage2Report, result.Report);
                Assert.AreEqual(1, gateway.GenerateSettlementCalls);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void CompleteStage2DoesNotGenerateWhenPreflightHasIssuesAndUserCancels()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeGateway { AddPreflightIssue = true };
                var workflow = CreateWorkflow(gateway);
                var options = CreateStage2Options(root);

                var plan = workflow.PlanStage2(options);
                var result = workflow.CompleteStage2(plan, confirmed: false, log: null);

                Assert.IsTrue(plan.RequiresConfirmation);
                Assert.IsTrue(result.WasCancelled);
                Assert.IsNull(result.Report);
                Assert.AreEqual(0, gateway.GenerateSettlementCalls);
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
                new Stage2Service(gateway),
                new EmployeeRewardService(gateway));
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

        private static EmployeeRewardOptions CreateEmployeeRewardOptions(string root)
        {
            return new EmployeeRewardOptions
            {
                Year = 2026,
                StartMonth = 1,
                EndMonth = 5,
                LedgerPath = CreateFile(root, "reward-ledger.xlsx"),
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

        private sealed class FakeGateway : IStage1ExcelGateway, IStage2ExcelGateway, IEmployeeRewardExcelGateway, IProvinceStage1ExcelGateway
        {
            public bool AddPreflightIssue { get; set; }

            public int GenerateSettlementCalls { get; private set; }

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

            public readonly EmployeeRewardResult EmployeeRewardResult = new EmployeeRewardResult
            {
                SummaryPath = "employee-reward.xlsx",
                ReportPath = "employee-reward-report.json",
                TotalPower = 123.4567,
                TotalReward = 123.4567,
                PersonalWorkbookPaths = new List<string> { "a.xlsx", "b.xlsx" }
            };

            public readonly ProvinceStage1CleanResult ProvinceStage1CleanResult = new ProvinceStage1CleanResult
            {
                Province = ProvinceCode.Chongqing,
                Unit = "兆瓦时",
                OutputWorkbookPath = "chongqing-power.xlsx",
                ReportPath = "chongqing-power-report.json",
                CustomerRows = 2,
                AccountRows = 3,
                TotalPower = 12.3456
            };

            public readonly ProvinceStage1LedgerUpdatePlan ProvinceStage1LedgerUpdatePlan = new ProvinceStage1LedgerUpdatePlan
            {
                Province = ProvinceCode.Chongqing,
                Month = 5,
                Unit = "兆瓦时",
                LedgerCustomerRows = 2,
                PowerCustomerRows = 2,
                MatchedRows = 2,
                MultiAccountRows = 1
            };

            public readonly ProvinceStage1LedgerUpdateResult ProvinceStage1LedgerUpdateResult = new ProvinceStage1LedgerUpdateResult
            {
                Province = ProvinceCode.Chongqing,
                Month = 5,
                Unit = "兆瓦时",
                OutputLedgerPath = "chongqing-ledger.xlsx",
                ReportPath = "chongqing-ledger-report.json",
                MatchedRows = 2,
                UpdatedPowerRows = 2,
                MultiAccountRows = 1,
                TotalPower = 12.3456
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

            public ProvinceStage1CleanResult CleanPowerData(ProvinceStage1CleanOptions options)
            {
                return ProvinceStage1CleanResult;
            }

            public ProvinceStage1LedgerUpdatePlan PlanLedgerUpdate(ProvinceStage1LedgerUpdateOptions options)
            {
                return ProvinceStage1LedgerUpdatePlan;
            }

            public ProvinceStage1LedgerUpdateResult UpdateLedger(ProvinceStage1LedgerUpdateOptions options)
            {
                return ProvinceStage1LedgerUpdateResult;
            }

            public Stage2PreflightReport AnalyzeSettlement(Stage2Options options)
            {
                var report = new Stage2PreflightReport();
                if (AddPreflightIssue)
                {
                    report.Issues.Add(new Stage2CheckIssue
                    {
                        Severity = "提示",
                        Category = "测试预检"
                    });
                }

                return report;
            }

            public Stage2Report GenerateSettlement(Stage2Options options)
            {
                GenerateSettlementCalls++;
                return Stage2Report;
            }

            public IList<EmployeeRewardLedgerRow> ReadLedgerRows(EmployeeRewardOptions options)
            {
                return new List<EmployeeRewardLedgerRow>
                {
                    new EmployeeRewardLedgerRow
                    {
                        SourceRow = 4,
                        CustomerCode = "001",
                        CustomerName = "客户A",
                        Owner = "员工A",
                        MonthPowers = new Dictionary<int, double> { { 1, 123.4567 } }
                    }
                };
            }

            public EmployeeRewardOutput GenerateWorkbooks(EmployeeRewardOptions options, EmployeeRewardResult result)
            {
                return new EmployeeRewardOutput
                {
                    SummaryPath = EmployeeRewardResult.SummaryPath,
                    ReportPath = EmployeeRewardResult.ReportPath,
                    PersonalWorkbookPaths = EmployeeRewardResult.PersonalWorkbookPaths
                };
            }
        }
    }
}
