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
    public sealed class HainanEmployeePowerRewardServiceTests
    {
        [TestMethod]
        public void RunAggregatesSelectedMonthsByResponsiblePerson()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeHainanEmployeePowerRewardGateway
                {
                    Rows = new List<HainanEmployeePowerRewardLedgerRow>
                    {
                        CreateRow(4, "001", "客户A", "员工A", 1, 10, 20),
                        CreateRow(5, "002", "客户B", "员工A", 2, 30, 40),
                        CreateRow(6, "003", "客户C", "员工B", 3, 50, 60)
                    }
                };
                var service = new HainanEmployeePowerRewardService(gateway);
                var options = CreateOptions(root, 4, 5);

                var result = service.Run(options, null);

                Assert.AreEqual(3, result.TotalCustomers);
                Assert.AreEqual(2, result.ResponsiblePersonSummaries.Count);
                Assert.AreEqual(210, result.TotalPower);
                Assert.AreEqual(210, result.TotalReward);

                var responsiblePersonA = result.ResponsiblePersonSummaries.Single(row => row.ResponsiblePerson == "员工A");
                Assert.AreEqual(100, responsiblePersonA.TotalPower);
                Assert.AreEqual(100, responsiblePersonA.RewardAmount);
                Assert.AreEqual(40, responsiblePersonA.MonthlyPowers[4]);
                Assert.AreEqual(60, responsiblePersonA.MonthlyPowers[5]);

                CollectionAssert.AreEqual(new[] { 4, 5 }, result.Months.ToArray());
                Assert.AreSame(result, gateway.GeneratedResult);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunStopsWhenLedgerHasBlockingErrors()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeHainanEmployeePowerRewardGateway
                {
                    Rows = new List<HainanEmployeePowerRewardLedgerRow>
                    {
                        CreateRow(4, "001", "客户A", string.Empty, 1, 10),
                        CreateRow(5, "002", string.Empty, "员工A", 2, 20),
                        CreateRow(6, "002", "客户B", "员工B", 3, 30)
                    }
                };
                var service = new HainanEmployeePowerRewardService(gateway);

                var ex = Assert.ThrowsException<InvalidOperationException>(() => service.Run(CreateOptions(root, 4, 4), null));

                StringAssert.Contains(ex.Message, "负责人为空");
                StringAssert.Contains(ex.Message, "企业名称为空但有电量");
                StringAssert.Contains(ex.Message, "客户编号重复");
                Assert.IsFalse(gateway.GenerateCalled);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunAllowsSingleMonthRange()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeHainanEmployeePowerRewardGateway
                {
                    Rows = new List<HainanEmployeePowerRewardLedgerRow>
                    {
                        CreateRow(7, "001", "客户A", "员工A", 1, 10, 20)
                    }
                };
                var service = new HainanEmployeePowerRewardService(gateway);

                var result = service.Run(CreateOptions(root, 5, 5), null);

                CollectionAssert.AreEqual(new[] { 5 }, result.Months.ToArray());
                Assert.AreEqual(20, result.TotalPower);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static HainanEmployeePowerRewardOptions CreateOptions(string root, int startMonth, int endMonth)
        {
            Directory.CreateDirectory(root);
            return new HainanEmployeePowerRewardOptions
            {
                Year = 2026,
                StartMonth = startMonth,
                EndMonth = endMonth,
                LedgerPath = CreateFile(root, "ledger.xlsx"),
                OutputDirectory = Path.Combine(root, "out")
            };
        }

        private static HainanEmployeePowerRewardLedgerRow CreateRow(
            int sourceRow,
            string customerCode,
            string customerName,
            string responsiblePerson,
            int sequence,
            params double[] powers)
        {
            var monthlyPowers = new Dictionary<int, double>();
            for (var index = 0; index < powers.Length; index++)
            {
                monthlyPowers[4 + index] = powers[index];
            }

            return new HainanEmployeePowerRewardLedgerRow
            {
                SourceRow = sourceRow,
                Sequence = sequence,
                CustomerCode = customerCode,
                CustomerName = customerName,
                ContractStartMonth = "202601",
                ProjectDeveloper = "项目开发人",
                AgentType = "代理",
                ResponsiblePerson = responsiblePerson,
                MonthlyPowers = monthlyPowers
            };
        }

        private static string CreateFile(string root, string name)
        {
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

        private sealed class FakeHainanEmployeePowerRewardGateway : IHainanEmployeePowerRewardExcelGateway
        {
            public IList<HainanEmployeePowerRewardLedgerRow> Rows { get; set; } = new List<HainanEmployeePowerRewardLedgerRow>();
            public bool GenerateCalled { get; private set; }
            public HainanEmployeePowerRewardResult GeneratedResult { get; private set; }

            public IList<HainanEmployeePowerRewardLedgerRow> ReadLedgerRows(HainanEmployeePowerRewardOptions options)
            {
                return Rows;
            }

            public HainanEmployeePowerRewardOutput GenerateWorkbooks(HainanEmployeePowerRewardOptions options, HainanEmployeePowerRewardResult result)
            {
                GenerateCalled = true;
                GeneratedResult = result;
                return new HainanEmployeePowerRewardOutput
                {
                    SummaryPath = Path.Combine(options.OutputDirectory, "summary.xlsx"),
                    ReportPath = Path.Combine(options.OutputDirectory, "report.json"),
                    PersonalWorkbookPaths = result.ResponsiblePersonSummaries
                        .Select(row => Path.Combine(options.OutputDirectory, row.ResponsiblePerson + ".xlsx"))
                        .ToList()
                };
            }
        }
    }
}
