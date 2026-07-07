using System;
using System.IO;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class ChongqingStage2ServiceTests
    {
        [TestMethod]
        public void RunAllowsMissingIntermediaryTemplateDirectoryForCurrentChongqingBoundary()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeChongqingStage2Gateway();
                var service = new ChongqingStage2Service(gateway);
                var options = CreateOptions(root);

                var report = service.Run(options, null);

                Assert.AreSame(gateway.Report, report);
                Assert.AreEqual(1, gateway.GenerateCalls);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunRejectsUnsupportedNewSummarySubjectPaymentParty()
        {
            var root = CreateTempRoot();
            try
            {
                var service = new ChongqingStage2Service(new FakeChongqingStage2Gateway());
                var options = CreateOptions(root);
                options.SummarySubjectDecisions.Add(new ChongqingStage2SummarySubjectDecision
                {
                    SettlementKind = ChongqingStage2SettlementKinds.Proxy,
                    Entity = "新增代理主体",
                    PaymentParty = "未知支付方"
                });

                var ex = Assert.ThrowsException<ArgumentException>(() => service.Run(options, null));
                StringAssert.Contains(ex.Message, "清能或清辉");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunAcceptsExplicitNewSummarySubjectPaymentParty()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeChongqingStage2Gateway();
                var service = new ChongqingStage2Service(gateway);
                var options = CreateOptions(root);
                options.SummarySubjectDecisions.Add(new ChongqingStage2SummarySubjectDecision
                {
                    SettlementKind = ChongqingStage2SettlementKinds.Refund,
                    Entity = "新增退补主体",
                    PaymentParty = ChongqingStage2PaymentParties.Qingneng
                });

                var report = service.Run(options, null);

                Assert.AreSame(gateway.Report, report);
                Assert.AreEqual(1, gateway.GenerateCalls);
            }
            finally
            {
                DeleteTempRoot(root);
            }
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
                LedgerPath = CreateFile(root, "ledger.xlsx"),
                ProxyTemplateDirectory = proxyDir,
                RefundTemplateDirectory = refundDir,
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

        private sealed class FakeChongqingStage2Gateway : IChongqingStage2ExcelGateway
        {
            public int GenerateCalls { get; private set; }

            public readonly ChongqingStage2Report Report = new ChongqingStage2Report
            {
                Summary = "summary.xlsx",
                ReportPath = "report.json"
            };

            public ChongqingStage2PreflightReport AnalyzeSettlement(ChongqingStage2Options options)
            {
                return new ChongqingStage2PreflightReport { Month = options.Month };
            }

            public ChongqingStage2Report GenerateSettlement(ChongqingStage2Options options)
            {
                GenerateCalls++;
                return Report;
            }
        }
    }
}
