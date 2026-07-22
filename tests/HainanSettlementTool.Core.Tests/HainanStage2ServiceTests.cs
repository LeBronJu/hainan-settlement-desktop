using System;
using System.IO;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class HainanStage2ServiceTests
    {
        [TestMethod]
        public void RunRejectsUnsupportedNewSummarySubjectPaymentParty()
        {
            var root = CreateTempRoot();
            try
            {
                var service = new HainanStage2Service(new FakeHainanStage2Gateway());
                var options = CreateOptions(root);
                options.SummarySubjectDecisions.Add(new HainanStage2SummarySubjectDecision
                {
                    SettlementKind = "代理费",
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
                var gateway = new FakeHainanStage2Gateway();
                var service = new HainanStage2Service(gateway);
                var options = CreateOptions(root);
                options.SummarySubjectDecisions.Add(new HainanStage2SummarySubjectDecision
                {
                    SettlementKind = "居间费",
                    Entity = "新增居间主体",
                    PaymentParty = HainanStage2PaymentParties.Qingneng
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

        [TestMethod]
        public void AnalyzeDoesNotAuthorizeDirectGeneration()
        {
            var root = CreateTempRoot();
            try
            {
                var gateway = new FakeHainanStage2Gateway();
                var service = new HainanStage2Service(gateway);
                var options = CreateOptions(root);
                options.ExpectedPreflightSignature = null;
                options.ExpectedInputFingerprint = null;

                service.Analyze(options);

                Assert.IsNull(options.ExpectedPreflightSignature);
                Assert.IsNull(options.ExpectedInputFingerprint);
                Assert.ThrowsException<InvalidOperationException>(() => service.Run(options, null));
                Assert.AreEqual(0, gateway.GenerateCalls);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static HainanStage2Options CreateOptions(string root)
        {
            var proxyDir = Path.Combine(root, "proxy");
            var intermediaryDir = Path.Combine(root, "intermediary");
            Directory.CreateDirectory(proxyDir);
            Directory.CreateDirectory(intermediaryDir);

            return new HainanStage2Options
            {
                Month = 5,
                LedgerPath = CreateFile(root, "ledger.xlsx"),
                ProxyTemplateDirectory = proxyDir,
                IntermediaryTemplateDirectory = intermediaryDir,
                SummaryTemplatePath = CreateFile(root, "summary.xlsx"),
                OutputDirectory = Path.Combine(root, "out"),
                ExpectedPreflightSignature = "confirmed-preflight",
                ExpectedInputFingerprint = "confirmed-input"
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

        private sealed class FakeHainanStage2Gateway : IHainanStage2ExcelGateway
        {
            public int GenerateCalls { get; private set; }

            public readonly HainanStage2Report Report = new HainanStage2Report
            {
                Summary = "summary.xlsx",
                ReportPath = "report.json"
            };

            public HainanStage2PreflightReport AnalyzeSettlement(HainanStage2Options options)
            {
                return new HainanStage2PreflightReport
                {
                    Month = options.Month,
                    PreflightSignature = "preflight",
                    InputFingerprint = "input"
                };
            }

            public HainanStage2Report GenerateSettlement(HainanStage2Options options)
            {
                GenerateCalls++;
                return Report;
            }
        }
    }
}
