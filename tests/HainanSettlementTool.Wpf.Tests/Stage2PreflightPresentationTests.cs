using System;
using System.Linq;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Wpf.Tests
{
    [TestClass]
    public class Stage2PreflightPresentationTests
    {
        [TestMethod]
        public void CreateHainanCombinesNewSubjectChecksIntoOneActionCard()
        {
            var firstTemplate = @"C:\SyntheticTemplates\负责人甲 - 海南2026\模板代理甲 2026海南.xlsx";
            var secondTemplate = @"C:\SyntheticTemplates\负责人乙 - 海南2026\模板代理乙 2026海南.xlsx";
            var report = new HainanStage2PreflightReport
            {
                Month = 6,
                SubjectCount = 21
            };

            var templateIssue = new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.AmbiguousBorrowTemplates,
                Disposition = Stage2PreflightDisposition.RequiredDecision,
                Category = "新增主体分表模板选择",
                Kind = "代理费",
                SettlementKind = "代理费",
                Entity = "新增代理",
                Owner = "测试负责人",
                LedgerRow = 73,
                CurrentValue = "找到 2 个可借用的同类型模板",
                Message = "需要选择一份历史分表作为新分表的样式模板。",
                Suggestion = "请选择版式最接近的一份历史分表。",
                RequiresTemplateSelection = true
            };
            templateIssue.AvailableTemplateFiles.Add(firstTemplate);
            templateIssue.AvailableTemplateFiles.Add(secondTemplate);
            report.Issues.Add(templateIssue);

            var paymentIssue = new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.PaymentPartyRequired,
                Disposition = Stage2PreflightDisposition.RequiredDecision,
                Category = "新增汇总主体支付方选择",
                Kind = "代理费",
                SettlementKind = "代理费",
                Entity = "新增代理",
                Owner = "测试负责人",
                LedgerRow = 73,
                Message = "新增汇总主体没有可继承的支付方。",
                Suggestion = "请选择清能或清辉。",
                RequiresPaymentPartySelection = true
            };
            paymentIssue.AvailablePaymentParties.Add(HainanStage2PaymentParties.Qingneng);
            paymentIssue.AvailablePaymentParties.Add(HainanStage2PaymentParties.Qinghui);
            report.Issues.Add(paymentIssue);

            report.Issues.Add(new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.NewSummarySubject,
                Disposition = Stage2PreflightDisposition.Review,
                Category = "新增汇总主体默认资料",
                Kind = "代理费",
                SettlementKind = "代理费",
                Entity = "新增代理",
                Owner = "测试负责人",
                LedgerRow = 73,
                TemplateFile = @"C:\SyntheticTemplates\summary.xlsx",
                CurrentValue = "平台、税率0、扣税率13%",
                Message = "汇总表将新增主体并写入已确认的默认资料。",
                Suggestion = "生成后请检查。"
            });

            var evaluation = Stage2PreflightPolicy.Evaluate(
                report.Issues,
                Array.Empty<IStage2PaymentPartyDecision>(),
                Array.Empty<IStage2TemplateDecision>());

            var viewModel = Stage2PreflightPresentationAdapter.CreateHainan(report, evaluation);

            Assert.AreEqual(1, viewModel.IssueGroups.Count);
            Assert.AreEqual(3, viewModel.IssueGroups.Single().Issues.Count);
            Assert.AreEqual("代理费 · 新增代理", viewModel.IssueGroups.Single().Heading);
            Assert.AreEqual("必选主体 1", viewModel.RequiredDecisionCountText);
            StringAssert.Contains(viewModel.SummaryText, "预检涉及 1 个主体");
            Assert.IsFalse(viewModel.IssueRows.Any(row =>
                (row.ContextText ?? string.Empty).Contains("SyntheticTemplates")));

            var templateRow = viewModel.IssueRows.Single(row => row.RequiresTemplateSelection);
            Assert.AreEqual(2, templateRow.TemplateOptions.Count);
            Assert.IsTrue(templateRow.TemplateOptions.All(option =>
                !option.DisplayText.Contains(@"C:\SyntheticTemplates")));
            StringAssert.Contains(templateRow.TechnicalDetailsText, firstTemplate);
            StringAssert.Contains(templateRow.Category, "需要你选择");
            StringAssert.Contains(
                viewModel.IssueRows.Single(row => row.Disposition == Stage2PreflightDisposition.Review).Category,
                "程序将填写");
        }

        [TestMethod]
        public void CreateHainanKeepsSameEntityInDifferentFeeTypesSeparate()
        {
            var report = new HainanStage2PreflightReport
            {
                Month = 6,
                SubjectCount = 2
            };
            report.Issues.Add(ReviewIssue("代理费", "同名主体"));
            report.Issues.Add(ReviewIssue("居间费", "同名主体"));
            var evaluation = Stage2PreflightPolicy.Evaluate(
                report.Issues,
                Array.Empty<IStage2PaymentPartyDecision>());

            var viewModel = Stage2PreflightPresentationAdapter.CreateHainan(report, evaluation);

            Assert.AreEqual(2, viewModel.IssueGroups.Count);
            CollectionAssert.AreEquivalent(
                new[] { "代理费 · 同名主体", "居间费 · 同名主体" },
                viewModel.IssueGroups.Select(group => group.Heading).ToArray());
        }

        private static HainanStage2CheckIssue ReviewIssue(string kind, string entity)
        {
            return new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.NewCustomer,
                Disposition = Stage2PreflightDisposition.Review,
                Category = "新增客户",
                Kind = kind,
                SettlementKind = kind,
                Entity = entity,
                Message = "生成后请复核。"
            };
        }
    }
}
