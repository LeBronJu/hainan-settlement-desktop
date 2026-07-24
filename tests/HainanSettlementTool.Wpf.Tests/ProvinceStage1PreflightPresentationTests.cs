using System.Linq;
using HainanSettlementTool.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Wpf.Tests
{
    [TestClass]
    public sealed class ProvinceStage1PreflightPresentationTests
    {
        [TestMethod]
        public void GuangdongPresentationShowsGroupedFocusAndKeepsBulkDetailsCollapsed()
        {
            var options = new ProvinceStage1LedgerUpdateOptions
            {
                Province = ProvinceCode.Guangdong,
                Month = 6
            };
            var plan = new ProvinceStage1LedgerUpdatePlan
            {
                Province = ProvinceCode.Guangdong,
                Month = 6,
                LedgerCustomerRows = 150,
                PowerCustomerRows = 55,
                MatchedRows = 50,
                CreatedCustomerRows = 5,
                MissingInPowerRows = 100,
                MultiAccountRows = 3
            };
            for (var index = 0; index < 100; index++)
            {
                plan.Issues.Add(Issue(
                    ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower,
                    "台账客户本月无电量",
                    "提示",
                    "无电量客户" + index));
            }

            for (var index = 0; index < 20; index++)
            {
                plan.Issues.Add(Issue(
                    ProvinceStage1LedgerUpdateIssueKinds.ExistingPowerDifference,
                    "已有电量差异",
                    "警告",
                    "旧电量客户" + index));
            }

            for (var index = 0; index < 5; index++)
            {
                plan.Issues.Add(Issue(
                    ProvinceStage1LedgerUpdateIssueKinds.CreatedCustomer,
                    "新增客户到台账",
                    "警告",
                    "新增客户" + index));
            }

            var viewModel = ProvinceStage1PreflightPresentationAdapter.Create(options, plan);

            Assert.AreEqual("生成广东本月台账前确认", viewModel.Title);
            Assert.AreEqual("预检报告", viewModel.Heading);
            Assert.AreEqual("确认并生成本月台账", viewModel.ConfirmButtonText);
            Assert.AreEqual(2, viewModel.FocusGroups.Count);
            Assert.AreEqual(4, viewModel.Metrics.Count);
            CollectionAssert.AreEqual(
                new[] { "电量表客户", "台账已有", "新增客户", "重点检查" },
                viewModel.Metrics.Select(metric => metric.Label).ToArray());
            Assert.AreEqual("2 类", viewModel.Metrics.Last().Value);
            Assert.IsTrue(viewModel.AutomaticItems.Any(item =>
                item.Contains("100 个") && item.Contains("写入 0")));
            Assert.IsTrue(viewModel.AutomaticItems.Any(item => item.Contains("3 个客户有多个计量点")));
            Assert.IsFalse(viewModel.FocusGroups
                .Single(group => group.Kind == ProvinceStage1LedgerUpdateIssueKinds.ExistingPowerDifference)
                .CustomerPreview
                .Contains("旧电量客户19"));
            Assert.AreEqual(125, viewModel.DetailGroups.Sum(group => group.Issues.Count));
            StringAssert.Contains(viewModel.DetailsHeaderText, "125");
            Assert.IsFalse(viewModel.GuidanceText.Contains("预检"));
            Assert.IsFalse(viewModel.GuidanceText.Contains("副本"));
            Assert.IsFalse(viewModel.GuidanceText.Contains("HTML"));
            StringAssert.Contains(viewModel.GuidanceText, "检查报告");
        }

        [TestMethod]
        public void GuangdongPresentationUsesApprovedCoefficientConflictWording()
        {
            var options = new ProvinceStage1LedgerUpdateOptions
            {
                Province = ProvinceCode.Guangdong,
                Month = 6
            };
            var plan = new ProvinceStage1LedgerUpdatePlan
            {
                Province = ProvinceCode.Guangdong,
                Month = 6,
                PowerCustomerRows = 1,
                MatchedRows = 1
            };
            plan.Issues.Add(Issue(
                ProvinceStage1LedgerUpdateIssueKinds.CoefficientConflict,
                "系数不同",
                "警告",
                "测试客户"));

            var viewModel = ProvinceStage1PreflightPresentationAdapter.Create(options, plan);
            var group = viewModel.FocusGroups.Single();

            Assert.AreEqual("多计量点峰平谷系数不同客户", group.Title);
            Assert.AreEqual("不影响代理费结算，但建议检查台账。", group.ActionText);
        }

        private static ProvinceStage1LedgerUpdateIssue Issue(
            string kind,
            string category,
            string severity,
            string customerName)
        {
            return new ProvinceStage1LedgerUpdateIssue
            {
                Kind = kind,
                Category = category,
                Severity = severity,
                CustomerName = customerName,
                Message = "测试说明"
            };
        }
    }
}
