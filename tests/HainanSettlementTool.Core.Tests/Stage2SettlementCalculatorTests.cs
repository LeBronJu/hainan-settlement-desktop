using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class Stage2SettlementCalculatorTests
    {
        [TestMethod]
        public void CalculateAmountsUsesCurrentStage2RoundingRules()
        {
            var amounts = Stage2SettlementCalculator.CalculateAmounts(
                total: 12.3456,
                ratio: 0.8,
                unitPrice: 12.34,
                taxRate: 0.13);

            Assert.AreEqual(121.8758, amounts.Gross, 0.00001);
            Assert.AreEqual(121.8758, amounts.AdjustedGross, 0.00001);
            Assert.AreEqual(14.0211, amounts.TaxAmount, 0.00001);
            Assert.AreEqual(107.8547, amounts.CalculatedNet, 0.00001);
            Assert.AreEqual(107.8547, amounts.ExpectedNet, 0.00001);
        }

        [TestMethod]
        public void CreateLedgerDifferenceIssueReturnsNullWithinTolerance()
        {
            var row = new DetailSettlementRow
            {
                Customer = "测试客户",
                Owner = "负责人",
                Entity = "代理主体",
                LedgerNet = 10.0000,
                CalculatedNet = 10.00005
            };

            var issue = Stage2SettlementCalculator.CreateLedgerDifferenceIssue(row, "代理", "分表.xlsx", "6月");

            Assert.IsNull(issue);
        }

        [TestMethod]
        public void CreateLedgerDifferenceIssueDescribesLedgerAndCalculatedDifference()
        {
            var row = new DetailSettlementRow
            {
                LedgerRow = 12,
                Customer = "测试客户",
                Owner = "负责人",
                Entity = "代理主体",
                LedgerNet = 10.0000,
                CalculatedNet = 10.2345
            };

            var issue = Stage2SettlementCalculator.CreateLedgerDifferenceIssue(row, "代理", "分表.xlsx", "6月");

            Assert.IsNotNull(issue);
            Assert.AreEqual("错误", issue.Severity);
            Assert.AreEqual("台账与分表金额不一致", issue.Category);
            Assert.AreEqual("代理费", issue.Kind);
            Assert.AreEqual("测试客户", issue.Customer);
            Assert.AreEqual("负责人", issue.Owner);
            Assert.AreEqual("代理主体", issue.Entity);
            Assert.AreEqual(12, issue.LedgerRow);
            Assert.AreEqual("分表.xlsx", issue.TemplateFile);
            Assert.AreEqual("6月", issue.SheetName);
            Assert.AreEqual("台账：10", issue.PreviousValue);
            Assert.AreEqual("分表自算：10.2345", issue.CurrentValue);
            StringAssert.Contains(issue.Message, "差额 0.2345 万元");
            StringAssert.Contains(issue.Suggestion, "台账第12行");
        }
    }
}
