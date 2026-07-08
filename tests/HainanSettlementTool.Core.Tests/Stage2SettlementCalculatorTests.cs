using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
    }
}
