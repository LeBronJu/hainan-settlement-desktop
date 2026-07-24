using HainanSettlementTool.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Wpf.Tests
{
    [TestClass]
    public sealed class ProvinceUiProfileTests
    {
        [TestMethod]
        public void GuangdongProfileExposesBothStageOneActions()
        {
            var profile = ProvinceUiProfile.For(ProvinceCode.Guangdong);

            Assert.IsTrue(profile.SupportsStage1CleanPower);
            Assert.IsTrue(profile.SupportsStage1LedgerUpdate);
            Assert.AreEqual("生成本月台账", profile.RunStageOneButtonText);
            Assert.AreEqual("整理本月电量", profile.CleanPowerButtonText);
            Assert.AreEqual("广东阶段一：整理本月电量", profile.StageOneTitle);
            Assert.AreEqual("本月电量处理", profile.StageOneResultLabel);
            Assert.IsFalse(profile.StageOneCaption.Contains("聚合"));
            Assert.IsFalse(profile.StageOneCaption.Contains("副本"));
            StringAssert.Contains(profile.RawDetailLabel, "零售结算明细");
        }
    }
}
