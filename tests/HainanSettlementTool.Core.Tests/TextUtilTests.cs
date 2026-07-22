using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class TextUtilTests
    {
        [TestMethod]
        public void SafeFileNameRejectsTraversalSegmentsAndWindowsDeviceNames()
        {
            Assert.AreEqual("_", TextUtil.SafeFileName("."));
            Assert.AreEqual("_", TextUtil.SafeFileName(".."));
            Assert.AreEqual("_CON", TextUtil.SafeFileName("CON"));
            Assert.AreEqual("_com1.xlsx", TextUtil.SafeFileName("com1.xlsx"));
        }

        [TestMethod]
        public void SafeFileNameRemovesSeparatorsInvalidCharactersAndTrailingDots()
        {
            var result = TextUtil.SafeFileName(" A/B:*? . ");

            Assert.IsFalse(result.Contains("/"));
            Assert.IsFalse(result.Contains(":"));
            Assert.IsFalse(result.EndsWith("."));
            Assert.IsFalse(result.EndsWith(" "));
        }
    }
}
