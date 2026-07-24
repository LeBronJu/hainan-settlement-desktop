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

        [TestMethod]
        public void CustomerKeyNormalizesCompatibilityWhitespaceAndInvisibleCharacters()
        {
            var key = TextUtil.CustomerKey("Ａ公司\u3000分部\r\n\u00a0\u200b\u2060\ufeff");

            Assert.AreEqual("A公司分部", key);
        }

        [TestMethod]
        public void CustomerKeyKeepsMeaningfulPunctuation()
        {
            Assert.AreNotEqual(
                TextUtil.CustomerKey("广东测试（第一）公司"),
                TextUtil.CustomerKey("广东测试第一公司"));
        }
    }
}
