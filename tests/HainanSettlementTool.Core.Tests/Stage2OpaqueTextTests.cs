using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class Stage2OpaqueTextTests
    {
        [TestMethod]
        public void ComparisonIgnoresOnlyOuterWhitespaceAndLineEndingEncoding()
        {
            Assert.IsTrue(Stage2OpaqueText.AreEquivalent(
                "  张三、李四\r\n王五\r\n",
                "张三、李四\n王五"));
        }

        [TestMethod]
        public void ComparisonDoesNotReorderOrParseNames()
        {
            Assert.IsFalse(Stage2OpaqueText.AreEquivalent("张三、李四", "李四、张三"));
            Assert.IsFalse(Stage2OpaqueText.AreEquivalent("张三、李四", "张三,李四"));
        }

        [TestMethod]
        public void NormalizationKeepsBodyWhitespace()
        {
            Assert.AreEqual("张三  李四", Stage2OpaqueText.NormalizeForComparison("\r\n张三  李四\t"));
        }
    }
}
