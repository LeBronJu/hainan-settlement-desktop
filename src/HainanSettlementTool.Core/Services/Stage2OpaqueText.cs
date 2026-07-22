using System;

namespace HainanSettlementTool.Core.Services
{
    public static class Stage2OpaqueText
    {
        public static string NormalizeForComparison(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            return value
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Trim();
        }

        public static bool AreEquivalent(string left, string right)
        {
            return string.Equals(
                NormalizeForComparison(left),
                NormalizeForComparison(right),
                StringComparison.Ordinal);
        }
    }
}
