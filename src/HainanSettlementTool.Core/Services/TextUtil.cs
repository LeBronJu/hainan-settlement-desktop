using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace HainanSettlementTool.Core.Services
{
    public static class TextUtil
    {
        public static string S(object value)
        {
            return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
        }

        public static double N(object value)
        {
            if (value == null)
            {
                return 0d;
            }

            if (value is double)
            {
                return (double)value;
            }

            if (value is float)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }

            if (value is decimal)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }

            if (value is int || value is long || value is short)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }

            var text = S(value).Replace(",", string.Empty);
            if (text.Length == 0 || text.StartsWith("=", StringComparison.Ordinal))
            {
                return 0d;
            }

            double parsed;
            return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed) ? parsed : 0d;
        }

        public static string CustomerKey(object value)
        {
            var normalized = S(value).Normalize(NormalizationForm.FormKC);
            var result = new StringBuilder(normalized.Length);
            foreach (var character in normalized)
            {
                if (char.IsWhiteSpace(character)
                    || CharUnicodeInfo.GetUnicodeCategory(character) == UnicodeCategory.Format)
                {
                    continue;
                }

                result.Append(character);
            }

            return result.ToString();
        }

        public static string SafeFileName(string value)
        {
            var invalid = new string(System.IO.Path.GetInvalidFileNameChars());
            var result = Regex.Replace(S(value), "[" + Regex.Escape(invalid) + "]", "_")
                .TrimEnd(' ', '.');
            if (string.IsNullOrWhiteSpace(result) || result == "." || result == "..")
            {
                return "_";
            }

            var stem = result.Split('.')[0];
            if (Regex.IsMatch(
                stem,
                "^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return "_" + result;
            }

            return result;
        }
    }
}
