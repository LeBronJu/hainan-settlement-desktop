using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public static class Stage2PreflightSignature
    {
        public static string Create(
            int month,
            int subjectCount,
            IEnumerable<Stage2PreflightIssue> issues)
        {
            var signatures = (issues ?? Enumerable.Empty<Stage2PreflightIssue>())
                .Select(IssueSignature)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToList();
            var builder = new StringBuilder();
            Append(builder, month.ToString(CultureInfo.InvariantCulture));
            Append(builder, subjectCount.ToString(CultureInfo.InvariantCulture));
            foreach (var signature in signatures)
            {
                Append(builder, signature);
            }

            return Hash(builder.ToString());
        }

        public static bool Matches(string expected, string actual)
        {
            if (string.IsNullOrWhiteSpace(expected) || string.IsNullOrWhiteSpace(actual))
            {
                return false;
            }

            var left = Encoding.ASCII.GetBytes(expected.Trim());
            var right = Encoding.ASCII.GetBytes(actual.Trim());
            if (left.Length != right.Length)
            {
                return false;
            }

            var difference = 0;
            for (var index = 0; index < left.Length; index++)
            {
                difference |= left[index] ^ right[index];
            }

            return difference == 0;
        }

        private static string IssueSignature(Stage2PreflightIssue issue)
        {
            if (issue == null)
            {
                return "<null>";
            }

            var builder = new StringBuilder();
            Append(builder, issue.Code);
            Append(builder, ((int)issue.Disposition).ToString(CultureInfo.InvariantCulture));
            Append(builder, issue.Severity);
            Append(builder, issue.Category);
            Append(builder, issue.Kind);
            Append(builder, issue.SettlementKind);
            Append(builder, issue.Customer);
            Append(builder, issue.Owner);
            Append(builder, issue.Entity);
            Append(builder, issue.LedgerRow.ToString(CultureInfo.InvariantCulture));
            Append(builder, issue.TemplateFile);
            Append(builder, issue.SheetName);
            Append(builder, issue.PreviousValue);
            Append(builder, issue.CurrentValue);
            Append(builder, issue.Message);
            Append(builder, issue.Suggestion);
            Append(builder, issue.RequiresPaymentPartySelection ? "1" : "0");
            foreach (var option in (issue.PaymentPartyOptions ?? new string[0])
                .OrderBy(value => value, StringComparer.Ordinal))
            {
                Append(builder, option);
            }

            Append(builder, issue.RequiresTemplateSelection ? "1" : "0");
            foreach (var option in (issue.TemplateOptions ?? new string[0])
                .OrderBy(value => value, StringComparer.Ordinal))
            {
                Append(builder, option);
            }

            return builder.ToString();
        }

        private static void Append(StringBuilder builder, string value)
        {
            var text = value ?? string.Empty;
            builder.Append(text.Length.ToString(CultureInfo.InvariantCulture));
            builder.Append(':');
            builder.Append(text);
            builder.Append('|');
        }

        private static string Hash(string value)
        {
            using (var sha256 = SHA256.Create())
            {
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(value));
                return string.Concat(bytes.Select(item => item.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }
    }
}
