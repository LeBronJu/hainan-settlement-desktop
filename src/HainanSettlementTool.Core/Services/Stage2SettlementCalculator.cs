using System;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public static class Stage2SettlementCalculator
    {
        public const double AmountTolerance = 0.0001;

        public static Stage2SettlementAmounts CalculateAmounts(
            double total,
            double ratio,
            double unitPrice,
            double taxRate,
            double adjustment = 0d)
        {
            var gross = Round4(total * ratio * unitPrice);
            var adjustedGross = Round4(gross - adjustment);
            var taxAmount = Round4(adjustedGross / 1.13 * taxRate);
            var calculatedNet = Round4(adjustedGross - taxAmount);

            return new Stage2SettlementAmounts
            {
                Gross = gross,
                Adjustment = adjustment,
                AdjustedGross = adjustedGross,
                TaxAmount = taxAmount,
                CalculatedNet = calculatedNet,
                ExpectedNet = calculatedNet
            };
        }

        public static bool HasLedgerDifference(DetailSettlementRow row)
        {
            if (row == null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            return Math.Abs(row.LedgerNet - row.CalculatedNet) > AmountTolerance;
        }

        public static Stage2CheckIssue CreateLedgerDifferenceIssue(
            DetailSettlementRow row,
            string kind,
            string outputPath,
            string sheetName)
        {
            if (!HasLedgerDifference(row))
            {
                return null;
            }

            return new Stage2CheckIssue
            {
                Severity = "错误",
                Category = "台账与分表金额不一致",
                Kind = kind + "费",
                Customer = row.Customer,
                Owner = row.Owner,
                Entity = row.Entity,
                LedgerRow = row.LedgerRow,
                TemplateFile = outputPath,
                SheetName = sheetName,
                PreviousValue = "台账：" + FormatAmount(row.LedgerNet),
                CurrentValue = "分表自算：" + FormatAmount(row.CalculatedNet),
                Message = kind + "费主体“" + row.Entity + "”下的客户“" + row.Customer + "”金额不一致，差额 " + FormatAmount(row.CalculatedNet - row.LedgerNet) + " 万元。",
                Suggestion = "请检查台账第" + row.LedgerRow + "行的电量、比例、利润单价、税率及公式缓存；当前汇总表采用分表自算结果，如果确认台账正确，请同步检查/修改分表和汇总表。"
            };
        }

        public static string FormatAmount(double value)
        {
            return value.ToString("0.####");
        }

        private static double Round4(double value)
        {
            return Math.Round(value, 4);
        }
    }
}
