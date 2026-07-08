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
