namespace HainanSettlementTool.Core.Models
{
    public sealed class Stage2SettlementAmounts
    {
        public double Gross { get; set; }
        public double Adjustment { get; set; }
        public double AdjustedGross { get; set; }
        public double TaxAmount { get; set; }
        public double CalculatedNet { get; set; }
        public double ExpectedNet { get; set; }
    }
}
