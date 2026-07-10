namespace HainanSettlementTool.Core.Models
{
    public sealed class GuangdongStage2MonthPreparationOptions
    {
        public int Year { get; set; } = 2026;
        public int Month { get; set; }
        public string ProxyDirectory { get; set; }
        public string IntermediaryDirectory { get; set; }
        public string RefundDirectory { get; set; }
        public string OutputDirectory { get; set; }
    }
}
