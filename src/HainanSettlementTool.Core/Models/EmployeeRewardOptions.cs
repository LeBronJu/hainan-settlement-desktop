namespace HainanSettlementTool.Core.Models
{
    public sealed class EmployeeRewardOptions
    {
        public int Year { get; set; }
        public int StartMonth { get; set; }
        public int EndMonth { get; set; }
        public string LedgerPath { get; set; }
        public string OutputDirectory { get; set; }
        public string OutputSummaryName { get; set; }
    }
}
