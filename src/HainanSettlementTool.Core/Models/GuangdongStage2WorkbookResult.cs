namespace HainanSettlementTool.Core.Models
{
    public sealed class GuangdongStage2WorkbookResult
    {
        public string SettlementKind { get; set; }
        public string SourcePath { get; set; }
        public string RelativePath { get; set; }
        public string OutputPath { get; set; }
        public string ReviewCopyPath { get; set; }
        public string Action { get; set; }
        public string IssueKind { get; set; }
        public string Message { get; set; }
        public int DetailRowCount { get; set; }
        public bool PowerCleared { get; set; }
        public bool PeriodUpdated { get; set; }
        public bool SettlementDateUpdated { get; set; }
        public bool TotalPowerReset { get; set; }
    }
}
