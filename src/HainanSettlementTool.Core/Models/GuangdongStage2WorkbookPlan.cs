namespace HainanSettlementTool.Core.Models
{
    public sealed class GuangdongStage2WorkbookPlan
    {
        public string SettlementKind { get; set; }
        public string SourceRoot { get; set; }
        public string SourcePath { get; set; }
        public string RelativePath { get; set; }
        public string Action { get; set; }
        public string IssueKind { get; set; }
        public string Message { get; set; }
        public string SourceSheetName { get; set; }
        public string TargetSheetName { get; set; }
        public bool TargetSheetExisted { get; set; }
        public bool PowerNeedsClearing { get; set; }
        public bool PeriodNeedsUpdate { get; set; }
        public bool SettlementDateNeedsUpdate { get; set; }
        public bool TotalPowerNeedsReset { get; set; }
        public bool WorksheetViewNeedsUpdate { get; set; }
        public int DetailRowCount { get; set; }

        public bool CanProcess
        {
            get { return Action != GuangdongStage2PreparationActions.Skipped; }
        }
    }
}
