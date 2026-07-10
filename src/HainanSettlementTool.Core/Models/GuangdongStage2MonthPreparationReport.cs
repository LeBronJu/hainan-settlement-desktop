using System.Collections.Generic;
using System.Linq;

namespace HainanSettlementTool.Core.Models
{
    public sealed class GuangdongStage2MonthPreparationReport
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string OutputDirectory { get; set; }
        public string ReportPath { get; set; }
        public string ValidationReportPath { get; set; }
        public List<GuangdongStage2WorkbookResult> Workbooks { get; } = new List<GuangdongStage2WorkbookResult>();

        public int CreatedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.CreateTargetMonth);
        public int NormalizedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.NormalizeExistingTargetMonth);
        public int AlreadyPreparedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.AlreadyPrepared);
        public int SkippedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.Skipped);
        public int FailedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.Failed);
        public int SuccessfulCount => CreatedCount + NormalizedCount + AlreadyPreparedCount;

        public int CountFor(string settlementKind)
        {
            return Workbooks.Count(item => item.SettlementKind == settlementKind
                && item.Action != GuangdongStage2PreparationActions.Skipped
                && item.Action != GuangdongStage2PreparationActions.Failed);
        }
    }
}
