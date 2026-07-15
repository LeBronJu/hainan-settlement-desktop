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
        public string HtmlReportPath { get; set; }
        public List<GuangdongStage2WorkbookResult> Workbooks { get; } = new List<GuangdongStage2WorkbookResult>();

        public int InputCount { get; set; }
        public int CreatedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.CreateTargetMonth);
        public int NormalizedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.NormalizeExistingTargetMonth);
        public int AlreadyPreparedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.AlreadyPrepared);
        public int SkippedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.Skipped);
        public int FailedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.Failed);
        public int PreservedSkippedCount => Workbooks.Count(item =>
            item.Action == GuangdongStage2PreparationActions.Skipped
            && !string.IsNullOrWhiteSpace(item.ReviewCopyPath));
        public int PreservedReviewCopyCount => Workbooks.Count(item =>
            !string.IsNullOrWhiteSpace(item.ReviewCopyPath));
        public int SuccessfulCount => CreatedCount + NormalizedCount + AlreadyPreparedCount;
        public int ClassifiedCount => SuccessfulCount + SkippedCount + FailedCount;
        public int AvailableWorkbookCount => Workbooks.Count(item =>
            !string.IsNullOrWhiteSpace(item.OutputPath)
            || !string.IsNullOrWhiteSpace(item.ReviewCopyPath));
        public bool IsClassificationComplete => InputCount == Workbooks.Count
            && InputCount == ClassifiedCount;
        public bool HasCompleteOutputSet => InputCount == AvailableWorkbookCount;
        public bool HasCriticalFailures => FailedCount > 0
            || !IsClassificationComplete
            || !HasCompleteOutputSet;
        public bool HasReviewItems => SkippedCount > 0 || HasCriticalFailures;

        public int CountFor(string settlementKind)
        {
            return Workbooks.Count(item => item.SettlementKind == settlementKind
                && (item.Action == GuangdongStage2PreparationActions.CreateTargetMonth
                    || item.Action == GuangdongStage2PreparationActions.NormalizeExistingTargetMonth
                    || item.Action == GuangdongStage2PreparationActions.AlreadyPrepared));
        }

        public int InputCountFor(string settlementKind)
        {
            return Workbooks.Count(item => item.SettlementKind == settlementKind);
        }

        public int SkippedCountFor(string settlementKind)
        {
            return Workbooks.Count(item => item.SettlementKind == settlementKind
                && item.Action == GuangdongStage2PreparationActions.Skipped);
        }

        public int FailedCountFor(string settlementKind)
        {
            return Workbooks.Count(item => item.SettlementKind == settlementKind
                && item.Action == GuangdongStage2PreparationActions.Failed);
        }
    }
}
