using System.Collections.Generic;
using System.Linq;

namespace HainanSettlementTool.Core.Models
{
    public sealed class GuangdongStage2PreflightReport
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public List<GuangdongStage2WorkbookPlan> Workbooks { get; } = new List<GuangdongStage2WorkbookPlan>();

        public int CreateCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.CreateTargetMonth);
        public int NormalizeCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.NormalizeExistingTargetMonth);
        public int AlreadyPreparedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.AlreadyPrepared);
        public int SkippedCount => Workbooks.Count(item => item.Action == GuangdongStage2PreparationActions.Skipped);
        public int ProcessableCount => Workbooks.Count(item => item.CanProcess);
        public bool HasReviewItems => SkippedCount > 0;
    }
}
