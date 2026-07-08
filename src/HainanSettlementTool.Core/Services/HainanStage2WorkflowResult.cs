using System.Collections.Generic;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class HainanStage2WorkflowResult
    {
        private HainanStage2WorkflowResult(bool wasCancelled, StageWorkflowResult<HainanStage2Report> completed)
        {
            WasCancelled = wasCancelled;
            Completed = completed;
        }

        public bool WasCancelled { get; }

        public StageWorkflowResult<HainanStage2Report> Completed { get; }

        public HainanStage2Report Report
        {
            get { return Completed == null ? null : Completed.Report; }
        }

        public IReadOnlyList<string> SummaryLines
        {
            get
            {
                return Completed == null
                    ? new List<string>()
                    : Completed.SummaryLines;
            }
        }

        public static HainanStage2WorkflowResult Cancelled()
        {
            return new HainanStage2WorkflowResult(true, null);
        }

        public static HainanStage2WorkflowResult Complete(StageWorkflowResult<HainanStage2Report> completed)
        {
            return new HainanStage2WorkflowResult(false, completed);
        }
    }
}
