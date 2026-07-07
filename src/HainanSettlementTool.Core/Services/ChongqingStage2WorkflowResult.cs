using System.Collections.Generic;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class ChongqingStage2WorkflowResult
    {
        private ChongqingStage2WorkflowResult(bool wasCancelled, StageWorkflowResult<ChongqingStage2Report> completed)
        {
            WasCancelled = wasCancelled;
            Completed = completed;
        }

        public bool WasCancelled { get; }

        public StageWorkflowResult<ChongqingStage2Report> Completed { get; }

        public ChongqingStage2Report Report
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

        public static ChongqingStage2WorkflowResult Cancelled()
        {
            return new ChongqingStage2WorkflowResult(true, null);
        }

        public static ChongqingStage2WorkflowResult Complete(StageWorkflowResult<ChongqingStage2Report> completed)
        {
            return new ChongqingStage2WorkflowResult(false, completed);
        }
    }
}
