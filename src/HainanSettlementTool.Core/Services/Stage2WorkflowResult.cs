using System.Collections.Generic;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class Stage2WorkflowResult
    {
        private Stage2WorkflowResult(bool wasCancelled, StageWorkflowResult<Stage2Report> completed)
        {
            WasCancelled = wasCancelled;
            Completed = completed;
        }

        public bool WasCancelled { get; }

        public StageWorkflowResult<Stage2Report> Completed { get; }

        public Stage2Report Report
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

        public static Stage2WorkflowResult Cancelled()
        {
            return new Stage2WorkflowResult(true, null);
        }

        public static Stage2WorkflowResult Complete(StageWorkflowResult<Stage2Report> completed)
        {
            return new Stage2WorkflowResult(false, completed);
        }
    }
}
