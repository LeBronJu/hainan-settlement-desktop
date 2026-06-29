using System.Collections.Generic;

namespace HainanSettlementTool.Core.Services
{
    public sealed class StageWorkflowResult<TReport>
    {
        public StageWorkflowResult(TReport report, IEnumerable<string> summaryLines)
        {
            Report = report;
            SummaryLines = new List<string>(summaryLines);
        }

        public TReport Report { get; }

        public IReadOnlyList<string> SummaryLines { get; }
    }
}
