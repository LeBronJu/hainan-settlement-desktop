using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class Stage2Options
    {
        public int Month { get; set; }
        public string LedgerPath { get; set; }
        public string ProxyTemplateDirectory { get; set; }
        public string IntermediaryTemplateDirectory { get; set; }
        public string SummaryTemplatePath { get; set; }
        public string OutputDirectory { get; set; }
        public string OutputSummaryName { get; set; }
        public bool AllowMissingOwner { get; set; }
        public List<Stage2SummarySubjectDecision> SummarySubjectDecisions { get; } = new List<Stage2SummarySubjectDecision>();
    }
}
