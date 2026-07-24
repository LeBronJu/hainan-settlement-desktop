using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class ChongqingStage2Options
    {
        public int Month { get; set; }
        public string LedgerPath { get; set; }
        public string ProxyTemplateDirectory { get; set; }
        public string IntermediaryTemplateDirectory { get; set; }
        public string RefundTemplateDirectory { get; set; }
        public string SummaryTemplatePath { get; set; }
        public string OutputDirectory { get; set; }
        public string OutputSummaryName { get; set; }
        public string ExpectedPreflightSignature { get; set; }
        public string ExpectedInputFingerprint { get; set; }
        public List<ChongqingStage2SummarySubjectDecision> SummarySubjectDecisions { get; } = new List<ChongqingStage2SummarySubjectDecision>();
        public List<ChongqingStage2TemplateDecision> TemplateDecisions { get; } = new List<ChongqingStage2TemplateDecision>();
    }
}
