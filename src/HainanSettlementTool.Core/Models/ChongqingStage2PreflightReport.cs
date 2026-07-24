using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Core.Models
{
    public sealed class ChongqingStage2PreflightReport
    {
        public int Month { get; set; }
        public int SubjectCount { get; set; }
        public string PreflightSignature { get; set; }
        public string InputFingerprint { get; set; }
        public List<ChongqingStage2CheckIssue> Issues { get; } = new List<ChongqingStage2CheckIssue>();

        public bool HasIssues
        {
            get { return Issues.Count > 0; }
        }

        public bool RequiresPaymentPartySelection
        {
            get { return Issues.Any(issue => issue.RequiresPaymentPartySelection); }
        }

        public bool RequiresTemplateSelection
        {
            get { return Issues.Any(issue => issue.RequiresTemplateSelection); }
        }

        public bool HasBlockingIssues
        {
            get { return Stage2PreflightPolicy.HasBlockingIssues(Issues); }
        }

        public bool HasRequiredDecisions
        {
            get { return Stage2PreflightPolicy.HasRequiredDecisions(Issues); }
        }

        public bool CanGenerate
        {
            get { return Stage2PreflightPolicy.CanGenerate(Issues); }
        }

        public int BlockerCount
        {
            get { return Stage2PreflightPolicy.Count(Issues, Stage2PreflightDisposition.Blocker); }
        }

        public int RequiredDecisionCount
        {
            get { return Stage2PreflightPolicy.Count(Issues, Stage2PreflightDisposition.RequiredDecision); }
        }

        public int ReviewCount
        {
            get { return Stage2PreflightPolicy.Count(Issues, Stage2PreflightDisposition.Review); }
        }
    }
}
