using System.Collections.Generic;
using System.Linq;

namespace HainanSettlementTool.Core.Models
{
    public sealed class Stage2PreflightEvaluation
    {
        public int BlockerCount { get; set; }
        public int ReviewCount { get; set; }
        public int InformationCount { get; set; }
        public List<string> InvalidDefinitions { get; } = new List<string>();
        public List<Stage2PaymentPartyDecisionResolution> DecisionResolutions { get; } = new List<Stage2PaymentPartyDecisionResolution>();
        public List<Stage2TemplateDecisionResolution> TemplateDecisionResolutions { get; } = new List<Stage2TemplateDecisionResolution>();

        public bool HasBlockingIssues
        {
            get { return BlockerCount > 0 || InvalidDefinitions.Count > 0; }
        }

        public bool HasOutstandingRequiredDecisions
        {
            get
            {
                return DecisionResolutions.Any(item => item.Status == Stage2PaymentPartyDecisionStatus.Outstanding)
                    || TemplateDecisionResolutions.Any(item => item.Status == Stage2TemplateDecisionStatus.Outstanding);
            }
        }

        public bool HasInvalidDecisions
        {
            get
            {
                return DecisionResolutions.Any(item =>
                    item.Status == Stage2PaymentPartyDecisionStatus.Invalid
                    || item.Status == Stage2PaymentPartyDecisionStatus.Conflicting
                    || item.Status == Stage2PaymentPartyDecisionStatus.Stale)
                    || TemplateDecisionResolutions.Any(item =>
                        item.Status == Stage2TemplateDecisionStatus.Invalid
                        || item.Status == Stage2TemplateDecisionStatus.Conflicting
                        || item.Status == Stage2TemplateDecisionStatus.Stale);
            }
        }

        public bool HasReviewIssues
        {
            get { return ReviewCount > 0; }
        }

        public bool CanContinue
        {
            get
            {
                return !HasBlockingIssues
                    && !HasOutstandingRequiredDecisions
                    && !HasInvalidDecisions;
            }
        }

        public bool RequiresConfirmation
        {
            get { return HasReviewIssues || HasOutstandingRequiredDecisions; }
        }

        public bool CanGenerate(bool reviewConfirmed)
        {
            return CanContinue && (!HasReviewIssues || reviewConfirmed);
        }
    }
}
