using System;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class ChongqingStage2WorkflowPlan
    {
        public ChongqingStage2WorkflowPlan(ChongqingStage2Options options, ChongqingStage2PreflightReport preflight)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (preflight == null)
            {
                throw new ArgumentNullException(nameof(preflight));
            }

            Options = options;
            Preflight = preflight;
        }

        public ChongqingStage2Options Options { get; }

        public ChongqingStage2PreflightReport Preflight { get; }

        public Stage2PreflightEvaluation Evaluation
        {
            get
            {
                return Stage2PreflightPolicy.Evaluate(
                    Preflight.Issues,
                    Options.SummarySubjectDecisions,
                    Options.TemplateDecisions);
            }
        }

        public bool IsBlocked
        {
            get { return Evaluation.HasBlockingIssues || Evaluation.HasInvalidDecisions; }
        }

        public bool CanContinue
        {
            get { return Evaluation.CanContinue; }
        }

        public bool RequiresConfirmation
        {
            get
            {
                var evaluation = Evaluation;
                return evaluation.HasBlockingIssues
                    || evaluation.HasInvalidDecisions
                    || Preflight.Issues.Any(issue => issue.Disposition != Stage2PreflightDisposition.Information);
            }
        }
    }
}
