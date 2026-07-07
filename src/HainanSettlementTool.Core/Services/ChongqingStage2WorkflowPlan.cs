using System;
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

        public bool RequiresConfirmation
        {
            get { return Preflight.HasIssues; }
        }
    }
}
