using System;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class HainanStage2WorkflowPlan
    {
        public HainanStage2WorkflowPlan(HainanStage2Options options, HainanStage2PreflightReport preflight)
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

        public HainanStage2Options Options { get; }

        public HainanStage2PreflightReport Preflight { get; }

        public bool RequiresConfirmation
        {
            get { return Preflight.HasIssues; }
        }
    }
}
