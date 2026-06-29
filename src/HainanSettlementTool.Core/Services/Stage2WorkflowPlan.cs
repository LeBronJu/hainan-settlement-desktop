using System;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class Stage2WorkflowPlan
    {
        public Stage2WorkflowPlan(Stage2Options options, Stage2PreflightReport preflight)
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

        public Stage2Options Options { get; }

        public Stage2PreflightReport Preflight { get; }

        public bool RequiresConfirmation
        {
            get { return Preflight.HasIssues; }
        }
    }
}
