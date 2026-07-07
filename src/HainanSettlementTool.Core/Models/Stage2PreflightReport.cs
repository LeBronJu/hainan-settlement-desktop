using System.Collections.Generic;
using System.Linq;

namespace HainanSettlementTool.Core.Models
{
    public sealed class Stage2PreflightReport
    {
        public int Month { get; set; }
        public List<Stage2CheckIssue> Issues { get; } = new List<Stage2CheckIssue>();

        public bool HasIssues
        {
            get { return Issues.Count > 0; }
        }

        public bool RequiresPaymentPartySelection
        {
            get { return Issues.Any(issue => issue.RequiresPaymentPartySelection); }
        }
    }
}
