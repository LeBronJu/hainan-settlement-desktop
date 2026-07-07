using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class ProvinceStage1LedgerUpdatePlan
    {
        public ProvinceCode Province { get; set; }
        public int Month { get; set; }
        public string Unit { get; set; }
        public int LedgerCustomerRows { get; set; }
        public int PowerCustomerRows { get; set; }
        public int MatchedRows { get; set; }
        public int ManualMatchedRows { get; set; }
        public int CreatedCustomerRows { get; set; }
        public int SkippedCustomerRows { get; set; }
        public int MultiAccountRows { get; set; }
        public int ExistingDifferentPowerRows { get; set; }
        public int MissingInLedgerRows { get; set; }
        public int MissingInPowerRows { get; set; }
        public int AliasCandidateRows { get; set; }
        public List<string> PowerOnlyCustomers { get; set; } = new List<string>();
        public List<string> LedgerOnlyCustomers { get; set; } = new List<string>();
        public List<ProvinceStage1CustomerDecision> CustomerDecisions { get; set; } = new List<ProvinceStage1CustomerDecision>();
        public List<ProvinceStage1CustomerMatch> ManualCustomerMatches { get; set; } = new List<ProvinceStage1CustomerMatch>();
        public List<string> Warnings { get; set; } = new List<string>();
        public List<ProvinceStage1LedgerUpdateIssue> Issues { get; set; } = new List<ProvinceStage1LedgerUpdateIssue>();

        public bool RequiresConfirmation => Issues.Count > 0;
    }
}
