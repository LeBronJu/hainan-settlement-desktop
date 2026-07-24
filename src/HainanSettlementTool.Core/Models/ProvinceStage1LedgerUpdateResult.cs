using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class ProvinceStage1LedgerUpdateResult
    {
        public ProvinceCode Province { get; set; }
        public int Month { get; set; }
        public string Unit { get; set; }
        public string LedgerPath { get; set; }
        public string RawDetailPath { get; set; }
        public string OutputPowerWorkbookPath { get; set; }
        public string OutputLedgerPath { get; set; }
        public string ReportPath { get; set; }
        public string HtmlReportPath { get; set; }
        public int LedgerCustomerRows { get; set; }
        public int PowerCustomerRows { get; set; }
        public int MatchedRows { get; set; }
        public int UpdatedPowerRows { get; set; }
        public int ManualMatchedRows { get; set; }
        public int CreatedCustomerRows { get; set; }
        public int SkippedCustomerRows { get; set; }
        public int MultiAccountRows { get; set; }
        public int SkippedRows { get; set; }
        public double TotalPower { get; set; }
        public List<ProvinceStage1CustomerDecision> CustomerDecisions { get; set; } = new List<ProvinceStage1CustomerDecision>();
        public List<ProvinceStage1CustomerMatch> ManualCustomerMatches { get; set; } = new List<ProvinceStage1CustomerMatch>();
        public List<string> Warnings { get; set; } = new List<string>();
        public List<ProvinceStage1LedgerUpdateIssue> Issues { get; set; } = new List<ProvinceStage1LedgerUpdateIssue>();
    }
}
