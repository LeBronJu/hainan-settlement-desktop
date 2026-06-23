using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class Stage2Report
    {
        public int Month { get; set; }
        public string Ledger { get; set; }
        public string ProxyTemplateDirectory { get; set; }
        public string IntermediaryTemplateDirectory { get; set; }
        public string SummaryTemplate { get; set; }
        public string OutputDirectory { get; set; }
        public string Summary { get; set; }
        public string ReportPath { get; set; }
        public int ProxyRows { get; set; }
        public int IntermediaryRows { get; set; }
        public int ProxyGroups { get; set; }
        public int IntermediaryGroups { get; set; }
        public double ProxyTotal { get; set; }
        public double IntermediaryTotal { get; set; }

        public List<string> MissingOwners { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();
        public List<Stage2CheckIssue> AuditIssues { get; } = new List<Stage2CheckIssue>();
        public List<GroupSettlementTotal> Groups { get; } = new List<GroupSettlementTotal>();
    }
}
