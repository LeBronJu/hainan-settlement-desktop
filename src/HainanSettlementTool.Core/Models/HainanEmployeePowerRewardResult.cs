using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanEmployeePowerRewardResult
    {
        public int Year { get; set; }
        public List<int> Months { get; set; } = new List<int>();
        public List<HainanEmployeePowerRewardDetail> Details { get; set; } = new List<HainanEmployeePowerRewardDetail>();
        public List<HainanEmployeePowerRewardSummary> ResponsiblePersonSummaries { get; set; } = new List<HainanEmployeePowerRewardSummary>();
        public Dictionary<int, double> MonthTotals { get; set; } = new Dictionary<int, double>();
        public int TotalCustomers { get; set; }
        public double TotalPower { get; set; }
        public double TotalReward { get; set; }
        public string SummaryPath { get; set; }
        public string ReportPath { get; set; }
        public List<string> PersonalWorkbookPaths { get; set; } = new List<string>();
    }
}
