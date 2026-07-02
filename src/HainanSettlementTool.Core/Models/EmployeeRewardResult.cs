using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class EmployeeRewardResult
    {
        public int Year { get; set; }
        public List<int> Months { get; set; } = new List<int>();
        public List<EmployeeRewardDetail> Details { get; set; } = new List<EmployeeRewardDetail>();
        public List<EmployeeRewardSummary> EmployeeSummaries { get; set; } = new List<EmployeeRewardSummary>();
        public Dictionary<int, double> MonthTotals { get; set; } = new Dictionary<int, double>();
        public int TotalCustomers { get; set; }
        public double TotalPower { get; set; }
        public double TotalReward { get; set; }
        public string SummaryPath { get; set; }
        public string ReportPath { get; set; }
        public List<string> PersonalWorkbookPaths { get; set; } = new List<string>();
    }
}
