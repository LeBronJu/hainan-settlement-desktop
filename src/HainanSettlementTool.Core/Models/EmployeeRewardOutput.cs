using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class EmployeeRewardOutput
    {
        public string SummaryPath { get; set; }
        public string ReportPath { get; set; }
        public List<string> PersonalWorkbookPaths { get; set; } = new List<string>();
    }
}
