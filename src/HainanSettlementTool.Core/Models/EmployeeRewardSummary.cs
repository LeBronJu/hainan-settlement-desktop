using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class EmployeeRewardSummary
    {
        public string Owner { get; set; }
        public int CustomerCount { get; set; }
        public Dictionary<int, double> MonthPowers { get; set; } = new Dictionary<int, double>();
        public double TotalPower { get; set; }
        public double RewardAmount { get; set; }
    }
}
