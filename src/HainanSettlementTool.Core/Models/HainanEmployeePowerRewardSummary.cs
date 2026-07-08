using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanEmployeePowerRewardSummary
    {
        public string ResponsiblePerson { get; set; }
        public int CustomerCount { get; set; }
        public Dictionary<int, double> MonthlyPowers { get; set; } = new Dictionary<int, double>();
        public double TotalPower { get; set; }
        public double RewardAmount { get; set; }
    }
}
