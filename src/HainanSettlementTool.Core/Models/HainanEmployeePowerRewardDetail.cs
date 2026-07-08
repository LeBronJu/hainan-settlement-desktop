using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanEmployeePowerRewardDetail
    {
        public int SourceRow { get; set; }
        public int Sequence { get; set; }
        public string CustomerCode { get; set; }
        public string CustomerName { get; set; }
        public string ContractStartMonth { get; set; }
        public string ProjectDeveloper { get; set; }
        public string AgentType { get; set; }
        public string ResponsiblePerson { get; set; }
        public Dictionary<int, double> MonthlyPowers { get; set; } = new Dictionary<int, double>();
        public double TotalPower { get; set; }
    }
}
