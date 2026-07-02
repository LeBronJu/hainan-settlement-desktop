using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class EmployeeRewardLedgerRow
    {
        public int SourceRow { get; set; }
        public int Sequence { get; set; }
        public string CustomerCode { get; set; }
        public string CustomerName { get; set; }
        public string ContractStartMonth { get; set; }
        public string Developer { get; set; }
        public string AgentType { get; set; }
        public string Owner { get; set; }
        public Dictionary<int, double> MonthPowers { get; set; } = new Dictionary<int, double>();
    }
}
