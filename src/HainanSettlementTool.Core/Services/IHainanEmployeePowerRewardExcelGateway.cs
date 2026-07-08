using System.Collections.Generic;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IHainanEmployeePowerRewardExcelGateway
    {
        IList<HainanEmployeePowerRewardLedgerRow> ReadLedgerRows(HainanEmployeePowerRewardOptions options);
        HainanEmployeePowerRewardOutput GenerateWorkbooks(HainanEmployeePowerRewardOptions options, HainanEmployeePowerRewardResult result);
    }
}
