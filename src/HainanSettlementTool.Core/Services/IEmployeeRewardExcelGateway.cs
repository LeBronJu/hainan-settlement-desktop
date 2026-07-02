using System.Collections.Generic;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IEmployeeRewardExcelGateway
    {
        IList<EmployeeRewardLedgerRow> ReadLedgerRows(EmployeeRewardOptions options);
        EmployeeRewardOutput GenerateWorkbooks(EmployeeRewardOptions options, EmployeeRewardResult result);
    }
}
