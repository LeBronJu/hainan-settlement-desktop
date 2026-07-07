using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IProvinceStage1ExcelGateway
    {
        ProvinceStage1CleanResult CleanPowerData(ProvinceStage1CleanOptions options);
        ProvinceStage1LedgerUpdatePlan PlanLedgerUpdate(ProvinceStage1LedgerUpdateOptions options);
        ProvinceStage1LedgerUpdateResult UpdateLedger(ProvinceStage1LedgerUpdateOptions options);
    }
}
