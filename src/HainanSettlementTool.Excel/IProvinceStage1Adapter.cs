using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal interface IProvinceStage1Adapter
    {
        ProvinceCode Province { get; }

        ProvinceStage1CleanResult CleanPowerData(ProvinceStage1CleanOptions options);

        ProvinceStage1LedgerUpdatePlan PlanLedgerUpdate(ProvinceStage1LedgerUpdateOptions options);

        ProvinceStage1LedgerUpdateResult UpdateLedger(ProvinceStage1LedgerUpdateOptions options);
    }
}
