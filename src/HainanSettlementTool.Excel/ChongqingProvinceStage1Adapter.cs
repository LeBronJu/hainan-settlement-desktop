using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class ChongqingProvinceStage1Adapter : IProvinceStage1Adapter
    {
        private readonly ChongqingPowerCleanGenerator _powerCleanGenerator;
        private readonly ChongqingLedgerStage1Updater _ledgerUpdater;

        public ChongqingProvinceStage1Adapter()
        {
            _powerCleanGenerator = new ChongqingPowerCleanGenerator();
            _ledgerUpdater = new ChongqingLedgerStage1Updater(_powerCleanGenerator);
        }

        public ProvinceCode Province => ProvinceCode.Chongqing;

        public ProvinceStage1CleanResult CleanPowerData(ProvinceStage1CleanOptions options)
        {
            return _powerCleanGenerator.Generate(options);
        }

        public ProvinceStage1LedgerUpdatePlan PlanLedgerUpdate(ProvinceStage1LedgerUpdateOptions options)
        {
            return _ledgerUpdater.Plan(options);
        }

        public ProvinceStage1LedgerUpdateResult UpdateLedger(ProvinceStage1LedgerUpdateOptions options)
        {
            return _ledgerUpdater.Update(options);
        }
    }
}
