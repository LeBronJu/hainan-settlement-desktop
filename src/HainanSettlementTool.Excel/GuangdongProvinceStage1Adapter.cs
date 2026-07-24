using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class GuangdongProvinceStage1Adapter : IProvinceStage1Adapter
    {
        private readonly GuangdongPowerCleanGenerator _powerCleanGenerator;
        private readonly GuangdongLedgerStage1Updater _ledgerUpdater;

        public GuangdongProvinceStage1Adapter()
        {
            var sourceReader = new GuangdongStage1SourceReader();
            _powerCleanGenerator = new GuangdongPowerCleanGenerator(sourceReader);
            _ledgerUpdater = new GuangdongLedgerStage1Updater(_powerCleanGenerator);
        }

        public ProvinceCode Province => ProvinceCode.Guangdong;

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
