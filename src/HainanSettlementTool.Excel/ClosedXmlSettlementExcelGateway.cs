using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    public sealed class ClosedXmlSettlementExcelGateway : IHainanStage1ExcelGateway, IHainanStage2ExcelGateway, IHainanEmployeePowerRewardExcelGateway, IProvinceStage1ExcelGateway, IChongqingStage2ExcelGateway
    {
        private readonly PowerWorkbookReader _powerWorkbookReader = new PowerWorkbookReader();
        private readonly RawDetailReader _rawDetailReader = new RawDetailReader();
        private readonly CustomerCodeReader _customerCodeReader = new CustomerCodeReader();
        private readonly LedgerStage1Updater _ledgerUpdater = new LedgerStage1Updater();
        private readonly HainanStage2SettlementGenerator _hainanStage2Generator = new HainanStage2SettlementGenerator();
        private readonly ChongqingStage2SettlementGenerator _chongqingStage2Generator = new ChongqingStage2SettlementGenerator();
        private readonly HainanEmployeePowerRewardGenerator _hainanEmployeePowerRewardGenerator = new HainanEmployeePowerRewardGenerator();
        private readonly Dictionary<ProvinceCode, IProvinceStage1Adapter> _provinceStage1Adapters;

        public ClosedXmlSettlementExcelGateway()
        {
            _provinceStage1Adapters = new IProvinceStage1Adapter[]
            {
                new ChongqingProvinceStage1Adapter()
            }.ToDictionary(adapter => adapter.Province);
        }

        public List<PowerRow> ReadPowerRows(string powerPath)
        {
            return _powerWorkbookReader.Read(powerPath);
        }

        public List<PowerRow> ReadRawPowerRows(string rawDetailPath)
        {
            return _rawDetailReader.Read(rawDetailPath);
        }

        public Dictionary<string, string> ReadCustomerCodes(string rawDetailPath)
        {
            return _customerCodeReader.Read(rawDetailPath);
        }

        public void WritePowerWorkbook(IEnumerable<PowerRow> rows, string outputPath)
        {
            _powerWorkbookReader.Write(rows, outputPath);
        }

        public Stage1Report UpdateLedger(Stage1Options options)
        {
            return _ledgerUpdater.Update(options, this);
        }

        public ProvinceStage1CleanResult CleanPowerData(ProvinceStage1CleanOptions options)
        {
            return ProvinceStage1AdapterFor(options.Province, "电量清洗").CleanPowerData(options);
        }

        public ProvinceStage1LedgerUpdatePlan PlanLedgerUpdate(ProvinceStage1LedgerUpdateOptions options)
        {
            return ProvinceStage1AdapterFor(options.Province, "台账更新").PlanLedgerUpdate(options);
        }

        public ProvinceStage1LedgerUpdateResult UpdateLedger(ProvinceStage1LedgerUpdateOptions options)
        {
            return ProvinceStage1AdapterFor(options.Province, "台账更新").UpdateLedger(options);
        }

        private IProvinceStage1Adapter ProvinceStage1AdapterFor(ProvinceCode province, string actionName)
        {
            IProvinceStage1Adapter adapter;
            if (_provinceStage1Adapters.TryGetValue(province, out adapter))
            {
                return adapter;
            }

            throw new System.NotSupportedException("当前省份暂未接入多省份阶段一" + actionName + "。");
        }

        public HainanStage2Report GenerateSettlement(HainanStage2Options options)
        {
            return _hainanStage2Generator.Generate(options);
        }

        public HainanStage2PreflightReport AnalyzeSettlement(HainanStage2Options options)
        {
            return _hainanStage2Generator.Analyze(options);
        }

        public ChongqingStage2Report GenerateSettlement(ChongqingStage2Options options)
        {
            return _chongqingStage2Generator.Generate(options);
        }

        public ChongqingStage2PreflightReport AnalyzeSettlement(ChongqingStage2Options options)
        {
            return _chongqingStage2Generator.Analyze(options);
        }

        public IList<HainanEmployeePowerRewardLedgerRow> ReadLedgerRows(HainanEmployeePowerRewardOptions options)
        {
            return _hainanEmployeePowerRewardGenerator.ReadLedgerRows(options);
        }

        public HainanEmployeePowerRewardOutput GenerateWorkbooks(HainanEmployeePowerRewardOptions options, HainanEmployeePowerRewardResult result)
        {
            return _hainanEmployeePowerRewardGenerator.GenerateWorkbooks(options, result);
        }
    }
}
