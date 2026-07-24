using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    public sealed class ClosedXmlSettlementExcelGateway : IHainanStage1ExcelGateway, IHainanStage2ExcelGateway, IHainanEmployeePowerRewardExcelGateway, IProvinceStage1ExcelGateway, IChongqingStage2ExcelGateway, IGuangdongStage2MonthPreparationExcelGateway
    {
        private readonly HainanPowerWorkbookReader _hainanPowerWorkbookReader = new HainanPowerWorkbookReader();
        private readonly HainanRawDetailReader _hainanRawDetailReader = new HainanRawDetailReader();
        private readonly HainanCustomerCodeReader _hainanCustomerCodeReader = new HainanCustomerCodeReader();
        private readonly HainanStage1LedgerUpdater _hainanStage1LedgerUpdater = new HainanStage1LedgerUpdater();
        private readonly HainanStage2SettlementGenerator _hainanStage2Generator = new HainanStage2SettlementGenerator();
        private readonly ChongqingStage2SettlementGenerator _chongqingStage2Generator = new ChongqingStage2SettlementGenerator();
        private readonly GuangdongStage2MonthPreparationGenerator _guangdongStage2MonthPreparationGenerator = new GuangdongStage2MonthPreparationGenerator();
        private readonly HainanEmployeePowerRewardGenerator _hainanEmployeePowerRewardGenerator = new HainanEmployeePowerRewardGenerator();
        private readonly Dictionary<ProvinceCode, IProvinceStage1Adapter> _provinceStage1Adapters;

        public ClosedXmlSettlementExcelGateway()
        {
            _provinceStage1Adapters = new IProvinceStage1Adapter[]
            {
                new ChongqingProvinceStage1Adapter(),
                new GuangdongProvinceStage1Adapter()
            }.ToDictionary(adapter => adapter.Province);
        }

        public List<HainanPowerRow> ReadPowerRows(string powerPath)
        {
            return _hainanPowerWorkbookReader.Read(powerPath);
        }

        public List<HainanPowerRow> ReadRawPowerRows(string rawDetailPath)
        {
            return _hainanRawDetailReader.Read(rawDetailPath);
        }

        public Dictionary<string, string> ReadCustomerCodes(string rawDetailPath)
        {
            return _hainanCustomerCodeReader.Read(rawDetailPath);
        }

        public void WritePowerWorkbook(IEnumerable<HainanPowerRow> rows, string outputPath)
        {
            _hainanPowerWorkbookReader.Write(rows, outputPath);
        }

        public HainanStage1Report UpdateLedger(HainanStage1Options options)
        {
            return _hainanStage1LedgerUpdater.Update(options, this);
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

        public GuangdongStage2PreflightReport AnalyzeMonthPreparation(GuangdongStage2MonthPreparationOptions options)
        {
            return _guangdongStage2MonthPreparationGenerator.Analyze(options);
        }

        public GuangdongStage2MonthPreparationReport GenerateMonthPreparation(GuangdongStage2MonthPreparationOptions options)
        {
            return _guangdongStage2MonthPreparationGenerator.Generate(options);
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
