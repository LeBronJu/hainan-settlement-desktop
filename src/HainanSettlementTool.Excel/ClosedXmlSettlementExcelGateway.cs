using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    public sealed class ClosedXmlSettlementExcelGateway : IHainanStage1ExcelGateway, IHainanStage2ExcelGateway, IEmployeeRewardExcelGateway, IProvinceStage1ExcelGateway
    {
        private readonly PowerWorkbookReader _powerWorkbookReader = new PowerWorkbookReader();
        private readonly RawDetailReader _rawDetailReader = new RawDetailReader();
        private readonly CustomerCodeReader _customerCodeReader = new CustomerCodeReader();
        private readonly LedgerStage1Updater _ledgerUpdater = new LedgerStage1Updater();
        private readonly Stage2SettlementGenerator _stage2Generator = new Stage2SettlementGenerator();
        private readonly EmployeeRewardGenerator _employeeRewardGenerator = new EmployeeRewardGenerator();
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

        public Stage2Report GenerateSettlement(Stage2Options options)
        {
            return _stage2Generator.Generate(options);
        }

        public Stage2PreflightReport AnalyzeSettlement(Stage2Options options)
        {
            return _stage2Generator.Analyze(options);
        }

        public IList<EmployeeRewardLedgerRow> ReadLedgerRows(EmployeeRewardOptions options)
        {
            return _employeeRewardGenerator.ReadLedgerRows(options);
        }

        public EmployeeRewardOutput GenerateWorkbooks(EmployeeRewardOptions options, EmployeeRewardResult result)
        {
            return _employeeRewardGenerator.GenerateWorkbooks(options, result);
        }
    }
}
