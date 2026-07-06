using System.Collections.Generic;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    public sealed class ClosedXmlStage1ExcelGateway : IStage1ExcelGateway, IStage2ExcelGateway, IEmployeeRewardExcelGateway, IProvinceStage1ExcelGateway
    {
        private readonly PowerWorkbookReader _powerWorkbookReader = new PowerWorkbookReader();
        private readonly RawDetailReader _rawDetailReader = new RawDetailReader();
        private readonly CustomerCodeReader _customerCodeReader = new CustomerCodeReader();
        private readonly LedgerStage1Updater _ledgerUpdater = new LedgerStage1Updater();
        private readonly Stage2SettlementGenerator _stage2Generator = new Stage2SettlementGenerator();
        private readonly EmployeeRewardGenerator _employeeRewardGenerator = new EmployeeRewardGenerator();
        private readonly ChongqingPowerCleanGenerator _chongqingPowerCleanGenerator = new ChongqingPowerCleanGenerator();
        private readonly ChongqingLedgerStage1Updater _chongqingLedgerStage1Updater;

        public ClosedXmlStage1ExcelGateway()
        {
            _chongqingLedgerStage1Updater = new ChongqingLedgerStage1Updater(_chongqingPowerCleanGenerator);
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
            if (options.Province == ProvinceCode.Chongqing)
            {
                return _chongqingPowerCleanGenerator.Generate(options);
            }

            throw new System.NotSupportedException("当前省份暂未接入多省份阶段一电量清洗。");
        }

        public ProvinceStage1LedgerUpdatePlan PlanLedgerUpdate(ProvinceStage1LedgerUpdateOptions options)
        {
            if (options.Province == ProvinceCode.Chongqing)
            {
                return _chongqingLedgerStage1Updater.Plan(options);
            }

            throw new System.NotSupportedException("当前省份暂未接入多省份阶段一台账更新。");
        }

        public ProvinceStage1LedgerUpdateResult UpdateLedger(ProvinceStage1LedgerUpdateOptions options)
        {
            if (options.Province == ProvinceCode.Chongqing)
            {
                return _chongqingLedgerStage1Updater.Update(options);
            }

            throw new System.NotSupportedException("当前省份暂未接入多省份阶段一台账更新。");
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
