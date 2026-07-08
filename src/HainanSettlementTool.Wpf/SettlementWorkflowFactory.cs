using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;

namespace HainanSettlementTool.Wpf
{
    internal static class SettlementWorkflowFactory
    {
        public static SettlementWorkflow Create()
        {
            var gateway = new ClosedXmlSettlementExcelGateway();
            return new SettlementWorkflow(
                new HainanStage1Service(gateway),
                new HainanStage2Service(gateway),
                new EmployeeRewardService(gateway),
                new ProvinceStage1Service(gateway),
                new ChongqingStage2Service(gateway));
        }
    }
}
