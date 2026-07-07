using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IHainanStage2ExcelGateway
    {
        Stage2PreflightReport AnalyzeSettlement(Stage2Options options);
        Stage2Report GenerateSettlement(Stage2Options options);
    }
}
