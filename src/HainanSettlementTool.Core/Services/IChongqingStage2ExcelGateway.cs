using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IChongqingStage2ExcelGateway
    {
        ChongqingStage2PreflightReport AnalyzeSettlement(ChongqingStage2Options options);
        ChongqingStage2Report GenerateSettlement(ChongqingStage2Options options);
    }
}
