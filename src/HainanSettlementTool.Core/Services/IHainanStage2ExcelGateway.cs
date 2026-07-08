using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IHainanStage2ExcelGateway
    {
        HainanStage2PreflightReport AnalyzeSettlement(HainanStage2Options options);
        HainanStage2Report GenerateSettlement(HainanStage2Options options);
    }
}
