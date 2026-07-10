using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IGuangdongStage2MonthPreparationExcelGateway
    {
        GuangdongStage2PreflightReport AnalyzeMonthPreparation(GuangdongStage2MonthPreparationOptions options);
        GuangdongStage2MonthPreparationReport GenerateMonthPreparation(GuangdongStage2MonthPreparationOptions options);
    }
}
