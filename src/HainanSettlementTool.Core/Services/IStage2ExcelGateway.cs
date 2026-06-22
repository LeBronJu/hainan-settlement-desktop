using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IStage2ExcelGateway
    {
        Stage2Report GenerateSettlement(Stage2Options options);
    }
}
