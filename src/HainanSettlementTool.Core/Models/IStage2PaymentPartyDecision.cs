namespace HainanSettlementTool.Core.Models
{
    public interface IStage2PaymentPartyDecision
    {
        string SettlementKind { get; }
        string Entity { get; }
        string PaymentParty { get; }
    }
}
