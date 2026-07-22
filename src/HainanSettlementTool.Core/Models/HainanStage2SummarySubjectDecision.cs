namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanStage2SummarySubjectDecision : IStage2PaymentPartyDecision
    {
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string PaymentParty { get; set; }
    }
}
