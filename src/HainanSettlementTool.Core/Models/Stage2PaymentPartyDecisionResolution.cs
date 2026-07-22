namespace HainanSettlementTool.Core.Models
{
    public enum Stage2PaymentPartyDecisionStatus
    {
        Outstanding = 0,
        Resolved = 1,
        Invalid = 2,
        Conflicting = 3,
        Stale = 4
    }

    public sealed class Stage2PaymentPartyDecisionResolution
    {
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string PaymentParty { get; set; }
        public Stage2PaymentPartyDecisionStatus Status { get; set; }
        public string Message { get; set; }
    }
}
