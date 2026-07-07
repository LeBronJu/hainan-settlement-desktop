namespace HainanSettlementTool.Core.Models
{
    public enum ProvinceStage1CustomerDecisionKind
    {
        MatchExisting = 0,
        CreateNew = 1,
        SkipWrite = 2
    }

    public sealed class ProvinceStage1CustomerDecision
    {
        public string SourceCustomerName { get; set; }
        public ProvinceStage1CustomerDecisionKind DecisionKind { get; set; }
        public string TargetCustomerName { get; set; }
    }

    public sealed class ProvinceStage1CustomerMatch
    {
        public string SourceCustomerName { get; set; }
        public string TargetCustomerName { get; set; }
    }
}
