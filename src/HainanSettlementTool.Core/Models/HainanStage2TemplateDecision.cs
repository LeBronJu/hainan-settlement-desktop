namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanStage2TemplateDecision : IStage2TemplateDecision
    {
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string TemplatePath { get; set; }
    }
}
