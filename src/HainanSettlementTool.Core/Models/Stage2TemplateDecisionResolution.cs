namespace HainanSettlementTool.Core.Models
{
    public enum Stage2TemplateDecisionStatus
    {
        Outstanding = 0,
        Resolved = 1,
        Invalid = 2,
        Conflicting = 3,
        Stale = 4
    }

    public sealed class Stage2TemplateDecisionResolution
    {
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string TemplatePath { get; set; }
        public Stage2TemplateDecisionStatus Status { get; set; }
        public string Message { get; set; }
    }
}
