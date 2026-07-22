namespace HainanSettlementTool.Core.Models
{
    public interface IStage2TemplateDecision
    {
        string SettlementKind { get; }
        string Entity { get; }
        string TemplatePath { get; }
    }
}
