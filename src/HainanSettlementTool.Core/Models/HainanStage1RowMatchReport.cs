namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanStage1RowMatchReport
    {
        public string Name { get; set; }
        public int TargetRow { get; set; }
        public int ReferenceRow { get; set; }
        public string Code { get; set; }
        public double Total { get; set; }
    }
}
