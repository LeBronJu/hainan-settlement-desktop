namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanPowerRow
    {
        public int SourceRow { get; set; }
        public string Name { get; set; }
        public string Key { get; set; }
        public double Total { get; set; }
        public double Sharp { get; set; }
        public double Peak { get; set; }
        public double Flat { get; set; }
        public double Valley { get; set; }
    }
}
