namespace HainanSettlementTool.Core.Models
{
    public sealed class GroupSettlementTotal
    {
        public string Kind { get; set; }
        public string Owner { get; set; }
        public string Entity { get; set; }
        public string DisplayEntity { get; set; }
        public int Rows { get; set; }
        public double ExpectedNet { get; set; }
        public string OutputFile { get; set; }
    }
}
