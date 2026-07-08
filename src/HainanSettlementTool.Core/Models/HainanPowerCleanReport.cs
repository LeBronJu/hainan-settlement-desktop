namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanPowerCleanReport
    {
        public string RawDetailPath { get; set; }
        public string OutputPath { get; set; }
        public int RawRows { get; set; }
        public int PowerRows { get; set; }
        public double MonthTotal { get; set; }
    }
}
