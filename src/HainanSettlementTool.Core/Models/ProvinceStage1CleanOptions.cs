namespace HainanSettlementTool.Core.Models
{
    public sealed class ProvinceStage1CleanOptions
    {
        public ProvinceCode Province { get; set; }
        public int Month { get; set; }
        public string RawDetailPath { get; set; }
        public string OutputDirectory { get; set; }
        public string OutputWorkbookName { get; set; }
    }
}
