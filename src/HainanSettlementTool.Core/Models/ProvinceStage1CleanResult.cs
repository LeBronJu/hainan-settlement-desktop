using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class ProvinceStage1CleanResult
    {
        public ProvinceCode Province { get; set; }
        public int Month { get; set; }
        public string Unit { get; set; }
        public string RawDetailPath { get; set; }
        public string OutputWorkbookPath { get; set; }
        public string ReportPath { get; set; }
        public string HtmlReportPath { get; set; }
        public string SourceSheetName { get; set; }
        public int RawRows { get; set; }
        public int CustomerRows { get; set; }
        public int AccountRows { get; set; }
        public double TotalPower { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
    }
}
