using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class HainanStage1Report
    {
        public int Month { get; set; }
        public string SourceLedger { get; set; }
        public string SourcePower { get; set; }
        public string RawDetailForCodes { get; set; }
        public string ReferenceLedger { get; set; }
        public string Output { get; set; }
        public string ReportPath { get; set; }
        public string TargetBlock { get; set; }
        public bool TargetMonthAlreadyPresent { get; set; }
        public int PowerRows { get; set; }
        public int MatchedRows { get; set; }
        public int NewRows { get; set; }
        public int RawDetailCodeRows { get; set; }
        public double MonthTotal { get; set; }

        public List<HainanStage1RowMatchReport> NewCustomers { get; } = new List<HainanStage1RowMatchReport>();
        public List<HainanStage1RowMatchReport> MatchedCustomers { get; } = new List<HainanStage1RowMatchReport>();
        public List<HainanStage1RowMatchReport> CopiedFromReference { get; } = new List<HainanStage1RowMatchReport>();
        public List<HainanStage1RowMatchReport> CodeFilledFromRaw { get; } = new List<HainanStage1RowMatchReport>();
        public List<string> MissingReference { get; } = new List<string>();
        public List<string> MissingCodes { get; } = new List<string>();
        public List<string> MissingManualInfo { get; } = new List<string>();
        public Dictionary<string, int> DuplicateNamesInPowerFile { get; } = new Dictionary<string, int>();
    }
}
