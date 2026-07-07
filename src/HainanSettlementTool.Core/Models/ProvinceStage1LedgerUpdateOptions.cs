using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class ProvinceStage1LedgerUpdateOptions
    {
        public ProvinceCode Province { get; set; }
        public int Month { get; set; }
        public string LedgerPath { get; set; }
        public string RawDetailPath { get; set; }
        public string OutputDirectory { get; set; }
        public List<ProvinceStage1CustomerDecision> CustomerDecisions { get; set; } = new List<ProvinceStage1CustomerDecision>();
        public List<ProvinceStage1CustomerMatch> ManualCustomerMatches { get; set; } = new List<ProvinceStage1CustomerMatch>();
    }
}
