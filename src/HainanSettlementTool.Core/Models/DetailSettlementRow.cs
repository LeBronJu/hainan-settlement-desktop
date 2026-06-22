namespace HainanSettlementTool.Core.Models
{
    public sealed class DetailSettlementRow
    {
        public int LedgerRow { get; set; }
        public string Customer { get; set; }
        public string Owner { get; set; }
        public string Entity { get; set; }
        public string Kind { get; set; }
        public double Total { get; set; }
        public double Sharp { get; set; }
        public double Peak { get; set; }
        public double Flat { get; set; }
        public double Valley { get; set; }
        public double PeakFlat { get; set; }
        public double ValleyFlat { get; set; }
        public double Ratio { get; set; }
        public double UnitPrice { get; set; }
        public double TaxRate { get; set; }
        public double ExpectedNet { get; set; }
    }
}
