using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class ChongqingStage2Layout
    {
        public const int Year = 2026;
        public const int LedgerDataStartRow = 4;
        public const int DetailDataStartRow = 5;
        public const int SummaryDataStartRow = 4;
    }

    internal sealed class ChongqingLedgerMap
    {
        public int CustomerNameColumn { get; set; }
        public int ProjectDeveloperColumn { get; set; }
        public int AgentOrSelfColumn { get; set; }
        public int OwnerColumn { get; set; }
        public int IntermediaryColumn { get; set; }
        public int PayeeColumn { get; set; }
        public int TotalPowerColumn { get; set; }
        public int SharpPowerColumn { get; set; }
        public int PeakPowerColumn { get; set; }
        public int FlatPowerColumn { get; set; }
        public int ValleyPowerColumn { get; set; }
        public int IntermediaryRatioColumn { get; set; }
        public int IntermediaryUnitPriceColumn { get; set; }
        public int IntermediaryTaxRateColumn { get; set; }
        public int IntermediaryNetColumn { get; set; }
        public int RefundRatioColumn { get; set; }
        public int RefundSharpPriceColumn { get; set; }
        public int RefundPeakPriceColumn { get; set; }
        public int RefundFlatPriceColumn { get; set; }
        public int RefundValleyPriceColumn { get; set; }
        public int RefundTaxRateColumn { get; set; }
        public int RefundNetColumn { get; set; }
        public int ProxyRatioColumn { get; set; }
        public int ProxyUnitPriceColumn { get; set; }
        public int ProxyTaxRateColumn { get; set; }
        public int ProxyNetColumn { get; set; }
        public int RecoverShortfallColumn { get; set; }
    }

    internal sealed class ChongqingSettlementDetail
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
        public double Ratio { get; set; }
        public double UnitPrice { get; set; }
        public double RefundSharpPrice { get; set; }
        public double RefundPeakPrice { get; set; }
        public double RefundFlatPrice { get; set; }
        public double RefundValleyPrice { get; set; }
        public double TaxRate { get; set; }
        public double RecoverShortfall { get; set; }
        public double Gross { get; set; }
        public double AdjustedGross { get; set; }
        public double TaxAmount { get; set; }
        public double CalculatedNet { get; set; }
        public double LedgerNet { get; set; }
        public double ExpectedNet { get; set; }
    }

    internal sealed class ChongqingSummaryMetaRow
    {
        public int Row { get; set; }
        public string Entity { get; set; }
        public string Kind { get; set; }
        public string PaymentParty { get; set; }
    }

    internal static class ChongqingStage2Keys
    {
        public static string TemplateKey(string kind, string owner, string entity)
        {
            return kind + "|" + TextUtil.CustomerKey(owner) + "|" + TextUtil.CustomerKey(entity);
        }

        public static string SummaryKey(string entity, string kind)
        {
            return TextUtil.CustomerKey(entity) + "|" + TextUtil.S(kind);
        }
    }
}
