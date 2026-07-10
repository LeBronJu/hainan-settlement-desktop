namespace HainanSettlementTool.Wpf
{
    internal sealed class ProvinceStage2UiProfile
    {
        public ProvinceStage2UiProfile(
            string title,
            string caption,
            string runButtonText,
            string completedLedgerLabel,
            string proxyDirectoryLabel,
            string intermediaryDirectoryLabel,
            string refundDirectoryLabel,
            string summaryTemplateLabel,
            string thirdResultLabel,
            bool showsCompletedLedger,
            bool showsProxyDirectory,
            bool showsIntermediaryDirectory,
            bool showsRefundDirectory,
            bool showsSummaryTemplate,
            bool showsAllowMissingOwner)
        {
            Title = title;
            Caption = caption;
            RunButtonText = runButtonText;
            CompletedLedgerLabel = completedLedgerLabel;
            ProxyDirectoryLabel = proxyDirectoryLabel;
            IntermediaryDirectoryLabel = intermediaryDirectoryLabel;
            RefundDirectoryLabel = refundDirectoryLabel;
            SummaryTemplateLabel = summaryTemplateLabel;
            ThirdResultLabel = thirdResultLabel;
            ShowsCompletedLedger = showsCompletedLedger;
            ShowsProxyDirectory = showsProxyDirectory;
            ShowsIntermediaryDirectory = showsIntermediaryDirectory;
            ShowsRefundDirectory = showsRefundDirectory;
            ShowsSummaryTemplate = showsSummaryTemplate;
            ShowsAllowMissingOwner = showsAllowMissingOwner;
        }

        public string Title { get; }
        public string Caption { get; }
        public string RunButtonText { get; }
        public string CompletedLedgerLabel { get; }
        public string ProxyDirectoryLabel { get; }
        public string IntermediaryDirectoryLabel { get; }
        public string RefundDirectoryLabel { get; }
        public string SummaryTemplateLabel { get; }
        public string ThirdResultLabel { get; }
        public bool ShowsCompletedLedger { get; }
        public bool ShowsProxyDirectory { get; }
        public bool ShowsIntermediaryDirectory { get; }
        public bool ShowsRefundDirectory { get; }
        public bool ShowsSummaryTemplate { get; }
        public bool ShowsAllowMissingOwner { get; }
    }
}
