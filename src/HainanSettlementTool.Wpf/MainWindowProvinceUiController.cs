using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowProvinceUiController
    {
        private readonly TabControl _mainTabControl;
        private readonly TabItem _mainSettlementTab;
        private readonly TabItem _employeeRewardTab;
        private readonly ComboBox _monthCombo;
        private readonly TextBlock _settlementMonthLabel;
        private readonly TextBlock _sharedSettingsCaption;
        private readonly FrameworkElement _provinceEmptyPanel;
        private readonly FrameworkElement _stageOnePanel;
        private readonly FrameworkElement _stageTwoPanel;
        private readonly TextBlock _stageOneTitleText;
        private readonly TextBlock _stageOneCaptionText;
        private readonly TextBlock _baseLedgerLabel;
        private readonly FrameworkElement _baseLedgerRow;
        private readonly TextBlock _powerLabel;
        private readonly FrameworkElement _powerRow;
        private readonly TextBlock _rawDetailLabel;
        private readonly TextBlock _referenceLedgerLabel;
        private readonly FrameworkElement _referenceLedgerRow;
        private readonly CheckBox _copyReferenceExistingCheckBox;
        private readonly Button _runStageOneButton;
        private readonly TextBlock _runStageOneButtonText;
        private readonly Button _cleanPowerButton;
        private readonly TextBlock _stageTwoTitleText;
        private readonly TextBlock _stageTwoCaptionText;
        private readonly TextBlock _completedLedgerLabel;
        private readonly FrameworkElement _completedLedgerRow;
        private readonly TextBlock _proxyTemplateDirLabel;
        private readonly FrameworkElement _proxyTemplateDirRow;
        private readonly TextBlock _intermediaryTemplateDirLabel;
        private readonly FrameworkElement _intermediaryTemplateDirRow;
        private readonly TextBlock _refundTemplateDirLabel;
        private readonly FrameworkElement _refundTemplateDirRow;
        private readonly TextBlock _summaryTemplateLabel;
        private readonly FrameworkElement _summaryTemplateRow;
        private readonly CheckBox _allowMissingOwnerCheckBox;
        private readonly Button _runStageTwoButton;
        private readonly TextBlock _runStageTwoButtonText;
        private readonly Button _runEmployeeRewardButton;
        private readonly Func<string, Brush> _brushOf;

        public MainWindowProvinceUiController(
            TabControl mainTabControl,
            TabItem mainSettlementTab,
            TabItem employeeRewardTab,
            ComboBox monthCombo,
            TextBlock settlementMonthLabel,
            TextBlock sharedSettingsCaption,
            FrameworkElement provinceEmptyPanel,
            FrameworkElement stageOnePanel,
            FrameworkElement stageTwoPanel,
            TextBlock stageOneTitleText,
            TextBlock stageOneCaptionText,
            TextBlock baseLedgerLabel,
            FrameworkElement baseLedgerRow,
            TextBlock powerLabel,
            FrameworkElement PowerInputRow,
            TextBlock rawDetailLabel,
            TextBlock referenceLedgerLabel,
            FrameworkElement referenceLedgerRow,
            CheckBox copyReferenceExistingCheckBox,
            Button runStageOneButton,
            TextBlock runStageOneButtonText,
            Button cleanPowerButton,
            TextBlock stageTwoTitleText,
            TextBlock stageTwoCaptionText,
            TextBlock completedLedgerLabel,
            FrameworkElement completedLedgerRow,
            TextBlock proxyTemplateDirLabel,
            FrameworkElement proxyTemplateDirRow,
            TextBlock intermediaryTemplateDirLabel,
            FrameworkElement intermediaryTemplateDirRow,
            TextBlock refundTemplateDirLabel,
            FrameworkElement refundTemplateDirRow,
            TextBlock summaryTemplateLabel,
            FrameworkElement summaryTemplateRow,
            CheckBox allowMissingOwnerCheckBox,
            Button runStageTwoButton,
            TextBlock runStageTwoButtonText,
            Button runEmployeeRewardButton,
            Func<string, Brush> brushOf)
        {
            _mainTabControl = mainTabControl;
            _mainSettlementTab = mainSettlementTab;
            _employeeRewardTab = employeeRewardTab;
            _monthCombo = monthCombo;
            _settlementMonthLabel = settlementMonthLabel;
            _sharedSettingsCaption = sharedSettingsCaption;
            _provinceEmptyPanel = provinceEmptyPanel;
            _stageOnePanel = stageOnePanel;
            _stageTwoPanel = stageTwoPanel;
            _stageOneTitleText = stageOneTitleText;
            _stageOneCaptionText = stageOneCaptionText;
            _baseLedgerLabel = baseLedgerLabel;
            _baseLedgerRow = baseLedgerRow;
            _powerLabel = powerLabel;
            _powerRow = PowerInputRow;
            _rawDetailLabel = rawDetailLabel;
            _referenceLedgerLabel = referenceLedgerLabel;
            _referenceLedgerRow = referenceLedgerRow;
            _copyReferenceExistingCheckBox = copyReferenceExistingCheckBox;
            _runStageOneButton = runStageOneButton;
            _runStageOneButtonText = runStageOneButtonText;
            _cleanPowerButton = cleanPowerButton;
            _stageTwoTitleText = stageTwoTitleText;
            _stageTwoCaptionText = stageTwoCaptionText;
            _completedLedgerLabel = completedLedgerLabel;
            _completedLedgerRow = completedLedgerRow;
            _proxyTemplateDirLabel = proxyTemplateDirLabel;
            _proxyTemplateDirRow = proxyTemplateDirRow;
            _intermediaryTemplateDirLabel = intermediaryTemplateDirLabel;
            _intermediaryTemplateDirRow = intermediaryTemplateDirRow;
            _refundTemplateDirLabel = refundTemplateDirLabel;
            _refundTemplateDirRow = refundTemplateDirRow;
            _summaryTemplateLabel = summaryTemplateLabel;
            _summaryTemplateRow = summaryTemplateRow;
            _allowMissingOwnerCheckBox = allowMissingOwnerCheckBox;
            _runStageTwoButton = runStageTwoButton;
            _runStageTwoButtonText = runStageTwoButtonText;
            _runEmployeeRewardButton = runEmployeeRewardButton;
            _brushOf = brushOf;
        }

        public ProvinceCode? ApplySharedSettings(bool isBusy, ProvinceUiProfile profile, bool employeeRewardSelected)
        {
            var monthEnabled = !isBusy && !employeeRewardSelected;
            _monthCombo.IsEnabled = monthEnabled;
            _settlementMonthLabel.Foreground = monthEnabled ? _brushOf("FieldTextBrush") : _brushOf("MutedBrush");

            var province = ApplyProvinceUi(isBusy, profile);
            if (employeeRewardSelected)
            {
                _sharedSettingsCaption.Text = "员工电量奖励使用本页的开始/结束月份，输出仍保存到这个文件夹中";
            }

            return province;
        }

        public ProvinceCode? ApplyProvinceUi(bool isBusy, ProvinceUiProfile profile)
        {
            var hasProvince = profile != null;
            var province = hasProvince ? (ProvinceCode?)profile.Province : null;
            var hasStageOne = hasProvince && (profile.SupportsStage1LedgerUpdate || profile.SupportsStage1CleanPower);
            var hasStageTwo = hasProvince && profile.SupportsStage2;
            var showEmployeeRewardTab = hasProvince
                && (profile.SupportsEmployeeReward || profile.ShowEmployeeRewardPlaceholder);
            var enableEmployeeRewardTab = hasProvince && profile.SupportsEmployeeReward;
            if (hasProvince && !enableEmployeeRewardTab && _mainTabControl.SelectedItem == _employeeRewardTab)
            {
                _mainTabControl.SelectedItem = _mainSettlementTab;
            }

            _mainSettlementTab.Header = hasProvince ? profile.MainSettlementTabHeader : "结算流程";
            _employeeRewardTab.Visibility = showEmployeeRewardTab ? Visibility.Visible : Visibility.Collapsed;
            _employeeRewardTab.IsEnabled = enableEmployeeRewardTab;
            _employeeRewardTab.ToolTip = showEmployeeRewardTab && !profile.SupportsEmployeeReward ? "正在开发中" : null;
            _provinceEmptyPanel.Visibility = hasProvince ? Visibility.Collapsed : Visibility.Visible;
            _stageOnePanel.Visibility = hasStageOne ? Visibility.Visible : Visibility.Collapsed;
            _stageTwoPanel.Visibility = hasStageTwo ? Visibility.Visible : Visibility.Collapsed;
            Grid.SetColumnSpan(_stageOnePanel, hasStageTwo ? 1 : 2);
            _stageOnePanel.Margin = hasStageTwo ? new Thickness(0, 0, 8, 0) : new Thickness(0);
            Grid.SetColumn(_stageTwoPanel, hasStageOne ? 1 : 0);
            Grid.SetColumnSpan(_stageTwoPanel, hasStageOne ? 1 : 2);
            _stageTwoPanel.Margin = hasStageOne ? new Thickness(8, 0, 0, 0) : new Thickness(0);

            _stageOneTitleText.Text = hasProvince ? profile.StageOneTitle : "请先选择结算省份";
            _stageOneCaptionText.Text = !hasProvince
                ? "不同省份的结算口径不同，必须先选择省份再执行。"
                : profile.StageOneCaption;
            _baseLedgerLabel.Text = hasProvince ? profile.BaseLedgerLabel : "基础台账（先选择省份）";
            _powerLabel.Text = hasProvince ? profile.ExistingPowerLabel : "电量处理表";
            _rawDetailLabel.Text = hasProvince ? profile.RawDetailLabel : "电量文件（先选择省份）";
            _referenceLedgerLabel.Text = hasProvince ? profile.ReferenceLedgerLabel : "参考台账（可选）";
            _runStageOneButtonText.Text = hasProvince ? profile.RunStageOneButtonText : "开始 执行阶段一";
            _cleanPowerButton.Content = hasProvince ? profile.CleanPowerButtonText : "只清洗电量";
            var stageTwo = hasStageTwo ? profile.StageTwo : null;
            _stageTwoTitleText.Text = stageTwo?.Title ?? "阶段二";
            _stageTwoCaptionText.Text = stageTwo?.Caption ?? "当前省份暂未开放阶段二。";
            _runStageTwoButtonText.Text = stageTwo?.RunButtonText ?? "暂未开放";
            _completedLedgerLabel.Text = stageTwo?.CompletedLedgerLabel ?? "人工整理后的台账";
            _proxyTemplateDirLabel.Text = stageTwo?.ProxyDirectoryLabel ?? "代理分表文件夹";
            _intermediaryTemplateDirLabel.Text = stageTwo?.IntermediaryDirectoryLabel ?? "居间分表文件夹";
            _refundTemplateDirLabel.Text = stageTwo?.RefundDirectoryLabel ?? "退补分表文件夹";
            _summaryTemplateLabel.Text = stageTwo?.SummaryTemplateLabel ?? "汇总表";

            var ledgerVisibility = hasProvince ? Visibility.Visible : Visibility.Collapsed;
            var existingPowerVisibility = hasProvince && profile.ShowsExistingPowerInput ? Visibility.Visible : Visibility.Collapsed;
            var referenceLedgerVisibility = hasProvince && profile.ShowsReferenceLedgerInput ? Visibility.Visible : Visibility.Collapsed;
            _baseLedgerLabel.Visibility = ledgerVisibility;
            _baseLedgerRow.Visibility = ledgerVisibility;
            _powerLabel.Visibility = existingPowerVisibility;
            _powerRow.Visibility = existingPowerVisibility;
            _referenceLedgerLabel.Visibility = referenceLedgerVisibility;
            _referenceLedgerRow.Visibility = referenceLedgerVisibility;
            _copyReferenceExistingCheckBox.Visibility = referenceLedgerVisibility;
            _completedLedgerLabel.Visibility = VisibilityOf(stageTwo?.ShowsCompletedLedger == true);
            _completedLedgerRow.Visibility = VisibilityOf(stageTwo?.ShowsCompletedLedger == true);
            _proxyTemplateDirLabel.Visibility = VisibilityOf(stageTwo?.ShowsProxyDirectory == true);
            _proxyTemplateDirRow.Visibility = VisibilityOf(stageTwo?.ShowsProxyDirectory == true);
            _intermediaryTemplateDirLabel.Visibility = VisibilityOf(stageTwo?.ShowsIntermediaryDirectory == true);
            _intermediaryTemplateDirRow.Visibility = VisibilityOf(stageTwo?.ShowsIntermediaryDirectory == true);
            _refundTemplateDirLabel.Visibility = VisibilityOf(stageTwo?.ShowsRefundDirectory == true);
            _refundTemplateDirRow.Visibility = VisibilityOf(stageTwo?.ShowsRefundDirectory == true);
            _summaryTemplateLabel.Visibility = VisibilityOf(stageTwo?.ShowsSummaryTemplate == true);
            _summaryTemplateRow.Visibility = VisibilityOf(stageTwo?.ShowsSummaryTemplate == true);
            _allowMissingOwnerCheckBox.Visibility = VisibilityOf(stageTwo?.ShowsAllowMissingOwner == true);

            _runStageOneButton.IsEnabled = !isBusy && hasProvince && profile.SupportsStage1LedgerUpdate;
            _cleanPowerButton.IsEnabled = !isBusy && hasProvince && profile.SupportsStage1CleanPower;
            _runStageTwoButton.IsEnabled = !isBusy && hasProvince && profile.SupportsStage2;
            _runEmployeeRewardButton.IsEnabled = !isBusy && enableEmployeeRewardTab;
            _sharedSettingsCaption.Text = !hasProvince
                ? "请先选择结算省份；选择后会显示对应省份的可用功能"
                : profile.SharedSettingsCaption;

            return province;
        }

        private static Visibility VisibilityOf(bool visible)
        {
            return visible ? Visibility.Visible : Visibility.Collapsed;
        }
    }
}
