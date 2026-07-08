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
        private readonly TextBlock _refundTemplateDirLabel;
        private readonly FrameworkElement _refundTemplateDirRow;
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
            FrameworkElement powerRow,
            TextBlock rawDetailLabel,
            TextBlock referenceLedgerLabel,
            FrameworkElement referenceLedgerRow,
            CheckBox copyReferenceExistingCheckBox,
            Button runStageOneButton,
            TextBlock runStageOneButtonText,
            Button cleanPowerButton,
            TextBlock stageTwoTitleText,
            TextBlock stageTwoCaptionText,
            TextBlock refundTemplateDirLabel,
            FrameworkElement refundTemplateDirRow,
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
            _powerRow = powerRow;
            _rawDetailLabel = rawDetailLabel;
            _referenceLedgerLabel = referenceLedgerLabel;
            _referenceLedgerRow = referenceLedgerRow;
            _copyReferenceExistingCheckBox = copyReferenceExistingCheckBox;
            _runStageOneButton = runStageOneButton;
            _runStageOneButtonText = runStageOneButtonText;
            _cleanPowerButton = cleanPowerButton;
            _stageTwoTitleText = stageTwoTitleText;
            _stageTwoCaptionText = stageTwoCaptionText;
            _refundTemplateDirLabel = refundTemplateDirLabel;
            _refundTemplateDirRow = refundTemplateDirRow;
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
            var isChongqing = province == ProvinceCode.Chongqing;
            if (hasProvince && !profile.SupportsEmployeeReward && _mainTabControl.SelectedItem == _employeeRewardTab)
            {
                _mainTabControl.SelectedItem = _mainSettlementTab;
            }

            _mainSettlementTab.Header = hasProvince ? profile.MainSettlementTabHeader : "结算流程";
            _employeeRewardTab.Visibility = hasProvince && profile.SupportsEmployeeReward ? Visibility.Visible : Visibility.Collapsed;
            _provinceEmptyPanel.Visibility = hasProvince ? Visibility.Collapsed : Visibility.Visible;
            _stageOnePanel.Visibility = hasProvince ? Visibility.Visible : Visibility.Collapsed;
            _stageTwoPanel.Visibility = hasProvince && profile.SupportsStage2 ? Visibility.Visible : Visibility.Collapsed;
            Grid.SetColumnSpan(_stageOnePanel, hasProvince && profile.SupportsStage2 ? 1 : 2);
            _stageOnePanel.Margin = hasProvince && profile.SupportsStage2 ? new Thickness(0, 0, 8, 0) : new Thickness(0);

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
            _stageTwoTitleText.Text = isChongqing ? "阶段二：重庆结算生成" : "阶段二：生成分表和汇总表";
            _stageTwoCaptionText.Text = isChongqing
                ? "生成代理/居间/退补分表和汇总表，生成前先确认预检项目"
                : "生成代理/居间分表和汇总表，输出结算结果";
            _runStageTwoButtonText.Text = isChongqing ? "开始 重庆阶段二" : "开始 执行阶段二";

            var ledgerVisibility = hasProvince ? Visibility.Visible : Visibility.Collapsed;
            var existingPowerVisibility = hasProvince && profile.ShowsExistingPowerInput ? Visibility.Visible : Visibility.Collapsed;
            var referenceLedgerVisibility = hasProvince && profile.ShowsReferenceLedgerInput ? Visibility.Visible : Visibility.Collapsed;
            var chongqingStage2Visibility = isChongqing ? Visibility.Visible : Visibility.Collapsed;
            _baseLedgerLabel.Visibility = ledgerVisibility;
            _baseLedgerRow.Visibility = ledgerVisibility;
            _powerLabel.Visibility = existingPowerVisibility;
            _powerRow.Visibility = existingPowerVisibility;
            _referenceLedgerLabel.Visibility = referenceLedgerVisibility;
            _referenceLedgerRow.Visibility = referenceLedgerVisibility;
            _copyReferenceExistingCheckBox.Visibility = referenceLedgerVisibility;
            _refundTemplateDirLabel.Visibility = chongqingStage2Visibility;
            _refundTemplateDirRow.Visibility = chongqingStage2Visibility;
            _allowMissingOwnerCheckBox.Visibility = isChongqing ? Visibility.Collapsed : Visibility.Visible;

            _runStageOneButton.IsEnabled = !isBusy && hasProvince && profile.SupportsStage1LedgerUpdate;
            _cleanPowerButton.IsEnabled = !isBusy && hasProvince && profile.SupportsStage1CleanPower;
            _runStageTwoButton.IsEnabled = !isBusy && hasProvince && profile.SupportsStage2;
            _runEmployeeRewardButton.IsEnabled = !isBusy && hasProvince && profile.SupportsEmployeeReward;
            _sharedSettingsCaption.Text = !hasProvince
                ? "请先选择结算省份；选择后会显示对应省份的可用功能"
                : profile.SharedSettingsCaption;

            return province;
        }
    }
}
