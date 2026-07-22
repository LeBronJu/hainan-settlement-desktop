using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using HainanSettlementTool.Core.Models;
using Microsoft.Win32;

namespace HainanSettlementTool.Wpf
{
    public partial class MainWindow : Window
    {
        private readonly MainWindowProgressController _progressController;
        private readonly MainWindowResultController _resultController;
        private readonly MainWindowDialogController _dialogController;
        private readonly MainWindowPathPickerController _pathPickerController;
        private readonly MainWindowLogController _logController;
        private readonly MainWindowInputController _inputController;
        private readonly MainWindowProvinceUiController _provinceUiController;
        private readonly MainWindowStage2WorkflowController _stage2WorkflowController;
        private readonly MainWindowHainanStage1WorkflowController _hainanStage1WorkflowController;
        private readonly MainWindowChongqingStage1WorkflowController _chongqingStage1WorkflowController;
        private readonly MainWindowHainanEmployeePowerRewardWorkflowController _hainanEmployeePowerRewardWorkflowController;
        private bool _isBusy;
        private bool _loadingInputs;
        private ProvinceCode? _activeInputProvince;
        private string _themeMode = ThemeService.SystemMode;

        public MainWindow()
        {
            var snapshot = UserInputStore.Load();
            _themeMode = ThemeService.NormalizeMode(snapshot.ThemeMode);
            ThemeService.Apply(_themeMode);

            InitializeComponent();

            _dialogController = new MainWindowDialogController(this);
            _pathPickerController = new MainWindowPathPickerController(this);
            _inputController = new MainWindowInputController(
                MonthCombo,
                RewardStartMonthCombo,
                RewardEndMonthCombo,
                ProvinceCombo,
                OutputDirBox,
                BaseLedgerBox,
                PowerBox,
                RawDetailBox,
                ReferenceLedgerBox,
                CompletedLedgerBox,
                ProxyTemplateDirBox,
                IntermediaryTemplateDirBox,
                RefundTemplateDirBox,
                SummaryTemplateBox,
                RewardLedgerBox,
                CopyReferenceExistingCheckBox,
                AllowMissingOwnerCheckBox);
            InitializeThemeCombo(_themeMode);
            InitializeProvinceCombo();

            for (var month = 2; month <= 12; month++)
            {
                MonthCombo.Items.Add("2026年" + month + "月");
            }

            for (var month = 1; month <= 12; month++)
            {
                RewardStartMonthCombo.Items.Add("2026年" + month + "月");
                RewardEndMonthCombo.Items.Add("2026年" + month + "月");
            }

            MonthCombo.SelectedIndex = -1;
            RewardStartMonthCombo.SelectedIndex = 0;
            RewardEndMonthCombo.SelectedIndex = -1;
            _progressController = new MainWindowProgressController(
                StatusText,
                StatusDot,
                StatusPill,
                ProgressTitle,
                ProgressDescriptionText,
                ProgressBar,
                ProgressPercentText,
                new[] { Step1Text, Step2Text, Step3Text, Step4Text, Step5Text },
                new[] { Step1Status, Step2Status, Step3Status, Step4Status, Step5Status },
                BrushOf);
            _resultController = new MainWindowResultController(
                CompletionCard,
                CompletionIconCircle,
                CompletionIconText,
                CompletionTitleText,
                CompletionDetailText,
                CompletionOutputLabel,
                CompletionOutputText,
                OpenReadableReportButton,
                NoProvinceResultHint,
                Stage1ResultRow,
                Stage1ResultLabel,
                Stage1ResultStatus,
                Stage1ResultCount,
                ProxyResultRow,
                ProxyResultStatus,
                ProxyResultCount,
                IntermediaryResultRow,
                IntermediaryResultStatus,
                IntermediaryResultCount,
                SummaryResultRow,
                SummaryResultLabel,
                SummaryResultStatus,
                SummaryResultCount,
                EmployeeRewardResultRow,
                EmployeeRewardResultStatus,
                EmployeeRewardResultCount,
                FinishedAtRow,
                FinishedAtText,
                BrushOf);
            _provinceUiController = new MainWindowProvinceUiController(
                MainTabControl,
                MainSettlementTab,
                EmployeeRewardTab,
                MonthCombo,
                SettlementMonthLabel,
                SharedSettingsCaption,
                ProvinceEmptyPanel,
                StageOnePanel,
                StageTwoPanel,
                Stage1TitleText,
                Stage1CaptionText,
                BaseLedgerLabel,
                BaseLedgerRow,
                PowerLabel,
                PowerInputRow,
                RawDetailLabel,
                ReferenceLedgerLabel,
                ReferenceLedgerRow,
                CopyReferenceExistingCheckBox,
                RunStage1Button,
                RunStage1ButtonText,
                CleanPowerButton,
                Stage2TitleText,
                Stage2CaptionText,
                CompletedLedgerLabel,
                CompletedLedgerRow,
                ProxyTemplateDirLabel,
                ProxyTemplateDirRow,
                IntermediaryTemplateDirLabel,
                IntermediaryTemplateDirRow,
                RefundTemplateDirLabel,
                RefundTemplateDirRow,
                SummaryTemplateLabel,
                SummaryTemplateRow,
                AllowMissingOwnerCheckBox,
                RunStage2Button,
                RunStage2ButtonText,
                RunEmployeeRewardButton,
                BrushOf);
            _logController = new MainWindowLogController(this, LogBox);
            _stage2WorkflowController = new MainWindowStage2WorkflowController(
                this,
                Dispatcher,
                _inputController,
                _progressController,
                _resultController,
                _logController,
                _dialogController,
                SetBusy,
                SaveInputs);
            _hainanStage1WorkflowController = new MainWindowHainanStage1WorkflowController(
                Dispatcher,
                _inputController,
                _progressController,
                _resultController,
                _logController,
                _dialogController,
                SetBusy,
                SaveInputs);
            _chongqingStage1WorkflowController = new MainWindowChongqingStage1WorkflowController(
                this,
                Dispatcher,
                _inputController,
                _progressController,
                _resultController,
                _logController,
                _dialogController,
                SetBusy,
                SaveInputs);
            _hainanEmployeePowerRewardWorkflowController = new MainWindowHainanEmployeePowerRewardWorkflowController(
                Dispatcher,
                _inputController,
                _progressController,
                _resultController,
                _logController,
                _dialogController,
                SetBusy,
                SaveInputs);
            _progressController.ResetProgress("等待执行", "尚未开始");
            ResetResults();
            AddLog("工具已就绪，等待操作。", "信息");
            LoadSavedInputs(snapshot);
            _activeInputProvince = ParseProvince(snapshot.ProvinceCode);
            UpdateSharedSettingsState();
            ResetResults();
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                ToggleMaximize();
                return;
            }

            DragMove();
        }

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void Maximize_Click(object sender, RoutedEventArgs e)
        {
            ToggleMaximize();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            SaveInputs();
            base.OnClosing(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            SystemEvents.UserPreferenceChanged -= SystemEvents_UserPreferenceChanged;
            base.OnClosed(e);
        }

        private void ToggleMaximize()
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }

        private void BrowseOutputDir_Click(object sender, RoutedEventArgs e)
        {
            BrowseFolder(OutputDirBox, "选择结果输出文件夹");
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(e.OriginalSource, MainTabControl))
            {
                return;
            }

            UpdateSharedSettingsState();
        }

        private void ThemeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loadingInputs)
            {
                return;
            }

            _themeMode = SelectedThemeMode();
            ThemeService.Apply(_themeMode);
            RefreshThemeDependentState();
            SaveInputs();
        }

        private void ProvinceCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loadingInputs)
            {
                return;
            }

            var selectedProvince = _inputController.SelectedProvinceOrNull();
            if (_activeInputProvince.HasValue
                && selectedProvince.HasValue
                && _activeInputProvince.Value != selectedProvince.Value)
            {
                _inputController.ClearStage1();
                _inputController.ClearStage2();
                _inputController.ClearEmployeeReward();
                AddLog("已切换结算省份，原省份专用输入路径已清空。", "提示");
            }

            _activeInputProvince = selectedProvince;
            UpdateProvinceUi();
            ResetResults();
            SaveInputs();
        }

        private void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (_themeMode != ThemeService.SystemMode)
            {
                return;
            }

            Dispatcher.Invoke(() =>
            {
                ThemeService.Apply(_themeMode);
                RefreshThemeDependentState();
            });
        }

        private void BrowseBaseLedger_Click(object sender, RoutedEventArgs e)
        {
            var profile = _inputController.SelectedProfileOrNull();
            var title = profile?.BaseLedgerDialogTitle ?? "选择基础台账";
            BrowseExcel(BaseLedgerBox, title);
        }

        private void BrowsePower_Click(object sender, RoutedEventArgs e)
        {
            BrowseExcel(PowerBox, "选择电量处理表");
        }

        private void BrowseRawDetail_Click(object sender, RoutedEventArgs e)
        {
            var profile = _inputController.SelectedProfileOrNull();
            if (profile == null)
            {
                ShowErrorMessage("请选择结算省份后再选择电量文件。");
                return;
            }

            BrowseFile(RawDetailBox, profile.RawDetailDialogTitle, "Excel/CSV|*.xlsx;*.xls;*.csv|Excel 文件|*.xlsx;*.xls|CSV 文件|*.csv|所有文件|*.*");
        }

        private void BrowseReferenceLedger_Click(object sender, RoutedEventArgs e)
        {
            BrowseExcel(ReferenceLedgerBox, "选择参考台账");
        }

        private void BrowseCompletedLedger_Click(object sender, RoutedEventArgs e)
        {
            BrowseExcel(CompletedLedgerBox, "选择人工整理后的台账");
        }

        private void BrowseProxyTemplateDir_Click(object sender, RoutedEventArgs e)
        {
            BrowseFolder(ProxyTemplateDirBox, "选择代理分表文件夹");
        }

        private void BrowseIntermediaryTemplateDir_Click(object sender, RoutedEventArgs e)
        {
            BrowseFolder(IntermediaryTemplateDirBox, "选择居间分表文件夹");
        }

        private void BrowseRefundTemplateDir_Click(object sender, RoutedEventArgs e)
        {
            BrowseFolder(RefundTemplateDirBox, "选择退补分表文件夹");
        }

        private void BrowseSummaryTemplate_Click(object sender, RoutedEventArgs e)
        {
            BrowseExcel(SummaryTemplateBox, "选择上月/修正版汇总表");
        }

        private void BrowseRewardLedger_Click(object sender, RoutedEventArgs e)
        {
            BrowseExcel(RewardLedgerBox, "选择最新售电结算台账");
        }

        private void BrowseExcel(TextBox target, string title)
        {
            if (_pathPickerController.BrowseExcel(target, title))
            {
                SaveInputs();
            }
        }

        private void BrowseFile(TextBox target, string title, string filter)
        {
            if (_pathPickerController.BrowseFile(target, title, filter))
            {
                SaveInputs();
            }
        }

        private void BrowseFolder(TextBox target, string title)
        {
            if (_pathPickerController.BrowseFolder(target, title))
            {
                SaveInputs();
            }
        }

        private async void RunStage1_Click(object sender, RoutedEventArgs e)
        {
            var profile = _inputController.SelectedProfileOrNull();
            if (profile == null)
            {
                ShowErrorMessage("请选择结算省份后再执行阶段一。");
                return;
            }

            if (!profile.SupportsStage1LedgerUpdate)
            {
                ShowErrorMessage(profile.DisplayName + "暂未开放阶段一台账更新。");
                return;
            }

            switch (profile.Province)
            {
                case ProvinceCode.Hainan:
                    await _hainanStage1WorkflowController.RunLedgerUpdateAsync();
                    return;
                case ProvinceCode.Chongqing:
                    await _chongqingStage1WorkflowController.RunLedgerUpdateAsync();
                    return;
                default:
                    ShowErrorMessage(profile.DisplayName + "暂未接入阶段一台账更新。");
                    return;
            }

        }

        private async void CleanPower_Click(object sender, RoutedEventArgs e)
        {
            var profile = _inputController.SelectedProfileOrNull();
            if (profile == null)
            {
                ShowErrorMessage("请选择结算省份后再清洗电量数据。");
                return;
            }

            if (!profile.SupportsStage1CleanPower)
            {
                ShowErrorMessage(profile.DisplayName + "暂未开放只清洗电量数据。");
                return;
            }

            switch (profile.Province)
            {
                case ProvinceCode.Hainan:
                    await _hainanStage1WorkflowController.CleanPowerAsync();
                    return;
                case ProvinceCode.Chongqing:
                    await _chongqingStage1WorkflowController.CleanPowerAsync();
                    return;
                default:
                    ShowErrorMessage(profile.DisplayName + "暂未接入只清洗电量数据。");
                    return;
            }

        }

        private async void RunStage2_Click(object sender, RoutedEventArgs e)
        {
            await _stage2WorkflowController.RunAsync();
        }

        private async void RunEmployeeReward_Click(object sender, RoutedEventArgs e)
        {
            await _hainanEmployeePowerRewardWorkflowController.RunAsync();
        }

        private void InitializeThemeCombo(string mode)
        {
            _loadingInputs = true;
            ThemeCombo.Items.Add("跟随系统");
            ThemeCombo.Items.Add("浅色");
            ThemeCombo.Items.Add("深色");
            ThemeCombo.SelectedIndex = ThemeIndex(mode);
            _loadingInputs = false;
        }

        private void InitializeProvinceCombo()
        {
            _loadingInputs = true;
            ProvinceCombo.DisplayMemberPath = nameof(ProvinceUiProfile.DisplayName);
            ProvinceCombo.ItemsSource = ProvinceUiProfile.Supported;
            ProvinceCombo.SelectedIndex = -1;
            _loadingInputs = false;
        }

        private static int ThemeIndex(string mode)
        {
            switch (ThemeService.NormalizeMode(mode))
            {
                case ThemeService.LightMode:
                    return 1;
                case ThemeService.DarkMode:
                    return 2;
                default:
                    return 0;
            }
        }

        private string SelectedThemeMode()
        {
            switch (ThemeCombo.SelectedIndex)
            {
                case 1:
                    return ThemeService.LightMode;
                case 2:
                    return ThemeService.DarkMode;
                default:
                    return ThemeService.SystemMode;
            }
        }

        private void SetBusy(bool busy)
        {
            _isBusy = busy;
            RunStage1Button.IsEnabled = !busy;
            CleanPowerButton.IsEnabled = !busy;
            RunStage2Button.IsEnabled = !busy;
            RunEmployeeRewardButton.IsEnabled = !busy;
            ClearStage1Button.IsEnabled = !busy;
            ClearStage2Button.IsEnabled = !busy;
            ClearEmployeeRewardButton.IsEnabled = !busy;
            SetStatus(
                busy ? "运行中" : "就绪",
                busy ? "WarningBrush" : "SuccessBrush",
                busy ? "StatusBusyBrush" : "StatusReadyBrush");
            UpdateSharedSettingsState();
        }

        private void UpdateSharedSettingsState()
        {
            if (_inputController == null || _provinceUiController == null)
            {
                return;
            }

            var employeeRewardSelected = MainTabControl != null && MainTabControl.SelectedIndex == 1;
            var province = _provinceUiController.ApplySharedSettings(
                _isBusy,
                _inputController.SelectedProfileOrNull(),
                employeeRewardSelected);
            UpdateResultVisibility(province);
        }

        private void UpdateProvinceUi()
        {
            if (_inputController == null || _provinceUiController == null)
            {
                return;
            }

            var province = _provinceUiController.ApplyProvinceUi(_isBusy, _inputController.SelectedProfileOrNull());
            UpdateResultVisibility(province);
        }

        private void RefreshThemeDependentState()
        {
            UpdateSharedSettingsState();
            _progressController.RefreshStatusBrushes();
        }

        private void SetStatus(string text, string dotBrushKey, string backgroundBrushKey)
        {
            _progressController.SetStatus(text, dotBrushKey, backgroundBrushKey);
        }

        private void UpdateResultVisibility(ProvinceCode? province)
        {
            _resultController.UpdateResultVisibility(province);
        }

        private void ResetResults()
        {
            _resultController.Reset(_inputController.SelectedProvinceOrNull());
        }

        private void LoadSavedInputs(UserInputSnapshot snapshot)
        {
            _loadingInputs = true;
            _inputController.LoadSavedInputs(snapshot);
            _loadingInputs = false;

            if (_inputController.HasSavedInputs(snapshot))
            {
                AddLog("已载入上次选择的文件路径；结算省份和月份仍需手动选择。", "信息");
            }
        }

        private void SaveInputs()
        {
            try
            {
                _inputController.SaveInputs(_themeMode);
            }
            catch (Exception ex)
            {
                AddLog("保存上次路径失败：" + ex.Message, "提示");
            }
        }

        private static ProvinceCode? ParseProvince(string value)
        {
            ProvinceCode province;
            return Enum.TryParse(value, out province) && Enum.IsDefined(typeof(ProvinceCode), province)
                ? (ProvinceCode?)province
                : null;
        }

        private void AddLog(string message, string level)
        {
            _logController.Add(message, level);
        }

        private void ShowError(Exception ex)
        {
            _dialogController.ShowError(ex);
        }

        private void ShowErrorMessage(string message)
        {
            _dialogController.ShowErrorMessage(message);
        }

        private void ClearStage1_Click(object sender, RoutedEventArgs e)
        {
            _inputController.ClearStage1();
            SaveInputs();
        }

        private void ClearStage2_Click(object sender, RoutedEventArgs e)
        {
            _inputController.ClearStage2();
            SaveInputs();
        }

        private void ClearEmployeeReward_Click(object sender, RoutedEventArgs e)
        {
            _inputController.ClearEmployeeReward();
            SaveInputs();
        }

        private void ClearLog_Click(object sender, RoutedEventArgs e)
        {
            _logController.Clear();
        }

        private void SaveLog_Click(object sender, RoutedEventArgs e)
        {
            _logController.Save();
        }

        private void OpenOutputDir_Click(object sender, RoutedEventArgs e)
        {
            var path = !string.IsNullOrWhiteSpace(_resultController.LastOutputDirectory)
                ? _resultController.LastOutputDirectory
                : OutputDirBox.Text.Trim();
            if (Directory.Exists(path))
            {
                Process.Start(path);
            }
        }

        private void OpenReadableReport_Click(object sender, RoutedEventArgs e)
        {
            var path = _resultController.LastReadableReportPath;
            if (File.Exists(path))
            {
                Process.Start(path);
            }
        }

        private Brush BrushOf(string key)
        {
            return (Brush)FindResource(key);
        }
    }
}
