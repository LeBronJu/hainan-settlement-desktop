using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;
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
        private bool _isBusy;
        private bool _loadingInputs;
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
                SummaryResultStatus,
                SummaryResultCount,
                EmployeeRewardResultRow,
                EmployeeRewardResultStatus,
                EmployeeRewardResultCount,
                FinishedAtRow,
                FinishedAtText,
                BrushOf);
            _logController = new MainWindowLogController(this, LogBox);
            ResetProgress("等待执行", "尚未开始");
            ResetResults();
            AddLog("工具已就绪，等待操作。", "信息");
            LoadSavedInputs(snapshot);
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
            BrowseFolder(ProxyTemplateDirBox, "选择上月代理分表文件夹");
        }

        private void BrowseIntermediaryTemplateDir_Click(object sender, RoutedEventArgs e)
        {
            BrowseFolder(IntermediaryTemplateDirBox, "选择上月居间分表文件夹");
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
                    break;
                case ProvinceCode.Chongqing:
                    await RunProvinceStage1LedgerUpdateAsync();
                    return;
                default:
                    ShowErrorMessage(profile.DisplayName + "暂未接入阶段一台账更新。");
                    return;
            }

            Stage1Options options;
            try
            {
                options = _inputController.CreateStage1Options();
                SaveInputs();
                if (!ConfirmRun("阶段一：写入电量到台账", options.Month, options.OutputDirectory))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
                return;
            }

            SetBusy(true);
            ResetResults();
            ResetProgress("正在执行阶段一...", "写入电量到台账");
            SetProgress(10, "检查输入文件");
            SetStepRunning(0);
            AddLog("开始执行阶段一。", "阶段一");

            try
            {
                StageWorkflowResult<Stage1Report> result = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        SetStepDone(0);
                        SetStepRunning(1);
                        SetProgress(28, "读取台账和电量文件");
                    });

                    result = CreateWorkflow().RunStage1(options, LogThreadSafe);
                });

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "阶段一执行完成");
                LogSummary(result.SummaryLines);

                SetStage1ResultSuccess("1 个文件");
                ShowCompletion("阶段一执行完成", "台账和检查报告已生成", options.OutputDirectory);
            }
            catch (Exception ex)
            {
                SetStepFailed();
                SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
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
                    break;
                case ProvinceCode.Chongqing:
                    await RunProvinceStage1CleanPowerAsync();
                    return;
                default:
                    ShowErrorMessage(profile.DisplayName + "暂未接入只清洗电量数据。");
                    return;
            }

            string rawDetailPath;
            string outputPath;
            string outputDirectory;
            try
            {
                rawDetailPath = RawDetailBox.Text.Trim();
                outputPath = _inputController.ResolvePowerOutputPath(rawDetailPath);
                outputDirectory = OutputDirBox.Text.Trim();
                PowerBox.Text = outputPath;
                SaveInputs();

                var message = "即将清洗原始零售侧明细并生成电量处理表。\n\n输出文件：\n" + outputPath;
                if (!ConfirmAction("确认清洗电量", "即将清洗电量数据", message, "开始清洗"))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
                return;
            }

            SetBusy(true);
            ResetResults();
            ResetProgress("正在清洗电量...", "生成电量处理表");
            SetProgress(10, "检查输入文件");
            SetStepRunning(0);
            AddLog("开始清洗电量数据。", "阶段一");

            try
            {
                StageWorkflowResult<PowerCleanReport> result = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        SetStepDone(0);
                        SetStepRunning(1);
                        SetProgress(35, "读取原始零售侧明细");
                    });

                    result = CreateWorkflow().CleanPowerData(rawDetailPath, outputPath, LogThreadSafe);
                });

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "电量清洗完成");
                LogSummary(result.SummaryLines);

                SetStage1ResultSuccess(result.Report.PowerRows + " 个客户");
                ShowCompletion("电量清洗完成", "电量处理表已生成", outputDirectory);
            }
            catch (Exception ex)
            {
                SetStepFailed();
                SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async Task RunProvinceStage1CleanPowerAsync()
        {
            ProvinceStage1CleanOptions options;
            try
            {
                options = _inputController.CreateProvinceStage1CleanOptions();
                SaveInputs();

                var message = "即将清洗重庆交易中心电量确认结算单。\n\n输出内容：用户电量汇总、户号明细、JSON校验报告\n输出文件夹：\n" + options.OutputDirectory;
                if (!ConfirmAction("确认清洗重庆电量", "即将清洗重庆电量数据", message, "开始清洗"))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
                return;
            }

            SetBusy(true);
            ResetResults();
            ResetProgress("正在清洗重庆电量...", "生成阶段一电量处理表");
            SetProgress(10, "检查输入文件");
            SetStepRunning(0);
            AddLog("开始清洗重庆阶段一电量数据。", "重庆");

            try
            {
                StageWorkflowResult<ProvinceStage1CleanResult> result = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        SetStepDone(0);
                        SetStepRunning(1);
                        SetProgress(35, "读取交易中心电量确认结算单");
                    });

                    result = CreateWorkflow().CleanProvinceStage1PowerData(options, LogThreadSafe);
                });

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "重庆电量清洗完成");
                LogSummary(result.SummaryLines);

                SetStage1ResultSuccess(result.Report.CustomerRows + " 个客户");
                ShowCompletion("重庆电量清洗完成", "电量处理表、户号明细和校验报告已生成", options.OutputDirectory);
            }
            catch (Exception ex)
            {
                SetStepFailed();
                SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async Task RunProvinceStage1LedgerUpdateAsync()
        {
            ProvinceStage1LedgerUpdateOptions options;
            try
            {
                options = _inputController.CreateProvinceStage1LedgerUpdateOptions();
                SaveInputs();
            }
            catch (Exception ex)
            {
                ShowError(ex);
                return;
            }

            SetBusy(true);
            ResetResults();
            ResetProgress("正在预检重庆台账...", "读取台账和电量明细");
            SetProgress(10, "检查输入文件");
            SetStepRunning(0);
            AddLog("开始预检重庆阶段一台账更新。", "重庆");

            try
            {
                var workflow = CreateWorkflow();
                ProvinceStage1LedgerUpdatePlan plan = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() => SetProgress(22, "读取重庆台账和交易中心电量确认单"));
                    plan = workflow.PlanProvinceStage1LedgerUpdate(options, LogThreadSafe);
                });

                SetProgress(30, plan.RequiresConfirmation ? "预检发现需要确认的项目" : "预检完成");
                if (!ConfirmProvinceStage1LedgerUpdate(options, plan))
                {
                    AddLog("已取消重庆阶段一台账更新。", "重庆");
                    ResetProgress("等待执行", "已取消重庆台账更新");
                    return;
                }

                SetStepDone(0);
                SetProgress(40, "写入重庆台账副本");
                StageWorkflowResult<ProvinceStage1LedgerUpdateResult> result = null;
                await Task.Run(() =>
                {
                    result = workflow.UpdateProvinceStage1Ledger(options, LogThreadSafe);
                });

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "重庆台账更新完成");
                LogSummary(result.SummaryLines);

                SetStage1ResultSuccess(result.Report.UpdatedPowerRows + " 个客户");
                ShowCompletion("重庆台账更新完成", "台账副本和更新报告已生成", options.OutputDirectory);
            }
            catch (Exception ex)
            {
                SetStepFailed();
                SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async void RunStage2_Click(object sender, RoutedEventArgs e)
        {
            if (_inputController.SelectedProvinceOrNull() == ProvinceCode.Chongqing)
            {
                await RunChongqingStage2PreflightAsync();
                return;
            }

            Stage2Options options;
            try
            {
                options = _inputController.CreateStage2Options();
                SaveInputs();
                if (!ConfirmRun("阶段二：生成分表和汇总表", options.Month, options.OutputDirectory))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
                return;
            }

            SetBusy(true);
            ResetResults();
            ResetProgress("正在执行阶段二...", "预检上月关键字段变化");
            SetProgress(8, "检查输入文件");
            SetStepRunning(0);
            AddLog("开始执行阶段二。", "阶段二");

            try
            {
                var workflow = CreateWorkflow();
                Stage2WorkflowPlan plan = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() => SetProgress(16, "读取台账并比对上月模板"));
                    plan = workflow.PlanStage2(options);
                });

                var preflight = plan.Preflight;
                SetProgress(24, preflight.HasIssues ? "预检发现需要确认的变化" : "预检完成");
                var confirmed = true;
                if (plan.RequiresConfirmation)
                {
                    SetStatus("待确认", "WarningBrush", "StatusBusyBrush");
                    SetStepNeedsConfirmation(0);
                    AddLog("阶段二预检发现 " + preflight.Issues.Count + " 条需要确认的变化。", "阶段二");
                    confirmed = ConfirmStage2Preflight(preflight, options);
                    if (confirmed)
                    {
                        SetStatus("运行中", "WarningBrush", "StatusBusyBrush");
                    }
                }
                else
                {
                    AddLog("阶段二预检通过，未发现需要确认的变化。", "阶段二");
                }

                if (!confirmed)
                {
                    var cancelled = workflow.CompleteStage2(plan, confirmed, LogThreadSafe);
                    if (cancelled.WasCancelled)
                    {
                        AddLog("已取消阶段二生成。", "阶段二");
                        ResetProgress("等待执行", "已取消阶段二");
                        return;
                    }
                }

                SetStepDone(0);
                SetProgress(34, "读取人工整理后的台账");
                Stage2WorkflowResult result = null;
                await Task.Run(() =>
                {
                    result = workflow.CompleteStage2(plan, confirmed, LogThreadSafe);
                });

                if (result.WasCancelled)
                {
                    AddLog("已取消阶段二生成。", "阶段二");
                    ResetProgress("等待执行", "已取消阶段二");
                    return;
                }

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "阶段二执行完成");
                LogSummary(result.SummaryLines);

                SetStage2ResultSuccess(result.Report.ProxyGroups + " 个文件", result.Report.IntermediaryGroups + " 个文件", "1 个文件");
                ShowCompletion("阶段二执行完成", "分表和汇总表已生成", options.OutputDirectory);
            }
            catch (Exception ex)
            {
                SetStepFailed();
                SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async Task RunChongqingStage2PreflightAsync()
        {
            ChongqingStage2Options options;
            try
            {
                options = _inputController.CreateChongqingStage2Options();
                SaveInputs();
                var message = new StringBuilder();
                message.AppendLine("结算月份：2026年" + options.Month + "月");
                message.AppendLine("输出文件夹：");
                message.AppendLine(options.OutputDirectory);
                message.AppendLine();
                message.AppendLine("将先读取重庆台账、代理/居间/退补模板和汇总表进行预检；确认后生成输出分表、退补表和汇总表副本。");
                if (!ConfirmAction("确认重庆阶段二生成", "即将执行重庆阶段二生成", message.ToString(), "开始生成"))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
                return;
            }

            SetBusy(true);
            ResetResults();
            ResetProgress("正在执行重庆阶段二...", "检查输入文件");
            SetProgress(8, "检查输入文件");
            SetStepRunning(0);
            AddLog("开始执行重庆阶段二。", "重庆阶段二");

            try
            {
                var workflow = CreateWorkflow();
                ChongqingStage2WorkflowPlan plan = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() => SetProgress(16, "读取重庆台账和汇总表模板"));
                    plan = workflow.PlanChongqingStage2(options);
                });

                var preflight = plan.Preflight;
                SetProgress(60, preflight.HasIssues ? "预检发现需要确认的变化" : "预检完成");
                var confirmed = true;
                if (plan.RequiresConfirmation)
                {
                    SetStatus("待确认", "WarningBrush", "StatusBusyBrush");
                    SetStepNeedsConfirmation(0);
                    AddLog("重庆阶段二预检发现 " + preflight.Issues.Count + " 条需要确认的变化。", "重庆阶段二");
                    confirmed = ConfirmChongqingStage2Preflight(preflight, options);
                    if (confirmed)
                    {
                        SetStatus("运行中", "WarningBrush", "StatusBusyBrush");
                    }
                }
                else
                {
                    AddLog("重庆阶段二预检通过，未发现需要确认的变化。", "重庆阶段二");
                }

                if (!confirmed)
                {
                    var cancelled = workflow.CompleteChongqingStage2(plan, confirmed, LogThreadSafe);
                    if (cancelled.WasCancelled)
                    {
                        AddLog("已取消重庆阶段二生成。", "重庆阶段二");
                        ResetProgress("等待执行", "已取消重庆阶段二");
                        return;
                    }
                }

                SetStepDone(0);
                SetProgress(34, "生成重庆分表和汇总表");
                ChongqingStage2WorkflowResult result = null;
                await Task.Run(() =>
                {
                    result = workflow.CompleteChongqingStage2(plan, confirmed, LogThreadSafe);
                });

                if (result.WasCancelled)
                {
                    AddLog("已取消重庆阶段二生成。", "重庆阶段二");
                    ResetProgress("等待执行", "已取消重庆阶段二");
                    return;
                }

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "重庆阶段二执行完成");
                LogSummary(result.Completed.SummaryLines);

                SetStage2ResultSuccess(
                    result.Report.ProxyGroups + " 个文件",
                    "居间" + result.Report.IntermediaryGroups + "/退补" + result.Report.RefundGroups,
                    "1 个文件");
                ShowCompletion("重庆阶段二执行完成", "分表、退补表和汇总表已生成", options.OutputDirectory);
            }
            catch (Exception ex)
            {
                SetStepFailed();
                SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async void RunEmployeeReward_Click(object sender, RoutedEventArgs e)
        {
            EmployeeRewardOptions options;
            try
            {
                options = _inputController.CreateEmployeeRewardOptions();
                SaveInputs();
                if (!ConfirmEmployeeRewardRun(options))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
                return;
            }

            SetBusy(true);
            ResetResults();
            ResetProgress("正在生成员工电量奖励...", "检查输入文件");
            SetProgress(10, "检查输入文件");
            SetStepRunning(0);
            AddLog("开始生成员工电量奖励表。", "员工奖励");

            try
            {
                StageWorkflowResult<EmployeeRewardResult> result = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        SetStepDone(0);
                        SetStepRunning(1);
                        SetProgress(30, "读取售电结算台账");
                    });

                    result = CreateWorkflow().RunEmployeeReward(options, LogThreadSafe);
                });

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "员工电量奖励生成完成");
                LogSummary(result.SummaryLines);

                SetEmployeeRewardResultSuccess(result.Report.PersonalWorkbookPaths.Count + " 个", "1 个文件");
                ShowCompletion("员工电量奖励生成完成", "奖励总表、个人确认表和校验报告已生成", options.OutputDirectory);
            }
            catch (Exception ex)
            {
                SetStepFailed();
                SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private static SettlementWorkflow CreateWorkflow()
        {
            var gateway = new ClosedXmlSettlementExcelGateway();
            return new SettlementWorkflow(
                new HainanStage1Service(gateway),
                new HainanStage2Service(gateway),
                new EmployeeRewardService(gateway),
                new ProvinceStage1Service(gateway),
                new ChongqingStage2Service(gateway));
        }

        private void LogSummary(IEnumerable<string> summaryLines)
        {
            var first = true;
            foreach (var line in summaryLines)
            {
                AddLog(line, first ? "成功" : "信息");
                first = false;
            }
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

        private bool ConfirmRun(string stageName, int month, string outputDirectory)
        {
            return _dialogController.ConfirmRun(stageName, month, outputDirectory);
        }

        private bool ConfirmEmployeeRewardRun(EmployeeRewardOptions options)
        {
            var period = options.StartMonth == options.EndMonth
                ? "2026年" + options.StartMonth + "月"
                : "2026年" + options.StartMonth + "-" + options.EndMonth + "月";
            var message = "即将生成员工电量奖励表。\n\n期间：" + period + "\n输出文件夹：\n" + options.OutputDirectory;
            return ConfirmAction("确认生成员工电量奖励", "即将生成员工电量奖励", message, "开始生成");
        }

        private bool ConfirmStage2Preflight(Stage2PreflightReport report, Stage2Options options)
        {
            var dialog = new Stage2PreflightWindow(report)
            {
                Owner = this
            };
            var confirmed = dialog.ShowDialog() == true;
            if (confirmed)
            {
                options.SummarySubjectDecisions.Clear();
                options.SummarySubjectDecisions.AddRange(dialog.SummarySubjectDecisions);
            }

            return confirmed;
        }

        private bool ConfirmChongqingStage2Preflight(ChongqingStage2PreflightReport report, ChongqingStage2Options options)
        {
            var dialog = new ChongqingStage2PreflightWindow(report)
            {
                Owner = this
            };
            var confirmed = dialog.ShowDialog() == true;
            if (confirmed)
            {
                options.SummarySubjectDecisions.Clear();
                options.SummarySubjectDecisions.AddRange(dialog.SummarySubjectDecisions);
            }

            return confirmed;
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
            var employeeRewardSelected = MainTabControl != null && MainTabControl.SelectedIndex == 1;
            var monthEnabled = !_isBusy && !employeeRewardSelected;

            MonthCombo.IsEnabled = monthEnabled;
            SettlementMonthLabel.Foreground = monthEnabled ? BrushOf("FieldTextBrush") : BrushOf("MutedBrush");
            UpdateProvinceUi();
            if (employeeRewardSelected)
            {
                SharedSettingsCaption.Text = "员工电量奖励使用本页的开始/结束月份，输出仍保存到这个文件夹中";
            }
        }

        private void UpdateProvinceUi()
        {
            if (_inputController == null || ProvinceCombo == null || StageOnePanel == null)
            {
                return;
            }

            var province = _inputController.SelectedProvinceOrNull();
            var hasProvince = province.HasValue;
            var profile = _inputController.SelectedProfileOrNull();
            var isChongqing = province == ProvinceCode.Chongqing;
            if (hasProvince && !profile.SupportsEmployeeReward && MainTabControl.SelectedItem == EmployeeRewardTab)
            {
                MainTabControl.SelectedItem = MainSettlementTab;
            }

            MainSettlementTab.Header = hasProvince ? profile.MainSettlementTabHeader : "结算流程";
            EmployeeRewardTab.Visibility = hasProvince && profile.SupportsEmployeeReward ? Visibility.Visible : Visibility.Collapsed;
            ProvinceEmptyPanel.Visibility = hasProvince ? Visibility.Collapsed : Visibility.Visible;
            StageOnePanel.Visibility = hasProvince ? Visibility.Visible : Visibility.Collapsed;
            StageTwoPanel.Visibility = hasProvince && profile.SupportsStage2 ? Visibility.Visible : Visibility.Collapsed;
            Grid.SetColumnSpan(StageOnePanel, hasProvince && profile.SupportsStage2 ? 1 : 2);
            StageOnePanel.Margin = hasProvince && profile.SupportsStage2 ? new Thickness(0, 0, 8, 0) : new Thickness(0);

            Stage1TitleText.Text = hasProvince ? profile.StageOneTitle : "请先选择结算省份";
            Stage1CaptionText.Text = !hasProvince
                ? "不同省份的结算口径不同，必须先选择省份再执行。"
                : profile.StageOneCaption;
            BaseLedgerLabel.Text = hasProvince ? profile.BaseLedgerLabel : "基础台账（先选择省份）";
            PowerLabel.Text = hasProvince ? profile.ExistingPowerLabel : "电量处理表";
            RawDetailLabel.Text = hasProvince ? profile.RawDetailLabel : "电量文件（先选择省份）";
            ReferenceLedgerLabel.Text = hasProvince ? profile.ReferenceLedgerLabel : "参考台账（可选）";
            RunStage1ButtonText.Text = hasProvince ? profile.RunStageOneButtonText : "开始 执行阶段一";
            CleanPowerButton.Content = hasProvince ? profile.CleanPowerButtonText : "只清洗电量";
            Stage2TitleText.Text = isChongqing ? "阶段二：重庆结算生成" : "阶段二：生成分表和汇总表";
            Stage2CaptionText.Text = isChongqing
                ? "生成代理/居间/退补分表和汇总表，生成前先确认预检项目"
                : "生成代理/居间分表和汇总表，输出结算结果";
            RunStage2ButtonText.Text = isChongqing ? "开始 重庆阶段二" : "开始 执行阶段二";

            var ledgerVisibility = hasProvince ? Visibility.Visible : Visibility.Collapsed;
            var existingPowerVisibility = hasProvince && profile.ShowsExistingPowerInput ? Visibility.Visible : Visibility.Collapsed;
            var referenceLedgerVisibility = hasProvince && profile.ShowsReferenceLedgerInput ? Visibility.Visible : Visibility.Collapsed;
            var chongqingStage2Visibility = isChongqing ? Visibility.Visible : Visibility.Collapsed;
            BaseLedgerLabel.Visibility = ledgerVisibility;
            BaseLedgerRow.Visibility = ledgerVisibility;
            PowerLabel.Visibility = existingPowerVisibility;
            PowerRow.Visibility = existingPowerVisibility;
            ReferenceLedgerLabel.Visibility = referenceLedgerVisibility;
            ReferenceLedgerRow.Visibility = referenceLedgerVisibility;
            CopyReferenceExistingCheckBox.Visibility = referenceLedgerVisibility;
            RefundTemplateDirLabel.Visibility = chongqingStage2Visibility;
            RefundTemplateDirRow.Visibility = chongqingStage2Visibility;
            AllowMissingOwnerCheckBox.Visibility = isChongqing ? Visibility.Collapsed : Visibility.Visible;

            RunStage1Button.IsEnabled = !_isBusy && hasProvince && profile.SupportsStage1LedgerUpdate;
            CleanPowerButton.IsEnabled = !_isBusy && hasProvince && profile.SupportsStage1CleanPower;
            RunStage2Button.IsEnabled = !_isBusy && hasProvince && profile.SupportsStage2;
            RunEmployeeRewardButton.IsEnabled = !_isBusy && hasProvince && profile.SupportsEmployeeReward;
            UpdateResultVisibility(province);
            SharedSettingsCaption.Text = !hasProvince
                ? "请先选择结算省份；选择后会显示对应省份的可用功能"
                : profile.SharedSettingsCaption;
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

        private void ResetProgress(string title, string description)
        {
            _progressController.ResetProgress(title, description);
        }

        private void SetProgress(int value, string description)
        {
            _progressController.SetProgress(value, description);
        }

        private void SetStepWaiting(int index)
        {
            _progressController.SetStepWaiting(index);
        }

        private void SetStepRunning(int index)
        {
            _progressController.SetStepRunning(index);
        }

        private void SetStepNeedsConfirmation(int index)
        {
            _progressController.SetStepNeedsConfirmation(index);
        }

        private void SetStepDone(int index)
        {
            _progressController.SetStepDone(index);
        }

        private void SetStepFailed()
        {
            _progressController.SetStepFailed();
        }

        private void UpdateResultVisibility(ProvinceCode? province)
        {
            _resultController.UpdateResultVisibility(province);
        }

        private void SetCompletionWaiting()
        {
            _resultController.ShowWaiting(_inputController.SelectedProvinceOrNull());
        }

        private void ResetResults()
        {
            _resultController.Reset(_inputController.SelectedProvinceOrNull());
        }

        private void ShowCompletion(string title, string detail, string outputDirectory)
        {
            _resultController.ShowCompletion(title, detail, outputDirectory);
        }

        private void SetStage1ResultSuccess(string countText)
        {
            _resultController.SetStage1Success(countText);
        }

        private void SetStage2ResultSuccess(string proxyCountText, string intermediaryCountText, string summaryCountText)
        {
            _resultController.SetStage2Success(proxyCountText, intermediaryCountText, summaryCountText);
        }

        private void SetStage2PreflightSuccess()
        {
            _resultController.SetStage2PreflightSuccess();
        }

        private void SetEmployeeRewardResultSuccess(string personalCountText, string summaryCountText)
        {
            _resultController.SetEmployeeRewardSuccess(personalCountText, summaryCountText);
        }

        private void LoadSavedInputs(UserInputSnapshot snapshot)
        {
            _loadingInputs = true;
            _inputController.LoadSavedInputs(snapshot);
            _loadingInputs = false;

            if (_inputController.HasSavedInputs(snapshot))
            {
                AddLog("已载入上次选择的文件路径；结算月份仍需手动选择。", "信息");
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

        private void LogThreadSafe(string message)
        {
            Dispatcher.Invoke(() => AddLog(message, "信息"));
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

        private bool ConfirmAction(string title, string heading, string message, string primaryButtonText)
        {
            return _dialogController.ConfirmAction(title, heading, message, primaryButtonText);
        }

        private bool ConfirmProvinceStage1LedgerUpdate(ProvinceStage1LedgerUpdateOptions options, ProvinceStage1LedgerUpdatePlan plan)
        {
            var message = new StringBuilder();
            message.AppendLine("结算月份：2026年" + options.Month + "月");
            message.AppendLine("匹配客户：" + plan.MatchedRows + " / " + plan.PowerCustomerRows);
            message.AppendLine("多户号客户：" + plan.MultiAccountRows + " 行（仅提示，不写入B列）");
            message.AppendLine("输出文件夹：");
            message.AppendLine(options.OutputDirectory);

            if (plan.RequiresConfirmation)
            {
                var dialog = new ProvinceStage1LedgerPreflightWindow(options, plan)
                {
                    Owner = this
                };

                if (dialog.ShowDialog() == true)
                {
                    options.CustomerDecisions = dialog.CustomerDecisions;
                    options.ManualCustomerMatches = dialog.ManualCustomerMatches;
                    return true;
                }

                return false;
            }

            message.AppendLine();
            message.AppendLine("未发现匹配异常。确认后会生成台账副本，不会覆盖原文件。");
            return ConfirmAction("确认重庆台账更新", "即将写入重庆台账副本", message.ToString(), "开始写入");
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

        private Brush BrushOf(string key)
        {
            return (Brush)FindResource(key);
        }
    }
}
