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
using Ookii.Dialogs.Wpf;

namespace HainanSettlementTool.Wpf
{
    public partial class MainWindow : Window
    {
        private readonly TextBlock[] _stepTexts;
        private readonly TextBlock[] _stepStatuses;
        private string _lastOutputDirectory;
        private bool _isBusy;
        private bool _loadingInputs;
        private string _themeMode = ThemeService.SystemMode;

        public MainWindow()
        {
            var snapshot = UserInputStore.Load();
            _themeMode = ThemeService.NormalizeMode(snapshot.ThemeMode);
            ThemeService.Apply(_themeMode);

            InitializeComponent();

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
            _stepTexts = new[] { Step1Text, Step2Text, Step3Text, Step4Text, Step5Text };
            _stepStatuses = new[] { Step1Status, Step2Status, Step3Status, Step4Status, Step5Status };
            ResetProgress("等待执行", "尚未开始");
            ResetResults();
            AddLog("工具已就绪，等待操作。", "信息");
            LoadSavedInputs(snapshot);
            UpdateSharedSettingsState();
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
            var title = SelectedProvinceOrNull() == ProvinceCode.Chongqing
                ? "选择重庆售电结算台账"
                : "选择基础台账";
            BrowseExcel(BaseLedgerBox, title);
        }

        private void BrowsePower_Click(object sender, RoutedEventArgs e)
        {
            BrowseExcel(PowerBox, "选择电量处理表");
        }

        private void BrowseRawDetail_Click(object sender, RoutedEventArgs e)
        {
            var province = SelectedProvinceOrNull();
            if (!province.HasValue)
            {
                ShowErrorMessage("请选择结算省份后再选择电量文件。");
                return;
            }

            var title = province.Value == ProvinceCode.Chongqing
                ? "选择重庆交易中心电量确认结算单"
                : "选择原始零售侧明细";
            BrowseFile(RawDetailBox, title, "Excel/CSV|*.xlsx;*.xls;*.csv|Excel 文件|*.xlsx;*.xls|CSV 文件|*.csv|所有文件|*.*");
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
            BrowseFile(target, title, "Excel 文件|*.xlsx|所有文件|*.*");
        }

        private void BrowseFile(TextBox target, string title, string filter)
        {
            var dialog = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                CheckFileExists = true
            };

            if (dialog.ShowDialog(this) == true)
            {
                target.Text = dialog.FileName;
                SaveInputs();
            }
        }

        private void BrowseFolder(TextBox target, string title)
        {
            var dialog = new VistaFolderBrowserDialog
            {
                Description = title,
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true
            };

            if (Directory.Exists(target.Text))
            {
                dialog.SelectedPath = target.Text;
            }

            if (dialog.ShowDialog(this) == true)
            {
                target.Text = dialog.SelectedPath;
                SaveInputs();
            }
        }

        private async void RunStage1_Click(object sender, RoutedEventArgs e)
        {
            var province = SelectedProvinceOrNull();
            if (!province.HasValue)
            {
                ShowErrorMessage("请选择结算省份后再执行阶段一。");
                return;
            }

            if (province.Value == ProvinceCode.Chongqing)
            {
                await RunProvinceStage1LedgerUpdateAsync();
                return;
            }

            Stage1Options options;
            try
            {
                options = CreateStage1Options();
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

                Stage1ResultStatus.Text = "成功";
                Stage1ResultCount.Text = "1 个文件";
                FinishedAtText.Text = DateTime.Now.ToString("HH:mm:ss");
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
            var province = SelectedProvinceOrNull();
            if (!province.HasValue)
            {
                ShowErrorMessage("请选择结算省份后再清洗电量数据。");
                return;
            }

            if (province.Value == ProvinceCode.Chongqing)
            {
                await RunProvinceStage1CleanPowerAsync();
                return;
            }

            string rawDetailPath;
            string outputPath;
            string outputDirectory;
            try
            {
                rawDetailPath = RawDetailBox.Text.Trim();
                outputPath = ResolvePowerOutputPath(rawDetailPath);
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

                Stage1ResultStatus.Text = "成功";
                Stage1ResultCount.Text = result.Report.PowerRows + " 个客户";
                FinishedAtText.Text = DateTime.Now.ToString("HH:mm:ss");
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
                options = CreateProvinceStage1CleanOptions();
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

                Stage1ResultStatus.Text = "成功";
                Stage1ResultCount.Text = result.Report.CustomerRows + " 个客户";
                FinishedAtText.Text = DateTime.Now.ToString("HH:mm:ss");
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
                options = CreateProvinceStage1LedgerUpdateOptions();
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

                Stage1ResultStatus.Text = "成功";
                Stage1ResultCount.Text = result.Report.UpdatedPowerRows + " 个客户";
                FinishedAtText.Text = DateTime.Now.ToString("HH:mm:ss");
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
            Stage2Options options;
            try
            {
                options = CreateStage2Options();
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
                    confirmed = ConfirmStage2Preflight(preflight);
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

                ProxyResultStatus.Text = "成功";
                ProxyResultCount.Text = result.Report.ProxyGroups + " 个文件";
                IntermediaryResultStatus.Text = "成功";
                IntermediaryResultCount.Text = result.Report.IntermediaryGroups + " 个文件";
                SummaryResultStatus.Text = "成功";
                SummaryResultCount.Text = "1 个文件";
                FinishedAtText.Text = DateTime.Now.ToString("HH:mm:ss");
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

        private async void RunEmployeeReward_Click(object sender, RoutedEventArgs e)
        {
            EmployeeRewardOptions options;
            try
            {
                options = CreateEmployeeRewardOptions();
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

                EmployeeRewardResultStatus.Text = "成功";
                EmployeeRewardResultCount.Text = result.Report.PersonalWorkbookPaths.Count + " 个";
                SummaryResultStatus.Text = "成功";
                SummaryResultCount.Text = "1 个文件";
                FinishedAtText.Text = DateTime.Now.ToString("HH:mm:ss");
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

        private Stage1Options CreateStage1Options()
        {
            var powerPath = PowerBox.Text.Trim();
            var rawDetailPath = RawDetailBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(rawDetailPath))
            {
                powerPath = ResolvePowerOutputPath(rawDetailPath);
                PowerBox.Text = powerPath;
            }

            return new Stage1Options
            {
                Month = SelectedMonth(),
                BaseLedgerPath = BaseLedgerBox.Text.Trim(),
                PowerPath = powerPath,
                RawDetailPath = rawDetailPath,
                ReferenceLedgerPath = ReferenceLedgerBox.Text.Trim(),
                OutputDirectory = OutputDirBox.Text.Trim(),
                CopyReferenceExisting = CopyReferenceExistingCheckBox.IsChecked == true
            };
        }

        private string ResolvePowerOutputPath(string rawDetailPath)
        {
            if (string.IsNullOrWhiteSpace(rawDetailPath))
            {
                throw new InvalidOperationException("请选择原始零售侧明细。");
            }

            var outputDirectory = OutputDirBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                throw new InvalidOperationException("请选择结果输出文件夹。");
            }

            return Path.Combine(outputDirectory, "零售侧用户电量数据处理表.xlsx");
        }

        private Stage2Options CreateStage2Options()
        {
            return new Stage2Options
            {
                Month = SelectedMonth(),
                LedgerPath = CompletedLedgerBox.Text.Trim(),
                ProxyTemplateDirectory = ProxyTemplateDirBox.Text.Trim(),
                IntermediaryTemplateDirectory = IntermediaryTemplateDirBox.Text.Trim(),
                SummaryTemplatePath = SummaryTemplateBox.Text.Trim(),
                OutputDirectory = OutputDirBox.Text.Trim(),
                AllowMissingOwner = AllowMissingOwnerCheckBox.IsChecked == true
            };
        }

        private EmployeeRewardOptions CreateEmployeeRewardOptions()
        {
            var startMonth = SelectedRewardStartMonth();
            var endMonth = SelectedRewardEndMonth();
            if (startMonth > endMonth)
            {
                throw new InvalidOperationException("员工电量奖励开始月份不能晚于结束月份。");
            }

            return new EmployeeRewardOptions
            {
                Year = 2026,
                StartMonth = startMonth,
                EndMonth = endMonth,
                LedgerPath = RewardLedgerBox.Text.Trim(),
                OutputDirectory = OutputDirBox.Text.Trim()
            };
        }

        private ProvinceStage1CleanOptions CreateProvinceStage1CleanOptions()
        {
            return new ProvinceStage1CleanOptions
            {
                Province = SelectedProvince(),
                Month = SelectedMonthOrZero(),
                RawDetailPath = RawDetailBox.Text.Trim(),
                OutputDirectory = OutputDirBox.Text.Trim()
            };
        }

        private ProvinceStage1LedgerUpdateOptions CreateProvinceStage1LedgerUpdateOptions()
        {
            return new ProvinceStage1LedgerUpdateOptions
            {
                Province = SelectedProvince(),
                Month = SelectedMonth(),
                LedgerPath = BaseLedgerBox.Text.Trim(),
                RawDetailPath = RawDetailBox.Text.Trim(),
                OutputDirectory = OutputDirBox.Text.Trim()
            };
        }

        private static SettlementWorkflow CreateWorkflow()
        {
            var gateway = new ClosedXmlStage1ExcelGateway();
            return new SettlementWorkflow(
                new Stage1Service(gateway),
                new Stage2Service(gateway),
                new EmployeeRewardService(gateway),
                new ProvinceStage1Service(gateway));
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

        private int SelectedMonth()
        {
            if (MonthCombo.SelectedIndex < 0)
            {
                throw new InvalidOperationException("请选择结算月份。");
            }

            return MonthCombo.SelectedIndex + 2;
        }

        private int SelectedMonthOrZero()
        {
            return MonthCombo.SelectedIndex < 0 ? 0 : MonthCombo.SelectedIndex + 2;
        }

        private int SelectedRewardStartMonth()
        {
            if (RewardStartMonthCombo.SelectedIndex < 0)
            {
                throw new InvalidOperationException("请选择员工电量奖励开始月份。");
            }

            return RewardStartMonthCombo.SelectedIndex + 1;
        }

        private int SelectedRewardEndMonth()
        {
            if (RewardEndMonthCombo.SelectedIndex < 0)
            {
                throw new InvalidOperationException("请选择员工电量奖励结束月份。");
            }

            return RewardEndMonthCombo.SelectedIndex + 1;
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
            ProvinceCombo.Items.Add("海南");
            ProvinceCombo.Items.Add("重庆");
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

        private ProvinceCode SelectedProvince()
        {
            var province = SelectedProvinceOrNull();
            if (!province.HasValue)
            {
                throw new InvalidOperationException("请选择结算省份。");
            }

            return province.Value;
        }

        private ProvinceCode? SelectedProvinceOrNull()
        {
            switch (ProvinceCombo.SelectedIndex)
            {
                case 0:
                    return ProvinceCode.Hainan;
                case 1:
                    return ProvinceCode.Chongqing;
                default:
                    return null;
            }
        }

        private bool ConfirmRun(string stageName, int month, string outputDirectory)
        {
            var dialog = new ConfirmRunWindow(stageName, month, outputDirectory)
            {
                Owner = this
            };
            return dialog.ShowDialog() == true;
        }

        private bool ConfirmEmployeeRewardRun(EmployeeRewardOptions options)
        {
            var period = options.StartMonth == options.EndMonth
                ? "2026年" + options.StartMonth + "月"
                : "2026年" + options.StartMonth + "-" + options.EndMonth + "月";
            var message = "即将生成员工电量奖励表。\n\n期间：" + period + "\n输出文件夹：\n" + options.OutputDirectory;
            return ConfirmAction("确认生成员工电量奖励", "即将生成员工电量奖励", message, "开始生成");
        }

        private bool ConfirmStage2Preflight(Stage2PreflightReport report)
        {
            var dialog = new Stage2PreflightWindow(report)
            {
                Owner = this
            };
            return dialog.ShowDialog() == true;
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
            if (ProvinceCombo == null || StageOnePanel == null)
            {
                return;
            }

            var province = SelectedProvinceOrNull();
            var hasProvince = province.HasValue;
            var hainan = province == ProvinceCode.Hainan;
            var chongqing = province == ProvinceCode.Chongqing;
            if (chongqing && MainTabControl.SelectedItem == EmployeeRewardTab)
            {
                MainTabControl.SelectedItem = MainSettlementTab;
            }

            MainSettlementTab.Header = !hasProvince ? "结算流程" : chongqing ? "阶段一：台账更新" : "代理费结算";
            EmployeeRewardTab.Visibility = hainan ? Visibility.Visible : Visibility.Collapsed;
            StageTwoPanel.Visibility = hainan ? Visibility.Visible : Visibility.Collapsed;
            Grid.SetColumnSpan(StageOnePanel, hainan ? 1 : 2);
            StageOnePanel.Margin = hainan ? new Thickness(0, 0, 8, 0) : new Thickness(0);

            Stage1TitleText.Text = !hasProvince ? "请先选择结算省份" : chongqing ? "阶段一：台账更新" : "阶段一：写入电量到台账";
            Stage1CaptionText.Text = !hasProvince
                ? "不同省份的结算口径不同，必须先选择省份再执行。"
                : chongqing
                ? "清洗交易中心电量确认结算单，并写入重庆台账副本"
                : "整理电量并写入基础台账，输出检查报告";
            BaseLedgerLabel.Text = chongqing ? "重庆售电结算台账（必填）" : "基础台账（必填）";
            RawDetailLabel.Text = !hasProvince ? "电量文件（先选择省份）" : chongqing ? "交易中心电量确认结算单（必填）" : "原始零售侧明细";
            RunStage1ButtonText.Text = chongqing ? "清洗并更新台账" : "开始 执行阶段一";
            CleanPowerButton.Content = chongqing ? "只清洗电量数据" : "只清洗电量";

            var ledgerVisibility = hasProvince ? Visibility.Visible : Visibility.Collapsed;
            var hainanVisibility = hainan ? Visibility.Visible : Visibility.Collapsed;
            BaseLedgerLabel.Visibility = ledgerVisibility;
            BaseLedgerRow.Visibility = ledgerVisibility;
            PowerLabel.Visibility = hainanVisibility;
            PowerRow.Visibility = hainanVisibility;
            ReferenceLedgerLabel.Visibility = hainanVisibility;
            ReferenceLedgerRow.Visibility = hainanVisibility;
            CopyReferenceExistingCheckBox.Visibility = hainanVisibility;

            RunStage1Button.IsEnabled = !_isBusy && hasProvince;
            CleanPowerButton.IsEnabled = !_isBusy && hasProvince;
            RunStage2Button.IsEnabled = !_isBusy && hainan;
            RunEmployeeRewardButton.IsEnabled = !_isBusy && hainan;
            SharedSettingsCaption.Text = !hasProvince
                ? "请先选择结算省份；选择后会显示对应省份的可用功能"
                : chongqing
                ? "重庆当前开放阶段一的只清洗电量数据，输出仍保存到这个文件夹中"
                : "阶段一和阶段二生成的所有文件都会保存到这个文件夹中";
        }

        private void RefreshThemeDependentState()
        {
            UpdateSharedSettingsState();
            if (StatusText.Text == "待确认" || StatusText.Text == "运行中")
            {
                SetStatus(StatusText.Text, "WarningBrush", "StatusBusyBrush");
            }
            else
            {
                SetStatus(StatusText.Text, "SuccessBrush", "StatusReadyBrush");
            }
        }

        private void SetStatus(string text, string dotBrushKey, string backgroundBrushKey)
        {
            StatusText.Text = text;
            StatusDot.Fill = BrushOf(dotBrushKey);
            StatusPill.Background = BrushOf(backgroundBrushKey);
        }

        private void ResetProgress(string title, string description)
        {
            ProgressTitle.Text = title;
            ProgressDescriptionText.Text = description;
            SetProgress(0, description);
            for (var i = 0; i < _stepTexts.Length; i++)
            {
                SetStepWaiting(i);
            }
        }

        private void SetProgress(int value, string description)
        {
            ProgressBar.Value = value;
            ProgressPercentText.Text = value + "%";
            ProgressDescriptionText.Text = description;
        }

        private void SetStepWaiting(int index)
        {
            _stepTexts[index].Text = "○  " + StepName(index);
            _stepTexts[index].Foreground = BrushOf("MutedBrush");
            _stepStatuses[index].Text = "等待中";
            _stepStatuses[index].Foreground = BrushOf("MutedBrush");
        }

        private void SetStepRunning(int index)
        {
            _stepTexts[index].Text = "●  " + StepName(index);
            _stepTexts[index].Foreground = BrushOf("AccentBrush");
            _stepStatuses[index].Text = "进行中";
            _stepStatuses[index].Foreground = BrushOf("AccentBrush");
        }

        private void SetStepNeedsConfirmation(int index)
        {
            _stepTexts[index].Text = "●  " + StepName(index);
            _stepTexts[index].Foreground = BrushOf("WarningBrush");
            _stepStatuses[index].Text = "待确认";
            _stepStatuses[index].Foreground = BrushOf("WarningBrush");
        }

        private void SetStepDone(int index)
        {
            _stepTexts[index].Text = "●  " + StepName(index);
            _stepTexts[index].Foreground = BrushOf("SuccessBrush");
            _stepStatuses[index].Text = "完成";
            _stepStatuses[index].Foreground = BrushOf("SuccessBrush");

            if (index + 1 < _stepTexts.Length)
            {
                SetStepRunning(index + 1);
                SetProgress(Math.Min(90, 25 + (index + 1) * 15), StepName(index + 1));
            }
        }

        private void SetStepFailed()
        {
            for (var i = 0; i < _stepStatuses.Length; i++)
            {
                if (_stepStatuses[i].Text == "进行中")
                {
                    _stepTexts[i].Text = "●  " + StepName(i);
                    _stepTexts[i].Foreground = BrushOf("ErrorBrush");
                    _stepStatuses[i].Text = "失败";
                    _stepStatuses[i].Foreground = BrushOf("ErrorBrush");
                    return;
                }
            }
        }

        private static string StepName(int index)
        {
            switch (index)
            {
                case 0:
                    return "检查输入文件";
                case 1:
                    return "读取台账数据";
                case 2:
                    return "生成结算文件";
                case 3:
                    return "写入结果报告";
                default:
                    return "保存结果文件";
            }
        }

        private void ResetResults()
        {
            Stage1ResultStatus.Text = "等待";
            Stage1ResultCount.Text = "-";
            ProxyResultStatus.Text = "等待";
            ProxyResultCount.Text = "-";
            IntermediaryResultStatus.Text = "等待";
            IntermediaryResultCount.Text = "-";
            SummaryResultStatus.Text = "等待";
            SummaryResultCount.Text = "-";
            EmployeeRewardResultStatus.Text = "等待";
            EmployeeRewardResultCount.Text = "-";
            FinishedAtText.Text = "-";
            _lastOutputDirectory = null;
            CompletionTitleText.Text = "等待生成结果";
            CompletionDetailText.Text = "运行完成后会在这里显示输出位置";
            CompletionOutputText.Text = "尚未生成";
            CompletionCard.Visibility = Visibility.Visible;
        }

        private void ShowCompletion(string title, string detail, string outputDirectory)
        {
            _lastOutputDirectory = outputDirectory;
            CompletionTitleText.Text = title;
            CompletionDetailText.Text = detail;
            CompletionOutputText.Text = outputDirectory;
            CompletionCard.Visibility = Visibility.Visible;
        }

        private void LoadSavedInputs(UserInputSnapshot snapshot)
        {
            OutputDirBox.Text = snapshot.OutputDirectory ?? string.Empty;
            BaseLedgerBox.Text = snapshot.BaseLedgerPath ?? string.Empty;
            PowerBox.Text = snapshot.PowerPath ?? string.Empty;
            RawDetailBox.Text = snapshot.RawDetailPath ?? string.Empty;
            ReferenceLedgerBox.Text = snapshot.ReferenceLedgerPath ?? string.Empty;
            CompletedLedgerBox.Text = snapshot.CompletedLedgerPath ?? string.Empty;
            ProxyTemplateDirBox.Text = snapshot.ProxyTemplateDirectory ?? string.Empty;
            IntermediaryTemplateDirBox.Text = snapshot.IntermediaryTemplateDirectory ?? string.Empty;
            SummaryTemplateBox.Text = snapshot.SummaryTemplatePath ?? string.Empty;
            RewardLedgerBox.Text = snapshot.RewardLedgerPath ?? string.Empty;

            if (HasSavedInputs(snapshot))
            {
                AddLog("已载入上次选择的文件路径；结算月份仍需手动选择。", "信息");
            }
        }

        private void SaveInputs()
        {
            try
            {
                UserInputStore.Save(new UserInputSnapshot
                {
                    OutputDirectory = OutputDirBox.Text.Trim(),
                    BaseLedgerPath = BaseLedgerBox.Text.Trim(),
                    PowerPath = PowerBox.Text.Trim(),
                    RawDetailPath = RawDetailBox.Text.Trim(),
                    ReferenceLedgerPath = ReferenceLedgerBox.Text.Trim(),
                    CompletedLedgerPath = CompletedLedgerBox.Text.Trim(),
                    ProxyTemplateDirectory = ProxyTemplateDirBox.Text.Trim(),
                    IntermediaryTemplateDirectory = IntermediaryTemplateDirBox.Text.Trim(),
                    SummaryTemplatePath = SummaryTemplateBox.Text.Trim(),
                    RewardLedgerPath = RewardLedgerBox.Text.Trim(),
                    ProvinceCode = SelectedProvinceOrNull()?.ToString() ?? string.Empty,
                    ThemeMode = _themeMode
                });
            }
            catch (Exception ex)
            {
                AddLog("保存上次路径失败：" + ex.Message, "提示");
            }
        }

        private static bool HasSavedInputs(UserInputSnapshot snapshot)
        {
            return !string.IsNullOrWhiteSpace(snapshot.OutputDirectory)
                || !string.IsNullOrWhiteSpace(snapshot.BaseLedgerPath)
                || !string.IsNullOrWhiteSpace(snapshot.PowerPath)
                || !string.IsNullOrWhiteSpace(snapshot.RawDetailPath)
                || !string.IsNullOrWhiteSpace(snapshot.ReferenceLedgerPath)
                || !string.IsNullOrWhiteSpace(snapshot.CompletedLedgerPath)
                || !string.IsNullOrWhiteSpace(snapshot.ProxyTemplateDirectory)
                || !string.IsNullOrWhiteSpace(snapshot.IntermediaryTemplateDirectory)
                || !string.IsNullOrWhiteSpace(snapshot.SummaryTemplatePath)
                || !string.IsNullOrWhiteSpace(snapshot.RewardLedgerPath);
        }

        private void LogThreadSafe(string message)
        {
            Dispatcher.Invoke(() => AddLog(message, "信息"));
        }

        private void AddLog(string message, string level)
        {
            LogBox.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] [" + level + "] " + message + Environment.NewLine);
            LogBox.ScrollToEnd();
        }

        private void ShowError(Exception ex)
        {
            ShowErrorMessage(ex.Message);
        }

        private void ShowErrorMessage(string message)
        {
            var dialog = new ModernDialogWindow("出错了", "需要先处理这个问题", message, "知道了", null, ModernDialogKind.Error)
            {
                Owner = this
            };
            dialog.ShowDialog();
        }

        private bool ConfirmAction(string title, string heading, string message, string primaryButtonText)
        {
            var dialog = new ModernDialogWindow(title, heading, message, primaryButtonText, "取消", ModernDialogKind.Warning)
            {
                Owner = this
            };
            return dialog.ShowDialog() == true;
        }

        private bool ConfirmProvinceStage1LedgerUpdate(ProvinceStage1LedgerUpdateOptions options, ProvinceStage1LedgerUpdatePlan plan)
        {
            var message = new StringBuilder();
            message.AppendLine("结算月份：2026年" + options.Month + "月");
            message.AppendLine("匹配客户：" + plan.MatchedRows + " / " + plan.PowerCustomerRows);
            message.AppendLine("预计补齐电力用户编码：" + plan.CodeFillRows + " 行");
            message.AppendLine("输出文件夹：");
            message.AppendLine(options.OutputDirectory);

            if (plan.RequiresConfirmation)
            {
                message.AppendLine();
                message.AppendLine("预检发现以下项目，请确认后再继续：");
                foreach (var issue in plan.Issues.Take(12))
                {
                    var customer = string.IsNullOrWhiteSpace(issue.CustomerName) ? string.Empty : "：" + issue.CustomerName;
                    message.AppendLine("- " + issue.Category + customer + "；" + issue.Message);
                }

                if (plan.Issues.Count > 12)
                {
                    message.AppendLine("- 其余 " + (plan.Issues.Count - 12) + " 项会写入更新报告。");
                }

                return ConfirmAction("确认重庆台账更新", "预检发现需要确认的项目", message.ToString(), "继续写入副本");
            }

            message.AppendLine();
            message.AppendLine("未发现匹配异常。确认后会生成台账副本，不会覆盖原文件。");
            return ConfirmAction("确认重庆台账更新", "即将写入重庆台账副本", message.ToString(), "开始写入");
        }

        private void ClearStage1_Click(object sender, RoutedEventArgs e)
        {
            BaseLedgerBox.Clear();
            PowerBox.Clear();
            RawDetailBox.Clear();
            ReferenceLedgerBox.Clear();
            CopyReferenceExistingCheckBox.IsChecked = false;
            SaveInputs();
        }

        private void ClearStage2_Click(object sender, RoutedEventArgs e)
        {
            CompletedLedgerBox.Clear();
            ProxyTemplateDirBox.Clear();
            IntermediaryTemplateDirBox.Clear();
            SummaryTemplateBox.Clear();
            AllowMissingOwnerCheckBox.IsChecked = false;
            SaveInputs();
        }

        private void ClearEmployeeReward_Click(object sender, RoutedEventArgs e)
        {
            RewardLedgerBox.Clear();
            RewardStartMonthCombo.SelectedIndex = 0;
            RewardEndMonthCombo.SelectedIndex = -1;
            SaveInputs();
        }

        private void ClearLog_Click(object sender, RoutedEventArgs e)
        {
            LogBox.Clear();
        }

        private void SaveLog_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Title = "保存运行日志",
                Filter = "文本文件|*.txt|所有文件|*.*",
                FileName = "海南售电结算运行日志.txt"
            };

            if (dialog.ShowDialog(this) == true)
            {
                File.WriteAllText(dialog.FileName, LogBox.Text, Encoding.UTF8);
            }
        }

        private void OpenOutputDir_Click(object sender, RoutedEventArgs e)
        {
            var path = !string.IsNullOrWhiteSpace(_lastOutputDirectory) ? _lastOutputDirectory : OutputDirBox.Text.Trim();
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
