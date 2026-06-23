using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
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

        public MainWindow()
        {
            InitializeComponent();

            for (var month = 2; month <= 12; month++)
            {
                MonthCombo.Items.Add("2026年" + month + "月");
            }

            MonthCombo.SelectedIndex = -1;
            _stepTexts = new[] { Step1Text, Step2Text, Step3Text, Step4Text, Step5Text };
            _stepStatuses = new[] { Step1Status, Step2Status, Step3Status, Step4Status, Step5Status };
            ResetProgress("等待执行", "尚未开始");
            ResetResults();
            AddLog("工具已就绪，等待操作。", "信息");
            LoadSavedInputs();
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

        private void ToggleMaximize()
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }

        private void BrowseOutputDir_Click(object sender, RoutedEventArgs e)
        {
            BrowseFolder(OutputDirBox, "选择结果输出文件夹");
        }

        private void BrowseBaseLedger_Click(object sender, RoutedEventArgs e)
        {
            BrowseExcel(BaseLedgerBox, "选择基础台账");
        }

        private void BrowsePower_Click(object sender, RoutedEventArgs e)
        {
            BrowseExcel(PowerBox, "选择电量处理表");
        }

        private void BrowseRawDetail_Click(object sender, RoutedEventArgs e)
        {
            BrowseFile(RawDetailBox, "选择原始零售侧明细", "Excel/CSV|*.xlsx;*.xls;*.csv|Excel 文件|*.xlsx;*.xls|CSV 文件|*.csv|所有文件|*.*");
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
                Stage1Report report = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        SetStepDone(0);
                        SetStepRunning(1);
                        SetProgress(28, "读取台账和电量文件");
                    });

                    var service = new Stage1Service(new ClosedXmlStage1ExcelGateway());
                    report = service.Run(options, LogThreadSafe);
                });

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "阶段一执行完成");
                AddLog("阶段一执行完成。", "成功");
                AddLog("输出台账：" + report.Output, "信息");
                AddLog("报告：" + report.ReportPath, "信息");

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
                if (MessageBox.Show(this, message, "确认清洗电量", MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK)
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
                PowerCleanReport report = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        SetStepDone(0);
                        SetStepRunning(1);
                        SetProgress(35, "读取原始零售侧明细");
                    });

                    var service = new Stage1Service(new ClosedXmlStage1ExcelGateway());
                    report = service.CleanPowerData(rawDetailPath, outputPath, LogThreadSafe);
                });

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "电量清洗完成");
                AddLog("电量清洗完成。", "成功");
                AddLog("电量处理表：" + report.OutputPath, "信息");
                AddLog("客户数量：" + report.PowerRows + "，合计电量：" + report.MonthTotal.ToString("0.####"), "信息");

                Stage1ResultStatus.Text = "成功";
                Stage1ResultCount.Text = report.PowerRows + " 个客户";
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
                Stage2PreflightReport preflight = null;
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() => SetProgress(16, "读取台账并比对上月模板"));
                    var preflightService = new Stage2Service(new ClosedXmlStage1ExcelGateway());
                    preflight = preflightService.Analyze(options);
                });

                SetProgress(24, preflight.HasIssues ? "预检发现需要确认的变化" : "预检完成");
                if (preflight.HasIssues)
                {
                    SetStatus("待确认", "WarningBrush", Color.FromRgb(255, 245, 224));
                    SetStepNeedsConfirmation(0);
                    AddLog("阶段二预检发现 " + preflight.Issues.Count + " 条需要确认的变化。", "阶段二");
                    if (!ConfirmStage2Preflight(preflight))
                    {
                        AddLog("已取消阶段二生成。", "阶段二");
                        ResetProgress("等待执行", "已取消阶段二");
                        return;
                    }

                    SetStatus("运行中", "WarningBrush", Color.FromRgb(255, 245, 224));
                }
                else
                {
                    AddLog("阶段二预检通过，未发现需要确认的变化。", "阶段二");
                }

                SetStepDone(0);
                SetProgress(34, "读取人工整理后的台账");
                Stage2Report report = null;
                await Task.Run(() =>
                {
                    var service = new Stage2Service(new ClosedXmlStage1ExcelGateway());
                    report = service.Run(options, LogThreadSafe);
                });

                SetStepDone(1);
                SetStepDone(2);
                SetStepDone(3);
                SetStepDone(4);
                SetProgress(100, "阶段二执行完成");
                AddLog("阶段二执行完成。", "成功");
                AddLog("汇总表：" + report.Summary, "信息");
                AddLog("报告：" + report.ReportPath, "信息");
                AddLog("代理费合计：" + report.ProxyTotal.ToString("0.####"), "信息");
                AddLog("居间费合计：" + report.IntermediaryTotal.ToString("0.####"), "信息");

                ProxyResultStatus.Text = "成功";
                ProxyResultCount.Text = report.ProxyGroups + " 个文件";
                IntermediaryResultStatus.Text = "成功";
                IntermediaryResultCount.Text = report.IntermediaryGroups + " 个文件";
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

        private int SelectedMonth()
        {
            if (MonthCombo.SelectedIndex < 0)
            {
                throw new InvalidOperationException("请选择结算月份。");
            }

            return MonthCombo.SelectedIndex + 2;
        }

        private bool ConfirmRun(string stageName, int month, string outputDirectory)
        {
            var dialog = new ConfirmRunWindow(stageName, month, outputDirectory)
            {
                Owner = this
            };
            return dialog.ShowDialog() == true;
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
            RunStage1Button.IsEnabled = !busy;
            CleanPowerButton.IsEnabled = !busy;
            RunStage2Button.IsEnabled = !busy;
            ClearStage1Button.IsEnabled = !busy;
            ClearStage2Button.IsEnabled = !busy;
            SetStatus(
                busy ? "运行中" : "就绪",
                busy ? "WarningBrush" : "SuccessBrush",
                busy ? Color.FromRgb(255, 245, 224) : Color.FromRgb(234, 247, 241));
        }

        private void SetStatus(string text, string dotBrushKey, Color background)
        {
            StatusText.Text = text;
            StatusDot.Fill = BrushOf(dotBrushKey);
            StatusPill.Background = new SolidColorBrush(background);
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
                    _stepTexts[i].Foreground = new SolidColorBrush(Color.FromRgb(192, 57, 43));
                    _stepStatuses[i].Text = "失败";
                    _stepStatuses[i].Foreground = new SolidColorBrush(Color.FromRgb(192, 57, 43));
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

        private void LoadSavedInputs()
        {
            var snapshot = UserInputStore.Load();
            OutputDirBox.Text = snapshot.OutputDirectory ?? string.Empty;
            BaseLedgerBox.Text = snapshot.BaseLedgerPath ?? string.Empty;
            PowerBox.Text = snapshot.PowerPath ?? string.Empty;
            RawDetailBox.Text = snapshot.RawDetailPath ?? string.Empty;
            ReferenceLedgerBox.Text = snapshot.ReferenceLedgerPath ?? string.Empty;
            CompletedLedgerBox.Text = snapshot.CompletedLedgerPath ?? string.Empty;
            ProxyTemplateDirBox.Text = snapshot.ProxyTemplateDirectory ?? string.Empty;
            IntermediaryTemplateDirBox.Text = snapshot.IntermediaryTemplateDirectory ?? string.Empty;
            SummaryTemplateBox.Text = snapshot.SummaryTemplatePath ?? string.Empty;

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
                    SummaryTemplatePath = SummaryTemplateBox.Text.Trim()
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
                || !string.IsNullOrWhiteSpace(snapshot.SummaryTemplatePath);
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
            MessageBox.Show(this, ex.Message, "出错了", MessageBoxButton.OK, MessageBoxImage.Error);
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
