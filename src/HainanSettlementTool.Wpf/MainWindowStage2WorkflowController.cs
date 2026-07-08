using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowStage2WorkflowController
    {
        private readonly Window _owner;
        private readonly Dispatcher _dispatcher;
        private readonly MainWindowInputController _inputController;
        private readonly MainWindowProgressController _progressController;
        private readonly MainWindowResultController _resultController;
        private readonly MainWindowLogController _logController;
        private readonly MainWindowDialogController _dialogController;
        private readonly Action<bool> _setBusy;
        private readonly Action _saveInputs;

        public MainWindowStage2WorkflowController(
            Window owner,
            Dispatcher dispatcher,
            MainWindowInputController inputController,
            MainWindowProgressController progressController,
            MainWindowResultController resultController,
            MainWindowLogController logController,
            MainWindowDialogController dialogController,
            Action<bool> setBusy,
            Action saveInputs)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            _inputController = inputController ?? throw new ArgumentNullException(nameof(inputController));
            _progressController = progressController ?? throw new ArgumentNullException(nameof(progressController));
            _resultController = resultController ?? throw new ArgumentNullException(nameof(resultController));
            _logController = logController ?? throw new ArgumentNullException(nameof(logController));
            _dialogController = dialogController ?? throw new ArgumentNullException(nameof(dialogController));
            _setBusy = setBusy ?? throw new ArgumentNullException(nameof(setBusy));
            _saveInputs = saveInputs ?? throw new ArgumentNullException(nameof(saveInputs));
        }

        public async Task RunAsync()
        {
            if (_inputController.SelectedProvinceOrNull() == ProvinceCode.Chongqing)
            {
                await RunChongqingStage2Async();
                return;
            }

            await RunHainanStage2Async();
        }

        private async Task RunHainanStage2Async()
        {
            HainanStage2Options options;
            try
            {
                options = _inputController.CreateHainanStage2Options();
                _saveInputs();
                if (!_dialogController.ConfirmRun("阶段二：生成分表和汇总表", options.Month, options.OutputDirectory))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                _dialogController.ShowError(ex);
                return;
            }

            _setBusy(true);
            ResetResults();
            _progressController.ResetProgress("正在执行阶段二...", "预检上月关键字段变化");
            _progressController.SetProgress(8, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始执行阶段二。", "阶段二");

            try
            {
                var workflow = SettlementWorkflowFactory.Create();
                HainanStage2WorkflowPlan plan = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() => _progressController.SetProgress(16, "读取台账并比对上月模板"));
                    plan = workflow.PlanHainanStage2(options);
                });

                var preflight = plan.Preflight;
                _progressController.SetProgress(24, preflight.HasIssues ? "预检发现需要确认的变化" : "预检完成");
                var confirmed = true;
                if (plan.RequiresConfirmation)
                {
                    _progressController.SetStatus("待确认", "WarningBrush", "StatusBusyBrush");
                    _progressController.SetStepNeedsConfirmation(0);
                    AddLog("阶段二预检发现 " + preflight.Issues.Count + " 条需要确认的变化。", "阶段二");
                    confirmed = ConfirmStage2Preflight(preflight, options);
                    if (confirmed)
                    {
                        _progressController.SetStatus("运行中", "WarningBrush", "StatusBusyBrush");
                    }
                }
                else
                {
                    AddLog("阶段二预检通过，未发现需要确认的变化。", "阶段二");
                }

                if (!confirmed)
                {
                    var cancelled = workflow.CompleteHainanStage2(plan, confirmed, LogThreadSafe);
                    if (cancelled.WasCancelled)
                    {
                        AddLog("已取消阶段二生成。", "阶段二");
                        _progressController.ResetProgress("等待执行", "已取消阶段二");
                        return;
                    }
                }

                _progressController.SetStepDone(0);
                _progressController.SetProgress(34, "读取人工整理后的台账");
                HainanStage2WorkflowResult result = null;
                await Task.Run(() =>
                {
                    result = workflow.CompleteHainanStage2(plan, confirmed, LogThreadSafe);
                });

                if (result.WasCancelled)
                {
                    AddLog("已取消阶段二生成。", "阶段二");
                    _progressController.ResetProgress("等待执行", "已取消阶段二");
                    return;
                }

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, "阶段二执行完成");
                LogSummary(result.SummaryLines);

                _resultController.SetStage2Success(result.Report.ProxyGroups + " 个文件", result.Report.IntermediaryGroups + " 个文件", "1 个文件");
                _resultController.ShowCompletion("阶段二执行完成", "分表和汇总表已生成", options.OutputDirectory);
            }
            catch (Exception ex)
            {
                _progressController.SetStepFailed();
                _progressController.SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                _dialogController.ShowError(ex);
            }
            finally
            {
                _setBusy(false);
            }
        }

        private async Task RunChongqingStage2Async()
        {
            ChongqingStage2Options options;
            try
            {
                options = _inputController.CreateChongqingStage2Options();
                _saveInputs();
                var message = new StringBuilder();
                message.AppendLine("结算月份：2026年" + options.Month + "月");
                message.AppendLine("输出文件夹：");
                message.AppendLine(options.OutputDirectory);
                message.AppendLine();
                message.AppendLine("将先读取重庆台账、代理/居间/退补模板和汇总表进行预检；确认后生成输出分表、退补表和汇总表副本。");
                if (!_dialogController.ConfirmAction("确认重庆阶段二生成", "即将执行重庆阶段二生成", message.ToString(), "开始生成"))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                _dialogController.ShowError(ex);
                return;
            }

            _setBusy(true);
            ResetResults();
            _progressController.ResetProgress("正在执行重庆阶段二...", "检查输入文件");
            _progressController.SetProgress(8, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始执行重庆阶段二。", "重庆阶段二");

            try
            {
                var workflow = SettlementWorkflowFactory.Create();
                ChongqingStage2WorkflowPlan plan = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() => _progressController.SetProgress(16, "读取重庆台账和汇总表模板"));
                    plan = workflow.PlanChongqingStage2(options);
                });

                var preflight = plan.Preflight;
                _progressController.SetProgress(60, preflight.HasIssues ? "预检发现需要确认的变化" : "预检完成");
                var confirmed = true;
                if (plan.RequiresConfirmation)
                {
                    _progressController.SetStatus("待确认", "WarningBrush", "StatusBusyBrush");
                    _progressController.SetStepNeedsConfirmation(0);
                    AddLog("重庆阶段二预检发现 " + preflight.Issues.Count + " 条需要确认的变化。", "重庆阶段二");
                    confirmed = ConfirmChongqingStage2Preflight(preflight, options);
                    if (confirmed)
                    {
                        _progressController.SetStatus("运行中", "WarningBrush", "StatusBusyBrush");
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
                        _progressController.ResetProgress("等待执行", "已取消重庆阶段二");
                        return;
                    }
                }

                _progressController.SetStepDone(0);
                _progressController.SetProgress(34, "生成重庆分表和汇总表");
                ChongqingStage2WorkflowResult result = null;
                await Task.Run(() =>
                {
                    result = workflow.CompleteChongqingStage2(plan, confirmed, LogThreadSafe);
                });

                if (result.WasCancelled)
                {
                    AddLog("已取消重庆阶段二生成。", "重庆阶段二");
                    _progressController.ResetProgress("等待执行", "已取消重庆阶段二");
                    return;
                }

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, "重庆阶段二执行完成");
                LogSummary(result.Completed.SummaryLines);

                _resultController.SetStage2Success(
                    result.Report.ProxyGroups + " 个文件",
                    "居间" + result.Report.IntermediaryGroups + "/退补" + result.Report.RefundGroups,
                    "1 个文件");
                _resultController.ShowCompletion("重庆阶段二执行完成", "分表、退补表和汇总表已生成", options.OutputDirectory);
            }
            catch (Exception ex)
            {
                _progressController.SetStepFailed();
                _progressController.SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                _dialogController.ShowError(ex);
            }
            finally
            {
                _setBusy(false);
            }
        }

        private bool ConfirmStage2Preflight(HainanStage2PreflightReport report, HainanStage2Options options)
        {
            var dialog = new HainanStage2PreflightWindow(report)
            {
                Owner = _owner
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
                Owner = _owner
            };
            var confirmed = dialog.ShowDialog() == true;
            if (confirmed)
            {
                options.SummarySubjectDecisions.Clear();
                options.SummarySubjectDecisions.AddRange(dialog.SummarySubjectDecisions);
            }

            return confirmed;
        }

        private void ResetResults()
        {
            _resultController.Reset(_inputController.SelectedProvinceOrNull());
        }

        private void LogThreadSafe(string message)
        {
            _dispatcher.Invoke(() => AddLog(message, "信息"));
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

        private void AddLog(string message, string level)
        {
            _logController.Add(message, level);
        }
    }
}
