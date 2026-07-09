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
    internal sealed class MainWindowChongqingStage1WorkflowController
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

        public MainWindowChongqingStage1WorkflowController(
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

        public async Task CleanPowerAsync()
        {
            ProvinceStage1CleanOptions options;
            try
            {
                options = _inputController.CreateProvinceStage1CleanOptions();
                _saveInputs();

                var message = "即将清洗重庆交易中心电量确认结算单。\n\n输出内容：用户电量汇总、户号明细、JSON校验报告\n输出文件夹：\n" + options.OutputDirectory;
                if (!_dialogController.ConfirmAction("确认清洗重庆电量", "即将清洗重庆电量数据", message, "开始清洗"))
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
            _progressController.ResetProgress("正在清洗重庆电量...", "生成阶段一电量处理表");
            _progressController.SetProgress(10, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始清洗重庆阶段一电量数据。", "重庆");

            try
            {
                StageWorkflowResult<ProvinceStage1CleanResult> result = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() =>
                    {
                        _progressController.SetStepDone(0);
                        _progressController.SetStepRunning(1);
                        _progressController.SetProgress(35, "读取交易中心电量确认结算单");
                    });

                    result = SettlementWorkflowFactory.Create().CleanProvinceStage1PowerData(options, LogThreadSafe);
                });

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, "重庆电量清洗完成");
                LogSummary(result.SummaryLines);

                _resultController.SetStage1Success(result.Report.CustomerRows + " 个客户");
                _resultController.ShowCompletion("重庆电量清洗完成", "电量处理表、户号明细和校验报告已生成", options.OutputDirectory);
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

        public async Task RunLedgerUpdateAsync()
        {
            ProvinceStage1LedgerUpdateOptions options;
            try
            {
                options = _inputController.CreateProvinceStage1LedgerUpdateOptions();
                _saveInputs();
            }
            catch (Exception ex)
            {
                _dialogController.ShowError(ex);
                return;
            }

            _setBusy(true);
            ResetResults();
            _progressController.ResetProgress("正在预检重庆台账...", "读取台账和电量明细");
            _progressController.SetProgress(10, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始预检重庆阶段一台账更新。", "重庆");

            try
            {
                var workflow = SettlementWorkflowFactory.Create();
                ProvinceStage1LedgerUpdatePlan plan = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() => _progressController.SetProgress(22, "读取重庆台账和交易中心电量确认单"));
                    plan = workflow.PlanProvinceStage1LedgerUpdate(options, LogThreadSafe);
                });

                _progressController.SetProgress(30, plan.RequiresConfirmation ? "预检发现需要确认的项目" : "预检完成");
                if (!ConfirmLedgerUpdate(options, plan))
                {
                    AddLog("已取消重庆阶段一台账更新。", "重庆");
                    _progressController.ResetProgress("等待执行", "已取消重庆台账更新");
                    return;
                }

                _progressController.SetStepDone(0);
                _progressController.SetProgress(40, "写入重庆台账副本");
                StageWorkflowResult<ProvinceStage1LedgerUpdateResult> result = null;
                await Task.Run(() =>
                {
                    result = workflow.UpdateProvinceStage1Ledger(options, LogThreadSafe);
                });

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, "重庆台账更新完成");
                LogSummary(result.SummaryLines);

                _resultController.SetStage1Success(result.Report.UpdatedPowerRows + " 个客户");
                _resultController.ShowCompletion("重庆台账更新完成", "台账副本和更新报告已生成", options.OutputDirectory);
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

        private bool ConfirmLedgerUpdate(ProvinceStage1LedgerUpdateOptions options, ProvinceStage1LedgerUpdatePlan plan)
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
                    Owner = _owner
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
            return _dialogController.ConfirmAction("确认重庆台账更新", "即将写入重庆台账副本", message.ToString(), "开始写入");
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
