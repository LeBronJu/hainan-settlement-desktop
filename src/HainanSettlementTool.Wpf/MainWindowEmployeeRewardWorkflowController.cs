using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Threading;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowEmployeeRewardWorkflowController
    {
        private readonly Dispatcher _dispatcher;
        private readonly MainWindowInputController _inputController;
        private readonly MainWindowProgressController _progressController;
        private readonly MainWindowResultController _resultController;
        private readonly MainWindowLogController _logController;
        private readonly MainWindowDialogController _dialogController;
        private readonly Action<bool> _setBusy;
        private readonly Action _saveInputs;

        public MainWindowEmployeeRewardWorkflowController(
            Dispatcher dispatcher,
            MainWindowInputController inputController,
            MainWindowProgressController progressController,
            MainWindowResultController resultController,
            MainWindowLogController logController,
            MainWindowDialogController dialogController,
            Action<bool> setBusy,
            Action saveInputs)
        {
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
            EmployeeRewardOptions options;
            try
            {
                options = _inputController.CreateEmployeeRewardOptions();
                _saveInputs();
                if (!ConfirmRun(options))
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
            _progressController.ResetProgress("正在生成员工电量奖励...", "检查输入文件");
            _progressController.SetProgress(10, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始生成员工电量奖励表。", "员工奖励");

            try
            {
                StageWorkflowResult<EmployeeRewardResult> result = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() =>
                    {
                        _progressController.SetStepDone(0);
                        _progressController.SetStepRunning(1);
                        _progressController.SetProgress(30, "读取售电结算台账");
                    });

                    result = SettlementWorkflowFactory.Create().RunEmployeeReward(options, LogThreadSafe);
                });

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, "员工电量奖励生成完成");
                LogSummary(result.SummaryLines);

                _resultController.SetEmployeeRewardSuccess(result.Report.PersonalWorkbookPaths.Count + " 个", "1 个文件");
                _resultController.ShowCompletion("员工电量奖励生成完成", "奖励总表、个人确认表和校验报告已生成", options.OutputDirectory);
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

        private bool ConfirmRun(EmployeeRewardOptions options)
        {
            var period = options.StartMonth == options.EndMonth
                ? "2026年" + options.StartMonth + "月"
                : "2026年" + options.StartMonth + "-" + options.EndMonth + "月";
            var message = "即将生成员工电量奖励表。\n\n期间：" + period + "\n输出文件夹：\n" + options.OutputDirectory;
            return _dialogController.ConfirmAction("确认生成员工电量奖励", "即将生成员工电量奖励", message, "开始生成");
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
