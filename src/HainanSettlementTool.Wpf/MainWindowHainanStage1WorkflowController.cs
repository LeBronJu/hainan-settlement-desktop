using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Threading;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowHainanStage1WorkflowController
    {
        private readonly Dispatcher _dispatcher;
        private readonly MainWindowInputController _inputController;
        private readonly MainWindowProgressController _progressController;
        private readonly MainWindowResultController _resultController;
        private readonly MainWindowLogController _logController;
        private readonly MainWindowDialogController _dialogController;
        private readonly Action<bool> _setBusy;
        private readonly Action _saveInputs;

        public MainWindowHainanStage1WorkflowController(
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

        public async Task RunLedgerUpdateAsync()
        {
            HainanStage1Options options;
            try
            {
                options = _inputController.CreateHainanStage1Options();
                _saveInputs();
                if (!_dialogController.ConfirmRun("阶段一：写入电量到台账", options.Month, options.OutputDirectory))
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
            _progressController.ResetProgress("正在执行阶段一...", "写入电量到台账");
            _progressController.SetProgress(10, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始执行阶段一。", "阶段一");

            try
            {
                StageWorkflowResult<HainanStage1Report> result = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() =>
                    {
                        _progressController.SetStepDone(0);
                        _progressController.SetStepRunning(1);
                        _progressController.SetProgress(28, "读取台账和电量文件");
                    });

                    result = SettlementWorkflowFactory.Create().RunHainanStage1(options, LogThreadSafe);
                });

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, "阶段一执行完成");
                LogSummary(result.SummaryLines);

                _resultController.SetStage1Success("1 个文件");
                _resultController.ShowCompletion("阶段一执行完成", "台账和检查报告已生成", options.OutputDirectory);
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

        public async Task CleanPowerAsync()
        {
            HainanPowerCleanInput input;
            try
            {
                input = _inputController.PrepareHainanPowerCleanInput();
                _saveInputs();

                var message = "即将清洗原始零售侧明细并生成电量处理表。\n\n输出文件：\n" + input.OutputPath;
                if (!_dialogController.ConfirmAction("确认清洗电量", "即将清洗电量数据", message, "开始清洗"))
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
            _progressController.ResetProgress("正在清洗电量...", "生成电量处理表");
            _progressController.SetProgress(10, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始清洗电量数据。", "阶段一");

            try
            {
                StageWorkflowResult<HainanPowerCleanReport> result = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() =>
                    {
                        _progressController.SetStepDone(0);
                        _progressController.SetStepRunning(1);
                        _progressController.SetProgress(35, "读取原始零售侧明细");
                    });

                    result = SettlementWorkflowFactory.Create().CleanHainanPowerData(input.RawDetailPath, input.OutputPath, LogThreadSafe);
                });

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, "电量清洗完成");
                LogSummary(result.SummaryLines);

                _resultController.SetStage1Success(result.Report.PowerRows + " 个客户");
                _resultController.ShowCompletion("电量清洗完成", "电量处理表已生成", input.OutputDirectory);
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
