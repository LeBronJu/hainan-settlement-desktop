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
    internal sealed class MainWindowProvinceStage1WorkflowController
    {
        private static readonly string[] CleanPowerStepNames =
        {
            "检查输入文件",
            "读取本月电量",
            "合并客户电量",
            "生成电量表",
            "生成检查报告"
        };

        private static readonly string[] LedgerUpdateStepNames =
        {
            "检查输入文件",
            "读取台账和电量",
            "整理本月台账",
            "生成电量表",
            "生成检查报告"
        };

        private readonly Window _owner;
        private readonly Dispatcher _dispatcher;
        private readonly MainWindowInputController _inputController;
        private readonly MainWindowProgressController _progressController;
        private readonly MainWindowResultController _resultController;
        private readonly MainWindowLogController _logController;
        private readonly MainWindowDialogController _dialogController;
        private readonly Action<bool> _setBusy;
        private readonly Action _saveInputs;

        public MainWindowProvinceStage1WorkflowController(
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
            string provinceName;
            try
            {
                options = _inputController.CreateProvinceStage1CleanOptions();
                provinceName = ProvinceDisplayNames.GetName(options.Province);
                _saveInputs();

                var outputDescription = options.Province == ProvinceCode.Guangdong
                    ? "八列电量表和可直接打开的检查报告"
                    : "用户电量表、户号明细和检查报告";
                var message = "即将整理"
                    + provinceName
                    + "本月电量。\n\n将生成："
                    + outputDescription
                    + "\n输出文件夹：\n"
                    + options.OutputDirectory;
                if (!_dialogController.ConfirmAction(
                    "确认整理" + provinceName + "电量",
                    "即将整理" + provinceName + "本月电量",
                    message,
                    "开始整理"))
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
            _progressController.ResetProgress(
                "正在整理" + provinceName + "电量...",
                "生成本月电量表",
                CleanPowerStepNames);
            _progressController.SetProgress(10, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始整理" + provinceName + "本月电量。", provinceName);

            try
            {
                StageWorkflowResult<ProvinceStage1CleanResult> result = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() =>
                    {
                        _progressController.SetStepDone(0);
                        _progressController.SetStepRunning(1);
                        _progressController.SetProgress(35, "读取" + provinceName + "本月电量明细");
                    });

                    result = SettlementWorkflowFactory.Create().CleanProvinceStage1PowerData(options, LogThreadSafe);
                });

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, provinceName + "本月电量整理完成");
                LogSummary(result.SummaryLines);

                _resultController.SetStage1Success(result.Report.CustomerRows + " 个客户");
                _resultController.ShowCompletion(
                    provinceName + "本月电量已整理",
                    "电量表和检查报告已生成",
                    options.OutputDirectory,
                    result.Report.HtmlReportPath);
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
            string provinceName;
            try
            {
                options = _inputController.CreateProvinceStage1LedgerUpdateOptions();
                provinceName = ProvinceDisplayNames.GetName(options.Province);
                _saveInputs();
            }
            catch (Exception ex)
            {
                _dialogController.ShowError(ex);
                return;
            }

            _setBusy(true);
            ResetResults();
            _progressController.ResetProgress(
                "正在检查" + provinceName + "本月资料...",
                "读取台账和电量明细",
                LedgerUpdateStepNames);
            _progressController.SetProgress(10, "检查输入文件");
            _progressController.SetStepRunning(0);
            AddLog("开始检查" + provinceName + "本月台账资料。", provinceName);

            try
            {
                var workflow = SettlementWorkflowFactory.Create();
                ProvinceStage1LedgerUpdatePlan plan = null;
                await Task.Run(() =>
                {
                    _dispatcher.Invoke(() => _progressController.SetProgress(22, "读取" + provinceName + "台账和电量来源"));
                    plan = workflow.PlanProvinceStage1LedgerUpdate(options, LogThreadSafe);
                });

                _progressController.SetProgress(30, plan.RequiresConfirmation ? "发现需要查看的事项" : "检查完成");
                if (!ConfirmLedgerUpdate(options, plan))
                {
                    AddLog("已取消" + provinceName + "阶段一台账更新。", provinceName);
                    _progressController.ResetProgress("等待执行", "已取消" + provinceName + "台账更新");
                    return;
                }

                _progressController.SetStepDone(0);
                _progressController.SetProgress(40, "生成" + provinceName + "本月台账");
                StageWorkflowResult<ProvinceStage1LedgerUpdateResult> result = null;
                await Task.Run(() =>
                {
                    result = workflow.UpdateProvinceStage1Ledger(options, LogThreadSafe);
                });

                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(100, provinceName + "本月台账已生成");
                LogSummary(result.SummaryLines);
                if (!string.IsNullOrWhiteSpace(result.Report.HtmlReportPath))
                {
                    AddLog("检查报告：" + result.Report.HtmlReportPath, "信息");
                }

                _resultController.SetStage1Success(result.Report.UpdatedPowerRows + " 个客户");
                var hasFocusItems = ProvinceStage1ReviewGuide.Build(result.Report.Issues)
                    .Exists(group => group.NeedsAttention);
                if (hasFocusItems)
                {
                    _resultController.ShowReviewCompletion(
                        provinceName + "本月台账已生成",
                        "请打开检查报告，按“下一步”和“重点检查”完成复核",
                        options.OutputDirectory,
                        false,
                        result.Report.HtmlReportPath);
                }
                else
                {
                    _resultController.ShowCompletion(
                        provinceName + "本月台账已生成",
                        "电量表、台账和检查报告已生成",
                        options.OutputDirectory,
                        result.Report.HtmlReportPath);
                }
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
            var provinceName = ProvinceDisplayNames.GetName(plan.Province);
            var message = new StringBuilder();
            message.AppendLine("结算月份：2026年" + options.Month + "月");
            message.AppendLine("匹配客户：" + plan.MatchedRows + " / " + plan.PowerCustomerRows);
            if (plan.Province == ProvinceCode.Guangdong)
            {
                message.AppendLine("多计量点客户：" + plan.MultiAccountRows + " 个（同编码电量已相加）");
                message.AppendLine("新增客户：" + plan.CreatedCustomerRows + " 个（自动追加安全字段，其余资料待人工补齐）");
            }
            else
            {
                message.AppendLine("多户号客户：" + plan.MultiAccountRows + " 行（仅提示，不写入B列）");
            }
            message.AppendLine("输出文件夹：");
            message.AppendLine(options.OutputDirectory);

            if (plan.RequiresConfirmation || plan.Province == ProvinceCode.Guangdong)
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
            message.AppendLine("没有需要额外处理的客户。确认后会生成新的台账文件，不会改原台账。");
            return _dialogController.ConfirmAction(
                "确认生成" + provinceName + "本月台账",
                "即将生成" + provinceName + "本月台账",
                message.ToString(),
                "开始生成");
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
