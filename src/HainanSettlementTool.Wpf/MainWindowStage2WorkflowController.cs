using System;
using System.Collections.Generic;
using System.Linq;
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
            var province = _inputController.SelectedProvinceOrNull();
            if (province == ProvinceCode.Chongqing)
            {
                await RunChongqingStage2Async();
                return;
            }

            if (province == ProvinceCode.Guangdong)
            {
                await RunGuangdongStage2MonthPreparationAsync();
                return;
            }

            if (province != ProvinceCode.Hainan)
            {
                _dialogController.ShowErrorMessage("请选择结算省份。");
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
                    confirmed = ConfirmHainanStage2Preflight(preflight, plan.Evaluation, options);
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
                    AddLog("已取消阶段二生成。", "阶段二");
                    _progressController.ResetProgress("等待执行", "已取消阶段二");
                    return;
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
                var needsReview = HainanStage2NeedsReview(result.Report);
                _progressController.SetProgress(100, needsReview ? "阶段二完成但需要复核" : "阶段二执行完成");
                LogSummary(result.SummaryLines);
                AddLog("海南阶段二可读报告：" + result.Report.HtmlReportPath, "信息");

                var status = needsReview ? "需复核" : "成功";
                _resultController.SetStage2Outcome(
                    status,
                    result.Report.ProxyGroups + " 个文件",
                    status,
                    result.Report.IntermediaryGroups + " 个文件",
                    status,
                    "1 个文件");
                if (needsReview)
                {
                    _progressController.SetStatus("需复核", "WarningBrush", "StatusBusyBrush");
                    _resultController.ShowReviewCompletion(
                        "阶段二生成完成但需要复核",
                        "正式分表和汇总表已完整生成；付款前请查看可读报告。",
                        options.OutputDirectory,
                        false,
                        result.Report.HtmlReportPath);
                    ShowHainanStage2ReviewReminder(result.Report);
                }
                else
                {
                    _resultController.ShowCompletion(
                        "阶段二执行完成",
                        "分表和汇总表已生成",
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

        private async Task RunChongqingStage2Async()
        {
            ChongqingStage2Options options;
            try
            {
                options = _inputController.CreateChongqingStage2Options();
                _saveInputs();
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
                    confirmed = ConfirmChongqingStage2Preflight(preflight, plan.Evaluation, options);
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
                    AddLog("已取消重庆阶段二生成。", "重庆阶段二");
                    _progressController.ResetProgress("等待执行", "已取消重庆阶段二");
                    return;
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
                var needsReview = ChongqingStage2NeedsReview(result.Report);
                _progressController.SetProgress(100, needsReview ? "重庆阶段二完成但需要复核" : "重庆阶段二执行完成");
                LogSummary(result.Completed.SummaryLines);
                AddLog("重庆阶段二可读报告：" + result.Report.HtmlReportPath, "信息");

                var status = needsReview ? "需复核" : "成功";
                _resultController.SetStage2Outcome(
                    status,
                    result.Report.ProxyGroups + " 个文件",
                    status,
                    "居间" + result.Report.IntermediaryGroups + "/退补" + result.Report.RefundGroups,
                    status,
                    "1 个文件");
                if (needsReview)
                {
                    _progressController.SetStatus("需复核", "WarningBrush", "StatusBusyBrush");
                    _resultController.ShowReviewCompletion(
                        "重庆阶段二生成完成但需要复核",
                        "正式分表和汇总表已完整生成；付款前请查看可读报告。",
                        options.OutputDirectory,
                        false,
                        result.Report.HtmlReportPath);
                    ShowChongqingStage2ReviewReminder(result.Report);
                }
                else
                {
                    _resultController.ShowCompletion(
                        "重庆阶段二执行完成",
                        "分表、退补表和汇总表已生成",
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

        private async Task RunGuangdongStage2MonthPreparationAsync()
        {
            GuangdongStage2MonthPreparationOptions options;
            try
            {
                options = _inputController.CreateGuangdongStage2MonthPreparationOptions();
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
                "正在扫描广东分表...",
                "识别标准数字月份 sheet",
                new[]
                {
                    "扫描分表文件",
                    "识别标准月份 sheet",
                    "创建或整理目标月",
                    "写入检查报告",
                    "保存输出副本"
                });
            _progressController.SetProgress(8, "扫描代理、居间和退补目录");
            _progressController.SetStepRunning(0);
            AddLog("开始扫描广东分表月份初始化条件。", "广东分表");
            var requiresReviewAfterRun = false;
            var criticalFailureAfterRun = false;

            try
            {
                var service = SettlementWorkflowFactory.CreateGuangdongStage2MonthPreparationService();
                GuangdongStage2PreflightReport preflight = null;
                await Task.Run(() =>
                {
                    preflight = service.Analyze(options);
                });

                _progressController.SetProgress(28, "广东分表扫描完成");
                if (preflight.Workbooks.Count == 0)
                {
                    _progressController.ResetProgress("等待执行", "没有找到可检查的 .xlsx 文件");
                    _dialogController.ShowWarningMessage(
                        "没有找到广东分表",
                        "所选文件夹中没有可处理的 .xlsx 文件",
                        "程序会递归扫描所选代理、居间和退补文件夹，并忽略 ~$ 或 ._ 开头的临时文件。");
                    return;
                }

                _progressController.SetStatus("待确认", "WarningBrush", "StatusBusyBrush");
                _progressController.SetStepNeedsConfirmation(0);
                if (!_dialogController.ConfirmAction(
                    "确认广东分表月份初始化",
                    "已完成只读预检",
                    BuildGuangdongPreflightMessage(preflight),
                    GuangdongConfirmationButtonText(preflight)))
                {
                    AddLog("已取消广东分表月份初始化。", "广东分表");
                    _progressController.ResetProgress("等待执行", "已取消广东分表月份初始化");
                    return;
                }

                _progressController.SetStatus("运行中", "WarningBrush", "StatusBusyBrush");
                _progressController.SetStepDone(0);
                _progressController.SetProgress(42, "创建或整理目标月份 sheet");
                GuangdongStage2MonthPreparationReport report = null;
                await Task.Run(() =>
                {
                    report = service.Run(options, LogThreadSafe);
                });

                requiresReviewAfterRun = report.HasReviewItems;
                criticalFailureAfterRun = report.HasCriticalFailures;
                _progressController.SetStepDone(1);
                _progressController.SetStepDone(2);
                _progressController.SetStepDone(3);
                _progressController.SetStepDone(4);
                _progressController.SetProgress(
                    100,
                    criticalFailureAfterRun
                        ? "广东分表月份初始化存在异常"
                        : requiresReviewAfterRun
                            ? "广东分表月份初始化部分完成"
                            : "广东分表月份初始化完成");
                AddLog(
                    "广东分表正常输出 " + report.SuccessfulCount + " / " + report.InputCount
                    + " 个；需人工复核 " + report.SkippedCount + " 个；失败 " + report.FailedCount + " 个。",
                    criticalFailureAfterRun ? "错误" : requiresReviewAfterRun ? "警告" : "成功");
                AddLog("广东分表校验报告：" + report.ValidationReportPath, "信息");
                AddLog("广东分表可读报告：" + report.HtmlReportPath, "信息");

                _resultController.SetStage2Outcome(
                    GuangdongKindStatus(report, GuangdongStage2SettlementKinds.Proxy),
                    GuangdongKindCount(report, GuangdongStage2SettlementKinds.Proxy),
                    GuangdongKindStatus(report, GuangdongStage2SettlementKinds.Intermediary),
                    GuangdongKindCount(report, GuangdongStage2SettlementKinds.Intermediary),
                    GuangdongKindStatus(report, GuangdongStage2SettlementKinds.Refund),
                    GuangdongKindCount(report, GuangdongStage2SettlementKinds.Refund));
                if (!report.HasReviewItems)
                {
                    _resultController.ShowCompletion(
                        "广东分表初始化完成",
                        "已生成目标月份分表副本和检查报告",
                        report.OutputDirectory,
                        report.HtmlReportPath);
                }
                else
                {
                    _resultController.ShowReviewCompletion(
                        criticalFailureAfterRun ? "广东分表初始化存在异常" : "广东分表初始化部分完成",
                        "正常输出 " + report.SuccessfulCount + " / " + report.InputCount
                        + "；需人工复核 " + report.SkippedCount + "；失败 " + report.FailedCount + "。",
                        report.OutputDirectory,
                        criticalFailureAfterRun,
                        report.HtmlReportPath);
                    ShowGuangdongReviewReminder(report);
                }
            }
            catch (Exception ex)
            {
                criticalFailureAfterRun = true;
                _progressController.SetStepFailed();
                _progressController.SetProgress(100, "执行失败");
                AddLog(ex.Message, "错误");
                _resultController.ShowReviewCompletion(
                    "广东分表初始化执行失败",
                    "本次运行未完成，请根据错误提示检查后重试。",
                    options.OutputDirectory,
                    true);
                _dialogController.ShowError(ex);
            }
            finally
            {
                _setBusy(false);
                if (criticalFailureAfterRun)
                {
                    _progressController.SetStatus("执行异常", "ErrorBrush", "StatusBusyBrush");
                }
                else if (requiresReviewAfterRun)
                {
                    _progressController.SetStatus("需复核", "WarningBrush", "StatusBusyBrush");
                }
            }
        }

        private static string GuangdongConfirmationButtonText(GuangdongStage2PreflightReport report)
        {
            if (report.SkippedCount == 0)
            {
                return "开始初始化";
            }

            return report.ProcessableCount > 0
                ? "继续处理其余 " + report.ProcessableCount + " 个"
                : "保存复核文件";
        }

        private static string BuildGuangdongPreflightMessage(GuangdongStage2PreflightReport report)
        {
            var message = new StringBuilder();
            message.AppendLine("结算月份：" + report.Year + "年" + report.Month + "月");
            message.AppendLine("扫描文件：" + report.Workbooks.Count + " 个");
            message.AppendLine();
            message.AppendLine("- 从标准上月 sheet 创建：" + report.CreateCount + " 个");
            message.AppendLine("- 保留现有目标月并整理：" + report.NormalizeCount + " 个");
            message.AppendLine("- 现有目标月已经准备完成：" + report.AlreadyPreparedCount + " 个");
            message.AppendLine("- 结构、日期或来源异常，将跳过：" + report.SkippedCount + " 个");
            if (report.SkippedCount > 0)
            {
                message.AppendLine();
                message.AppendLine("以下文件不会新增目标月份 sheet：");
                foreach (var workbook in report.Workbooks
                    .Where(item => item.Action == GuangdongStage2PreparationActions.Skipped)
                    .Take(3))
                {
                    message.AppendLine("- [" + workbook.SettlementKind + "] " + workbook.RelativePath);
                    message.AppendLine("  原因：" + workbook.Message);
                }

                if (report.SkippedCount > 3)
                {
                    message.AppendLine("- 其余 " + (report.SkippedCount - 3) + " 个请在运行报告中查看。");
                }

                message.AppendLine();
                message.AppendLine("继续后，这些原文件会原样保留到对应类别的“【未处理-需人工复核】”目录。"
                    + "程序不会自动修正异常日期。");
            }

            message.AppendLine();
            message.AppendLine("只识别名称完全等于数字月份的标准 sheet；其它 sheet 原样保留但不参与来源选择。所有结果写入新的输出目录，不修改原文件。");
            return message.ToString();
        }

        private void ShowGuangdongReviewReminder(GuangdongStage2MonthPreparationReport report)
        {
            if (!report.HasReviewItems)
            {
                return;
            }

            var message = new StringBuilder();
            message.AppendLine("扫描输入：" + report.InputCount + " 个；正常输出：" + report.SuccessfulCount
                + " 个；需人工复核：" + report.SkippedCount + " 个；失败：" + report.FailedCount + " 个。");
            message.AppendLine();
            message.AppendLine("需要关注的文件：");
            foreach (var workbook in report.Workbooks
                .Where(item => item.Action == GuangdongStage2PreparationActions.Skipped
                    || item.Action == GuangdongStage2PreparationActions.Failed)
                .Take(3))
            {
                message.AppendLine("- [" + workbook.SettlementKind + "] " + workbook.RelativePath);
                message.AppendLine("  原因：" + workbook.Message);
            }

            var reviewCount = report.SkippedCount + report.FailedCount;
            if (reviewCount > 3)
            {
                message.AppendLine("- 其余 " + (reviewCount - 3) + " 个请查看报告。");
            }

            if (report.PreservedReviewCopyCount > 0)
            {
                message.AppendLine();
                message.AppendLine("未自动处理的原文件已原样保留在对应类别的“【未处理-需人工复核】”目录，未新增目标月份 sheet。");
            }

            message.AppendLine();
            message.AppendLine("优先查看可读报告：");
            message.AppendLine(report.HtmlReportPath);
            _dialogController.ShowWarningMessage(
                report.HasCriticalFailures ? "广东分表执行异常" : "广东分表需要复核",
                report.HasCriticalFailures ? "部分输入未形成可用输出" : "本批次尚未完整处理",
                message.ToString());
        }

        private static string GuangdongKindStatus(
            GuangdongStage2MonthPreparationReport report,
            string settlementKind)
        {
            if (report.FailedCountFor(settlementKind) > 0)
            {
                return "失败";
            }

            if (!report.IsClassificationComplete || !report.HasCompleteOutputSet)
            {
                return "需复核";
            }

            return report.SkippedCountFor(settlementKind) > 0 ? "需复核" : "成功";
        }

        private static string GuangdongKindCount(
            GuangdongStage2MonthPreparationReport report,
            string settlementKind)
        {
            return report.CountFor(settlementKind) + " / " + report.InputCountFor(settlementKind);
        }

        private bool ConfirmHainanStage2Preflight(
            HainanStage2PreflightReport report,
            Stage2PreflightEvaluation evaluation,
            HainanStage2Options options)
        {
            var dialog = new Stage2PreflightWindow(
                Stage2PreflightPresentationAdapter.CreateHainan(report, evaluation))
            {
                Owner = _owner
            };
            var confirmed = dialog.ShowDialog() == true;
            if (confirmed)
            {
                options.SummarySubjectDecisions.Clear();
                options.SummarySubjectDecisions.AddRange(dialog.PaymentDecisions.Select(decision =>
                    new HainanStage2SummarySubjectDecision
                    {
                        SettlementKind = decision.SettlementKind,
                        Entity = decision.Entity,
                        PaymentParty = decision.PaymentParty
                    }));
                options.TemplateDecisions.Clear();
                options.TemplateDecisions.AddRange(dialog.TemplateDecisions.Select(decision =>
                    new HainanStage2TemplateDecision
                    {
                        SettlementKind = decision.SettlementKind,
                        Entity = decision.Entity,
                        TemplatePath = decision.TemplatePath
                    }));
            }

            return confirmed;
        }

        private void ShowHainanStage2ReviewReminder(HainanStage2Report report)
        {
            if (report == null
                || (report.Warnings.Count == 0
                    && report.AuditIssues.Count == 0
                    && report.MissingOwners.Count == 0))
            {
                return;
            }

            var message = new StringBuilder();
            message.AppendLine("生成已完成，但仍有项目需要人工复核。");
            message.AppendLine("校验项目：" + report.AuditIssues.Count + " 条；新增/自动填充提示："
                + report.Warnings.Count + " 条。");

            foreach (var group in report.AuditIssues
                .Where(item => item != null)
                .GroupBy(item => string.IsNullOrWhiteSpace(item.Category) ? "其他校验项目" : item.Category)
                .Take(3))
            {
                message.AppendLine("- " + group.Key + "：" + group.Count() + " 项");
            }

            foreach (var warning in report.Warnings.Where(item => !string.IsNullOrWhiteSpace(item)).Take(2))
            {
                message.AppendLine("- " + ShortenStage2ReviewWarning(warning));
            }

            if (report.MissingOwners.Count > 0)
            {
                message.AppendLine("- 负责人缺失：" + report.MissingOwners.Count + " 项");
            }

            message.AppendLine();
            message.AppendLine("优先查看可读报告：");
            message.AppendLine(report.HtmlReportPath);
            _dialogController.ShowWarningMessage(
                "海南阶段二需要复核",
                "生成完成，但有项目需要人工确认",
                message.ToString());
        }

        private static bool HainanStage2NeedsReview(HainanStage2Report report)
        {
            return report != null
                && (report.Warnings.Any(item => !string.IsNullOrWhiteSpace(item))
                    || report.MissingOwners.Count > 0
                    || report.AuditIssues.Any(issue =>
                        issue != null && issue.Disposition == Stage2PreflightDisposition.Review));
        }

        private void ShowChongqingStage2ReviewReminder(ChongqingStage2Report report)
        {
            if (report == null || (report.Warnings.Count == 0 && report.AuditIssues.Count == 0))
            {
                return;
            }

            _dialogController.ShowWarningMessage(
                "重庆阶段二需要复核",
                "生成完成，但有项目需要人工确认",
                BuildChongqingStage2ReviewMessage(report));
        }

        private static bool ChongqingStage2NeedsReview(ChongqingStage2Report report)
        {
            return report != null
                && (report.Warnings.Any(item => !string.IsNullOrWhiteSpace(item))
                    || report.AuditIssues.Any(issue =>
                        issue != null && issue.Disposition == Stage2PreflightDisposition.Review));
        }

        private static string BuildChongqingStage2ReviewMessage(ChongqingStage2Report report)
        {
            var message = new StringBuilder();
            message.AppendLine("校验问题：" + report.AuditIssues.Count + " 条；自动生成提示：" + report.Warnings.Count + " 条");
            message.AppendLine();

            var extraDeductionWarnings = report.Warnings.Count(item => item != null && item.IndexOf("额外扣减块", StringComparison.Ordinal) >= 0);
            if (extraDeductionWarnings > 0)
            {
                message.AppendLine("- 退补额外扣减块：" + extraDeductionWarnings + " 项。已同步 C-G 当月电量；H 列以后、汇总表当月抵扣和实际支付仍按模板保留，请人工复核。");
            }

            foreach (var group in report.AuditIssues
                .Where(item => item != null)
                .GroupBy(item => string.IsNullOrWhiteSpace(item.Category) ? "其他校验问题" : item.Category)
                .Take(3))
            {
                message.AppendLine("- " + group.Key + "：" + group.Count() + " 项");
            }

            var otherWarnings = report.Warnings
                .Where(item => item == null || item.IndexOf("额外扣减块", StringComparison.Ordinal) < 0)
                .Take(2)
                .Select(ShortenStage2ReviewWarning)
                .ToList();
            foreach (var warning in otherWarnings)
            {
                message.AppendLine("- " + warning);
            }

            if (report.AuditIssues.Count > 3 || report.Warnings.Count > extraDeductionWarnings + otherWarnings.Count)
            {
                message.AppendLine("- 其余复核项请查看校验报告。");
            }

            message.AppendLine();
            message.AppendLine("优先查看可读报告：");
            message.AppendLine(report.HtmlReportPath);
            return message.ToString();
        }

        private static string ShortenStage2ReviewWarning(string warning)
        {
            if (string.IsNullOrWhiteSpace(warning))
            {
                return "自动生成提示";
            }

            var fileIndex = warning.IndexOf("文件：", StringComparison.Ordinal);
            var text = fileIndex >= 0 ? warning.Substring(0, fileIndex).TrimEnd(' ', '；', ';') : warning.Trim();
            return text.Length <= 96 ? text : text.Substring(0, 96) + "...";
        }

        private bool ConfirmChongqingStage2Preflight(
            ChongqingStage2PreflightReport report,
            Stage2PreflightEvaluation evaluation,
            ChongqingStage2Options options)
        {
            var dialog = new Stage2PreflightWindow(
                Stage2PreflightPresentationAdapter.CreateChongqing(report, evaluation))
            {
                Owner = _owner
            };
            var confirmed = dialog.ShowDialog() == true;
            if (confirmed)
            {
                options.SummarySubjectDecisions.Clear();
                options.SummarySubjectDecisions.AddRange(dialog.PaymentDecisions.Select(decision =>
                    new ChongqingStage2SummarySubjectDecision
                    {
                        SettlementKind = decision.SettlementKind,
                        Entity = decision.Entity,
                        PaymentParty = decision.PaymentParty
                    }));
                options.TemplateDecisions.Clear();
                options.TemplateDecisions.AddRange(dialog.TemplateDecisions.Select(decision =>
                    new ChongqingStage2TemplateDecision
                    {
                        SettlementKind = decision.SettlementKind,
                        Entity = decision.Entity,
                        TemplatePath = decision.TemplatePath
                    }));
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
