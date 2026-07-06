using System;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class SettlementWorkflow
    {
        private readonly Stage1Service _stage1Service;
        private readonly Stage2Service _stage2Service;
        private readonly EmployeeRewardService _employeeRewardService;
        private readonly ProvinceStage1Service _provinceStage1Service;

        public SettlementWorkflow(Stage1Service stage1Service, Stage2Service stage2Service)
            : this(stage1Service, stage2Service, null)
        {
        }

        public SettlementWorkflow(Stage1Service stage1Service, Stage2Service stage2Service, EmployeeRewardService employeeRewardService)
            : this(stage1Service, stage2Service, employeeRewardService, null)
        {
        }

        public SettlementWorkflow(
            Stage1Service stage1Service,
            Stage2Service stage2Service,
            EmployeeRewardService employeeRewardService,
            ProvinceStage1Service provinceStage1Service)
        {
            if (stage1Service == null)
            {
                throw new ArgumentNullException(nameof(stage1Service));
            }

            if (stage2Service == null)
            {
                throw new ArgumentNullException(nameof(stage2Service));
            }

            _stage1Service = stage1Service;
            _stage2Service = stage2Service;
            _employeeRewardService = employeeRewardService;
            _provinceStage1Service = provinceStage1Service;
        }

        public StageWorkflowResult<Stage1Report> RunStage1(Stage1Options options, Action<string> log)
        {
            var report = _stage1Service.Run(options, log);
            return new StageWorkflowResult<Stage1Report>(
                report,
                new[]
                {
                    "阶段1完成。",
                    "输出台账：" + report.Output,
                    "报告：" + report.ReportPath
                });
        }

        public StageWorkflowResult<PowerCleanReport> CleanPowerData(string rawDetailPath, string outputPath, Action<string> log)
        {
            var report = _stage1Service.CleanPowerData(rawDetailPath, outputPath, log);
            return new StageWorkflowResult<PowerCleanReport>(
                report,
                new[]
                {
                    "电量清洗完成。",
                    "电量处理表：" + report.OutputPath,
                    "客户数量：" + report.PowerRows + "，合计电量：" + report.MonthTotal.ToString("0.####")
                });
        }

        public StageWorkflowResult<ProvinceStage1CleanResult> CleanProvinceStage1PowerData(
            ProvinceStage1CleanOptions options,
            Action<string> log)
        {
            if (_provinceStage1Service == null)
            {
                throw new InvalidOperationException("多省份阶段一服务未配置。");
            }

            var report = _provinceStage1Service.CleanPowerData(options, log);
            return new StageWorkflowResult<ProvinceStage1CleanResult>(
                report,
                new[]
                {
                    ProvinceStage1Service.ProvinceName(report.Province) + "阶段一电量清洗完成。",
                    "电量处理表：" + report.OutputWorkbookPath,
                    "报告：" + report.ReportPath,
                    "客户数量：" + report.CustomerRows + "，户号数量：" + report.AccountRows,
                    "合计电量：" + report.TotalPower.ToString("0.####") + " " + report.Unit
                });
        }

        public Stage2PreflightReport AnalyzeStage2(Stage2Options options)
        {
            return _stage2Service.Analyze(options);
        }

        public Stage2WorkflowPlan PlanStage2(Stage2Options options)
        {
            return new Stage2WorkflowPlan(options, AnalyzeStage2(options));
        }

        public Stage2WorkflowResult CompleteStage2(Stage2WorkflowPlan plan, bool confirmed, Action<string> log)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            if (plan.RequiresConfirmation && !confirmed)
            {
                return Stage2WorkflowResult.Cancelled();
            }

            return Stage2WorkflowResult.Complete(RunStage2(plan.Options, log));
        }

        public StageWorkflowResult<Stage2Report> RunStage2(Stage2Options options, Action<string> log)
        {
            var report = _stage2Service.Run(options, log);
            return new StageWorkflowResult<Stage2Report>(
                report,
                new[]
                {
                    "阶段2完成。",
                    "汇总表：" + report.Summary,
                    "报告：" + report.ReportPath,
                    "代理费合计：" + report.ProxyTotal.ToString("0.####"),
                    "居间费合计：" + report.IntermediaryTotal.ToString("0.####")
                });
        }

        public StageWorkflowResult<EmployeeRewardResult> RunEmployeeReward(EmployeeRewardOptions options, Action<string> log)
        {
            if (_employeeRewardService == null)
            {
                throw new InvalidOperationException("员工电量奖励服务未配置。");
            }

            var report = _employeeRewardService.Run(options, log);
            return new StageWorkflowResult<EmployeeRewardResult>(
                report,
                new[]
                {
                    "员工电量奖励生成完成。",
                    "奖励总表：" + report.SummaryPath,
                    "报告：" + report.ReportPath,
                    "员工确认表：" + report.PersonalWorkbookPaths.Count + " 个",
                    "电量合计：" + report.TotalPower.ToString("0.####") + " 万千瓦时",
                    "奖励金额：" + report.TotalReward.ToString("0.##") + " 元"
                });
        }
    }
}
