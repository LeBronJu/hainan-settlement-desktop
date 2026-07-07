using System;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class SettlementWorkflow
    {
        private readonly HainanStage1Service _stage1Service;
        private readonly HainanStage2Service _stage2Service;
        private readonly EmployeeRewardService _employeeRewardService;
        private readonly ProvinceStage1Service _provinceStage1Service;
        private readonly ChongqingStage2Service _chongqingStage2Service;

        public SettlementWorkflow(HainanStage1Service stage1Service, HainanStage2Service stage2Service)
            : this(stage1Service, stage2Service, null)
        {
        }

        public SettlementWorkflow(HainanStage1Service stage1Service, HainanStage2Service stage2Service, EmployeeRewardService employeeRewardService)
            : this(stage1Service, stage2Service, employeeRewardService, null)
        {
        }

        public SettlementWorkflow(
            HainanStage1Service stage1Service,
            HainanStage2Service stage2Service,
            EmployeeRewardService employeeRewardService,
            ProvinceStage1Service provinceStage1Service)
            : this(stage1Service, stage2Service, employeeRewardService, provinceStage1Service, null)
        {
        }

        public SettlementWorkflow(
            HainanStage1Service stage1Service,
            HainanStage2Service stage2Service,
            EmployeeRewardService employeeRewardService,
            ProvinceStage1Service provinceStage1Service,
            ChongqingStage2Service chongqingStage2Service)
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
            _chongqingStage2Service = chongqingStage2Service;
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
                    ProvinceDisplayNames.GetName(report.Province) + "阶段一电量清洗完成。",
                    "电量处理表：" + report.OutputWorkbookPath,
                    "报告：" + report.ReportPath,
                    "客户数量：" + report.CustomerRows + "，户号数量：" + report.AccountRows,
                    "合计电量：" + report.TotalPower.ToString("0.####") + " " + report.Unit
                });
        }

        public ProvinceStage1LedgerUpdatePlan PlanProvinceStage1LedgerUpdate(
            ProvinceStage1LedgerUpdateOptions options,
            Action<string> log)
        {
            if (_provinceStage1Service == null)
            {
                throw new InvalidOperationException("多省份阶段一服务未配置。");
            }

            return _provinceStage1Service.PlanLedgerUpdate(options, log);
        }

        public StageWorkflowResult<ProvinceStage1LedgerUpdateResult> UpdateProvinceStage1Ledger(
            ProvinceStage1LedgerUpdateOptions options,
            Action<string> log)
        {
            if (_provinceStage1Service == null)
            {
                throw new InvalidOperationException("多省份阶段一服务未配置。");
            }

            var report = _provinceStage1Service.UpdateLedger(options, log);
            return new StageWorkflowResult<ProvinceStage1LedgerUpdateResult>(
                report,
                new[]
                {
                    ProvinceDisplayNames.GetName(report.Province) + "阶段一台账更新完成。",
                    "输出台账：" + report.OutputLedgerPath,
                    "报告：" + report.ReportPath,
                    "匹配客户：" + report.MatchedRows + "，写入电量：" + report.UpdatedPowerRows,
                    "人工匹配：" + report.ManualMatchedRows,
                    "多户号提示：" + report.MultiAccountRows,
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

        public ChongqingStage2PreflightReport AnalyzeChongqingStage2(ChongqingStage2Options options)
        {
            if (_chongqingStage2Service == null)
            {
                throw new InvalidOperationException("重庆阶段二服务未配置。");
            }

            return _chongqingStage2Service.Analyze(options);
        }

        public ChongqingStage2WorkflowPlan PlanChongqingStage2(ChongqingStage2Options options)
        {
            return new ChongqingStage2WorkflowPlan(options, AnalyzeChongqingStage2(options));
        }

        public ChongqingStage2WorkflowResult CompleteChongqingStage2(
            ChongqingStage2WorkflowPlan plan,
            bool confirmed,
            Action<string> log)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            if (plan.RequiresConfirmation && !confirmed)
            {
                return ChongqingStage2WorkflowResult.Cancelled();
            }

            return ChongqingStage2WorkflowResult.Complete(RunChongqingStage2(plan.Options, log));
        }

        public StageWorkflowResult<ChongqingStage2Report> RunChongqingStage2(
            ChongqingStage2Options options,
            Action<string> log)
        {
            if (_chongqingStage2Service == null)
            {
                throw new InvalidOperationException("重庆阶段二服务未配置。");
            }

            var report = _chongqingStage2Service.Run(options, log);
            return new StageWorkflowResult<ChongqingStage2Report>(
                report,
                new[]
                {
                    "重庆阶段二完成。",
                    "汇总表：" + report.Summary,
                    "报告：" + report.ReportPath,
                    "代理费合计：" + report.ProxyTotal.ToString("0.####"),
                    "居间费合计：" + report.IntermediaryTotal.ToString("0.####"),
                    "退补电费合计：" + report.RefundTotal.ToString("0.####")
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
