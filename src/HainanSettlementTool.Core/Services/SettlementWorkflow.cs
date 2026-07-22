using System;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class SettlementWorkflow
    {
        private readonly HainanStage1Service _stage1Service;
        private readonly HainanStage2Service _hainanStage2Service;
        private readonly HainanEmployeePowerRewardService _hainanEmployeePowerRewardService;
        private readonly ProvinceStage1Service _provinceStage1Service;
        private readonly ChongqingStage2Service _chongqingStage2Service;

        public SettlementWorkflow(HainanStage1Service stage1Service, HainanStage2Service stage2Service)
            : this(stage1Service, stage2Service, null)
        {
        }

        public SettlementWorkflow(HainanStage1Service stage1Service, HainanStage2Service stage2Service, HainanEmployeePowerRewardService hainanEmployeePowerRewardService)
            : this(stage1Service, stage2Service, hainanEmployeePowerRewardService, null)
        {
        }

        public SettlementWorkflow(
            HainanStage1Service stage1Service,
            HainanStage2Service stage2Service,
            HainanEmployeePowerRewardService hainanEmployeePowerRewardService,
            ProvinceStage1Service provinceStage1Service)
            : this(stage1Service, stage2Service, hainanEmployeePowerRewardService, provinceStage1Service, null)
        {
        }

        public SettlementWorkflow(
            HainanStage1Service stage1Service,
            HainanStage2Service stage2Service,
            HainanEmployeePowerRewardService hainanEmployeePowerRewardService,
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
            _hainanStage2Service = stage2Service;
            _hainanEmployeePowerRewardService = hainanEmployeePowerRewardService;
            _provinceStage1Service = provinceStage1Service;
            _chongqingStage2Service = chongqingStage2Service;
        }

        public StageWorkflowResult<HainanStage1Report> RunHainanStage1(HainanStage1Options options, Action<string> log)
        {
            var report = _stage1Service.Run(options, log);
            return new StageWorkflowResult<HainanStage1Report>(
                report,
                new[]
                {
                    "阶段1完成。",
                    "输出台账：" + report.Output,
                    "报告：" + report.ReportPath
                });
        }

        public StageWorkflowResult<HainanPowerCleanReport> CleanHainanPowerData(string rawDetailPath, string outputPath, Action<string> log)
        {
            var report = _stage1Service.CleanHainanPowerData(rawDetailPath, outputPath, log);
            return new StageWorkflowResult<HainanPowerCleanReport>(
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

        public HainanStage2PreflightReport AnalyzeHainanStage2(HainanStage2Options options)
        {
            return _hainanStage2Service.Analyze(options);
        }

        public HainanStage2WorkflowPlan PlanHainanStage2(HainanStage2Options options)
        {
            var preflight = AnalyzeHainanStage2(options);
            options.ExpectedPreflightSignature = preflight.PreflightSignature;
            options.ExpectedInputFingerprint = preflight.InputFingerprint;
            return new HainanStage2WorkflowPlan(options, preflight);
        }

        public HainanStage2WorkflowResult CompleteHainanStage2(HainanStage2WorkflowPlan plan, bool confirmed, Action<string> log)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var evaluation = plan.Evaluation;
            if (evaluation.HasBlockingIssues)
            {
                throw new InvalidOperationException("海南阶段二预检存在阻断项，请修正台账或模板后重新预检。");
            }

            if (!confirmed && evaluation.RequiresConfirmation)
            {
                return HainanStage2WorkflowResult.Cancelled();
            }

            if (!evaluation.CanContinue)
            {
                throw new InvalidOperationException("海南阶段二仍有未完成或无效的必选项，请完成支付方选择后再生成。");
            }

            if (!evaluation.CanGenerate(confirmed))
            {
                return HainanStage2WorkflowResult.Cancelled();
            }

            EnsureHainanStage2PlanIsCurrent(plan);
            return HainanStage2WorkflowResult.Complete(RunHainanStage2(plan.Options, log));
        }

        public StageWorkflowResult<HainanStage2Report> RunHainanStage2(HainanStage2Options options, Action<string> log)
        {
            var report = _hainanStage2Service.Run(options, log);
            return new StageWorkflowResult<HainanStage2Report>(
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
            var preflight = AnalyzeChongqingStage2(options);
            options.ExpectedPreflightSignature = preflight.PreflightSignature;
            options.ExpectedInputFingerprint = preflight.InputFingerprint;
            return new ChongqingStage2WorkflowPlan(options, preflight);
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

            var evaluation = plan.Evaluation;
            if (evaluation.HasBlockingIssues)
            {
                throw new InvalidOperationException("重庆阶段二预检存在阻断项，请修正台账或模板后重新预检。");
            }

            if (!confirmed && evaluation.RequiresConfirmation)
            {
                return ChongqingStage2WorkflowResult.Cancelled();
            }

            if (!evaluation.CanContinue)
            {
                throw new InvalidOperationException("重庆阶段二仍有未完成或无效的必选项，请完成支付方选择后再生成。");
            }

            if (!evaluation.CanGenerate(confirmed))
            {
                return ChongqingStage2WorkflowResult.Cancelled();
            }

            EnsureChongqingStage2PlanIsCurrent(plan);
            return ChongqingStage2WorkflowResult.Complete(RunChongqingStage2(plan.Options, log));
        }

        private void EnsureHainanStage2PlanIsCurrent(HainanStage2WorkflowPlan plan)
        {
            if (string.IsNullOrWhiteSpace(plan.Options.ExpectedPreflightSignature)
                && string.IsNullOrWhiteSpace(plan.Options.ExpectedInputFingerprint))
            {
                return;
            }

            var current = AnalyzeHainanStage2(plan.Options);
            EnsureStage2PlanSignaturesMatch(
                "海南",
                plan.Options.ExpectedPreflightSignature,
                current.PreflightSignature,
                plan.Options.ExpectedInputFingerprint,
                current.InputFingerprint);
        }

        private void EnsureChongqingStage2PlanIsCurrent(ChongqingStage2WorkflowPlan plan)
        {
            if (string.IsNullOrWhiteSpace(plan.Options.ExpectedPreflightSignature)
                && string.IsNullOrWhiteSpace(plan.Options.ExpectedInputFingerprint))
            {
                return;
            }

            var current = AnalyzeChongqingStage2(plan.Options);
            EnsureStage2PlanSignaturesMatch(
                "重庆",
                plan.Options.ExpectedPreflightSignature,
                current.PreflightSignature,
                plan.Options.ExpectedInputFingerprint,
                current.InputFingerprint);
        }

        private static void EnsureStage2PlanSignaturesMatch(
            string province,
            string expectedPreflight,
            string currentPreflight,
            string expectedInput,
            string currentInput)
        {
            var preflightMatches = string.IsNullOrWhiteSpace(expectedPreflight)
                || Stage2PreflightSignature.Matches(expectedPreflight, currentPreflight);
            var inputMatches = string.IsNullOrWhiteSpace(expectedInput)
                || Stage2PreflightSignature.Matches(expectedInput, currentInput);
            if (!preflightMatches || !inputMatches)
            {
                throw new InvalidOperationException(
                    province
                    + "阶段二输入或预检项目在确认后发生变化。为避免使用未确认的新资料，本次未生成；请重新打开预检并确认。" );
            }
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

        public StageWorkflowResult<HainanEmployeePowerRewardResult> RunHainanEmployeePowerReward(HainanEmployeePowerRewardOptions options, Action<string> log)
        {
            if (_hainanEmployeePowerRewardService == null)
            {
                throw new InvalidOperationException("员工电量奖励服务未配置。");
            }

            var report = _hainanEmployeePowerRewardService.Run(options, log);
            return new StageWorkflowResult<HainanEmployeePowerRewardResult>(
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
