using System;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class SettlementWorkflow
    {
        private readonly Stage1Service _stage1Service;
        private readonly Stage2Service _stage2Service;

        public SettlementWorkflow(Stage1Service stage1Service, Stage2Service stage2Service)
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

        public Stage2PreflightReport AnalyzeStage2(Stage2Options options)
        {
            return _stage2Service.Analyze(options);
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
    }
}
