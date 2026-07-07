using System;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class ChongqingStage2Service
    {
        private readonly IChongqingStage2ExcelGateway _excel;

        public ChongqingStage2Service(IChongqingStage2ExcelGateway excel)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
        }

        public ChongqingStage2PreflightReport Analyze(ChongqingStage2Options options)
        {
            Validate(options);
            return _excel.AnalyzeSettlement(options);
        }

        public ChongqingStage2Report Run(ChongqingStage2Options options, Action<string> log)
        {
            Validate(options);
            Directory.CreateDirectory(options.OutputDirectory);

            log?.Invoke("正在生成重庆阶段二代理/居间/退补分表和汇总表。");
            var report = _excel.GenerateSettlement(options);
            log?.Invoke("重庆阶段二汇总表已生成：" + report.Summary);
            log?.Invoke("重庆阶段二报告已生成：" + report.ReportPath);
            return report;
        }

        private static void Validate(ChongqingStage2Options options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Month <= 1)
            {
                throw new ArgumentException("重庆阶段二要求月份大于1。");
            }

            FileAccessGuard.RequireReadableWorkbook(options.LedgerPath, "重庆售电结算台账");
            RequireExistingDirectory(options.ProxyTemplateDirectory, "重庆代理分表文件夹");
            RequireOptionalExistingDirectory(options.IntermediaryTemplateDirectory, "重庆居间分表文件夹");
            RequireExistingDirectory(options.RefundTemplateDirectory, "重庆退补分表文件夹");
            FileAccessGuard.RequireReadableWorkbook(options.SummaryTemplatePath, "重庆代理费汇总表模板");

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new ArgumentException("请选择输出文件夹。");
            }

            ValidateSummarySubjectDecisions(options);
        }

        private static void ValidateSummarySubjectDecisions(ChongqingStage2Options options)
        {
            foreach (var decision in options.SummarySubjectDecisions)
            {
                if (decision == null)
                {
                    throw new ArgumentException("重庆阶段二新增汇总主体决策不能为空。");
                }

                if (string.IsNullOrWhiteSpace(decision.SettlementKind))
                {
                    throw new ArgumentException("重庆阶段二新增汇总主体决策缺少结算类型。");
                }

                if (string.IsNullOrWhiteSpace(decision.Entity))
                {
                    throw new ArgumentException("重庆阶段二新增汇总主体决策缺少主体名称。");
                }

                if (!ChongqingStage2PaymentParties.Supported.Contains(decision.PaymentParty))
                {
                    throw new ArgumentException("重庆阶段二新增汇总主体支付方只能选择清能或清辉。");
                }
            }
        }

        private static void RequireExistingDirectory(string path, string label)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("请选择" + label + "。");
            }

            if (!Directory.Exists(path))
            {
                throw new DirectoryNotFoundException(label + "不存在：" + path);
            }
        }

        private static void RequireOptionalExistingDirectory(string path, string label)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            if (!Directory.Exists(path))
            {
                throw new DirectoryNotFoundException(label + "不存在：" + path);
            }
        }
    }
}
