using System;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class HainanStage2Service
    {
        private readonly IHainanStage2ExcelGateway _excel;

        public HainanStage2Service(IHainanStage2ExcelGateway excel)
        {
            _excel = excel;
        }

        public HainanStage2Report Run(HainanStage2Options options, Action<string> log)
        {
            Validate(options);
            RequireConfirmedPreflight(options.ExpectedPreflightSignature, options.ExpectedInputFingerprint, "海南");

            log?.Invoke("正在读取人工整理后的台账，并生成代理/居间分表和汇总表。");
            return _excel.GenerateSettlement(options);
        }

        public HainanStage2PreflightReport Analyze(HainanStage2Options options)
        {
            Validate(options);
            return _excel.AnalyzeSettlement(options);
        }

        private static void Validate(HainanStage2Options options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Month <= 1)
            {
                throw new ArgumentException("阶段2要求月份大于1。");
            }

            FileAccessGuard.RequireReadableWorkbook(options.LedgerPath, "人工整理后的台账");
            RequireExistingDirectory(options.ProxyTemplateDirectory, "上月代理分表文件夹");
            RequireExistingDirectory(options.IntermediaryTemplateDirectory, "上月居间分表文件夹");
            FileAccessGuard.RequireReadableWorkbook(options.SummaryTemplatePath, "上月/修正版汇总表");

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new ArgumentException("请选择输出文件夹。");
            }

            ValidateOutputSummaryName(options.OutputSummaryName);
            ValidateSummarySubjectDecisions(options);
            ValidateTemplateDecisions(options);
        }

        private static void ValidateOutputSummaryName(string outputSummaryName)
        {
            if (string.IsNullOrWhiteSpace(outputSummaryName))
            {
                return;
            }

            var invalid = false;
            try
            {
                var baseName = Path.GetFileNameWithoutExtension(outputSummaryName);
                invalid = Path.IsPathRooted(outputSummaryName)
                    || outputSummaryName.IndexOf(Path.DirectorySeparatorChar) >= 0
                    || outputSummaryName.IndexOf(Path.AltDirectorySeparatorChar) >= 0
                    || outputSummaryName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
                    || outputSummaryName == "."
                    || outputSummaryName == ".."
                    || !string.Equals(Path.GetFileName(outputSummaryName), outputSummaryName, StringComparison.Ordinal)
                    || !string.Equals(Path.GetExtension(outputSummaryName), ".xlsx", StringComparison.OrdinalIgnoreCase)
                    || string.IsNullOrWhiteSpace(baseName)
                    || outputSummaryName.EndsWith(" ", StringComparison.Ordinal)
                    || outputSummaryName.EndsWith(".", StringComparison.Ordinal)
                    || IsReservedWindowsFileName(baseName);
            }
            catch (ArgumentException)
            {
                invalid = true;
            }

            if (invalid)
            {
                throw new ArgumentException(
                    "海南阶段二输出汇总表名必须是合法的纯 .xlsx 文件名，不能使用绝对路径、目录分隔符或 . / .. 路径。",
                    nameof(outputSummaryName));
            }
        }

        private static bool IsReservedWindowsFileName(string baseName)
        {
            var name = (baseName ?? string.Empty).TrimEnd(' ', '.');
            return new[]
            {
                "CON", "PRN", "AUX", "NUL",
                "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
                "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            }.Contains(name, StringComparer.OrdinalIgnoreCase);
        }

        private static void ValidateSummarySubjectDecisions(HainanStage2Options options)
        {
            foreach (var decision in options.SummarySubjectDecisions)
            {
                if (decision == null)
                {
                    throw new ArgumentException("海南阶段二新增汇总主体决策不能为空。");
                }

                if (string.IsNullOrWhiteSpace(decision.SettlementKind))
                {
                    throw new ArgumentException("海南阶段二新增汇总主体决策缺少结算类型。");
                }

                if (string.IsNullOrWhiteSpace(decision.Entity))
                {
                    throw new ArgumentException("海南阶段二新增汇总主体决策缺少主体名称。");
                }

                if (!HainanStage2PaymentParties.Supported.Contains(decision.PaymentParty))
                {
                    throw new ArgumentException("海南阶段二新增汇总主体支付方只能选择清能或清辉。");
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

        private static void RequireConfirmedPreflight(
            string preflightSignature,
            string inputFingerprint,
            string province)
        {
            if (string.IsNullOrWhiteSpace(preflightSignature)
                || string.IsNullOrWhiteSpace(inputFingerprint))
            {
                throw new InvalidOperationException(
                    province + "阶段二正式生成必须先完成本次预检，不能直接绕过预检生成。");
            }
        }

        private static void ValidateTemplateDecisions(HainanStage2Options options)
        {
            foreach (var decision in options.TemplateDecisions)
            {
                if (decision == null)
                {
                    throw new ArgumentException("海南阶段二模板选择不能为空。");
                }

                if (string.IsNullOrWhiteSpace(decision.SettlementKind))
                {
                    throw new ArgumentException("海南阶段二模板选择缺少结算类型。");
                }

                if (string.IsNullOrWhiteSpace(decision.Entity))
                {
                    throw new ArgumentException("海南阶段二模板选择缺少主体名称。");
                }

                if (string.IsNullOrWhiteSpace(decision.TemplatePath))
                {
                    throw new ArgumentException("海南阶段二模板选择缺少模板文件路径。");
                }
            }
        }
    }
}
