using System;
using System.IO;
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

        public Stage2Report Run(Stage2Options options, Action<string> log)
        {
            Validate(options);
            Directory.CreateDirectory(options.OutputDirectory);

            log?.Invoke("正在读取人工整理后的台账，并生成代理/居间分表和汇总表。");
            return _excel.GenerateSettlement(options);
        }

        public Stage2PreflightReport Analyze(Stage2Options options)
        {
            Validate(options);
            return _excel.AnalyzeSettlement(options);
        }

        private static void Validate(Stage2Options options)
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
    }
}
