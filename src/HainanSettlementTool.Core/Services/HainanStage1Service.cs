using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class HainanStage1Service
    {
        private readonly IHainanStage1ExcelGateway _excel;

        public HainanStage1Service(IHainanStage1ExcelGateway excel)
        {
            _excel = excel;
        }

        public HainanStage1Report Run(HainanStage1Options options, Action<string> log)
        {
            Validate(options);
            Directory.CreateDirectory(options.OutputDirectory);

            if (!string.IsNullOrWhiteSpace(options.RawDetailPath))
            {
                log?.Invoke("正在清洗原始零售侧明细，生成电量处理表。");
                CleanHainanPowerData(options.RawDetailPath, options.PowerPath, log);
            }
            else if (!File.Exists(options.PowerPath))
            {
                throw new FileNotFoundException("找不到电量处理表，也没有可清洗的原始零售侧明细。", options.PowerPath);
            }

            log?.Invoke("正在把电量、新增客户名称和户号导入台账。");
            return _excel.UpdateLedger(options);
        }

        public HainanPowerCleanReport CleanHainanPowerData(string rawDetailPath, string outputPath, Action<string> log)
        {
            ValidateCleanPower(rawDetailPath, outputPath);
            var outputDirectory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            var rawRows = _excel.ReadRawPowerRows(rawDetailPath);
            _excel.WritePowerWorkbook(rawRows, outputPath);
            log?.Invoke("电量处理表已生成：" + outputPath);

            var grouped = rawRows
                .Where(row => !string.IsNullOrWhiteSpace(row.Key))
                .GroupBy(row => row.Key)
                .Select(group => new
                {
                    Total = group.Sum(row => row.Total)
                })
                .ToList();

            return new HainanPowerCleanReport
            {
                RawDetailPath = rawDetailPath,
                OutputPath = outputPath,
                RawRows = rawRows.Count,
                PowerRows = grouped.Count,
                MonthTotal = Math.Round(grouped.Sum(row => row.Total), 4)
            };
        }

        private static void Validate(HainanStage1Options options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Month <= 1)
            {
                throw new ArgumentException("阶段1要求月份大于1。");
            }

            FileAccessGuard.RequireReadableWorkbook(options.BaseLedgerPath, "基础台账");
            if (string.IsNullOrWhiteSpace(options.PowerPath))
            {
                throw new ArgumentException("请选择电量处理表，或选择原始零售侧明细让程序生成。");
            }

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new ArgumentException("请选择输出文件夹。");
            }

            if (!string.IsNullOrWhiteSpace(options.ReferenceLedgerPath))
            {
                FileAccessGuard.RequireReadableWorkbook(options.ReferenceLedgerPath, "参考台账");
            }

            if (!string.IsNullOrWhiteSpace(options.RawDetailPath) && !File.Exists(options.RawDetailPath))
            {
                throw new FileNotFoundException("原始零售侧明细不存在。", options.RawDetailPath);
            }
        }

        private static void ValidateCleanPower(string rawDetailPath, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(rawDetailPath))
            {
                throw new ArgumentException("请选择原始零售侧明细。");
            }

            if (!File.Exists(rawDetailPath))
            {
                throw new FileNotFoundException("原始零售侧明细不存在。", rawDetailPath);
            }

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("请选择输出文件夹，程序会在其中生成电量处理表。");
            }
        }
    }
}
