using System;
using System.IO;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class ProvinceStage1Service
    {
        private readonly IProvinceStage1ExcelGateway _excel;

        public ProvinceStage1Service(IProvinceStage1ExcelGateway excel)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
        }

        public ProvinceStage1CleanResult CleanPowerData(ProvinceStage1CleanOptions options, Action<string> log)
        {
            ValidateCleanOptions(options);
            Directory.CreateDirectory(options.OutputDirectory);

            log?.Invoke("正在整理" + ProvinceDisplayNames.GetName(options.Province) + "本月电量。");
            var result = _excel.CleanPowerData(options);
            log?.Invoke("电量处理表已生成：" + result.OutputWorkbookPath);
            if (!string.IsNullOrWhiteSpace(result.HtmlReportPath))
            {
                log?.Invoke("检查报告已生成：" + result.HtmlReportPath);
            }
            return result;
        }

        public ProvinceStage1LedgerUpdatePlan PlanLedgerUpdate(ProvinceStage1LedgerUpdateOptions options, Action<string> log)
        {
            ValidateLedgerUpdateOptions(options);

            log?.Invoke("正在检查" + ProvinceDisplayNames.GetName(options.Province) + "本月台账资料。");
            return _excel.PlanLedgerUpdate(options);
        }

        public ProvinceStage1LedgerUpdateResult UpdateLedger(ProvinceStage1LedgerUpdateOptions options, Action<string> log)
        {
            ValidateLedgerUpdateOptions(options);
            Directory.CreateDirectory(options.OutputDirectory);

            log?.Invoke("正在生成" + ProvinceDisplayNames.GetName(options.Province) + "本月台账。");
            var result = _excel.UpdateLedger(options);
            if (!string.IsNullOrWhiteSpace(result.OutputPowerWorkbookPath))
            {
                log?.Invoke("电量处理表已生成：" + result.OutputPowerWorkbookPath);
            }
            log?.Invoke("本月台账已生成：" + result.OutputLedgerPath);
            if (!string.IsNullOrWhiteSpace(result.HtmlReportPath))
            {
                log?.Invoke("检查报告已生成：" + result.HtmlReportPath);
            }
            return result;
        }

        private static void ValidateCleanOptions(ProvinceStage1CleanOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Province == ProvinceCode.Hainan)
            {
                throw new NotSupportedException("海南阶段一使用独立的成熟工作流，不通过多省份阶段一电量清洗入口。");
            }

            if (string.IsNullOrWhiteSpace(options.RawDetailPath))
            {
                throw new ArgumentException("请选择原始电量确认结算单。");
            }

            if (!File.Exists(options.RawDetailPath))
            {
                throw new FileNotFoundException("原始电量确认结算单不存在。", options.RawDetailPath);
            }

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new ArgumentException("请选择结果输出文件夹。");
            }
        }

        private static void ValidateLedgerUpdateOptions(ProvinceStage1LedgerUpdateOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Province == ProvinceCode.Hainan)
            {
                throw new NotSupportedException("海南阶段一使用独立的成熟工作流，不通过多省份阶段一台账更新入口。");
            }

            if (options.Month < 1 || options.Month > 12)
            {
                throw new ArgumentException("请选择结算月份。");
            }

            if (string.IsNullOrWhiteSpace(options.LedgerPath))
            {
                throw new ArgumentException("请选择" + ProvinceDisplayNames.GetName(options.Province) + "售电结算台账。");
            }

            if (!File.Exists(options.LedgerPath))
            {
                throw new FileNotFoundException(
                    ProvinceDisplayNames.GetName(options.Province) + "售电结算台账不存在。",
                    options.LedgerPath);
            }

            if (string.IsNullOrWhiteSpace(options.RawDetailPath))
            {
                throw new ArgumentException("请选择交易中心电量确认结算单。");
            }

            if (!File.Exists(options.RawDetailPath))
            {
                throw new FileNotFoundException("交易中心电量确认结算单不存在。", options.RawDetailPath);
            }

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new ArgumentException("请选择结果输出文件夹。");
            }
        }

    }
}
