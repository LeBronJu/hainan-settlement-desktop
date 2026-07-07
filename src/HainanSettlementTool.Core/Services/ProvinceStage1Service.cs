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

            log?.Invoke("正在清洗" + ProvinceDisplayNames.GetName(options.Province) + "阶段一电量数据。");
            var result = _excel.CleanPowerData(options);
            log?.Invoke("电量处理表已生成：" + result.OutputWorkbookPath);
            log?.Invoke("校验报告已生成：" + result.ReportPath);
            return result;
        }

        public ProvinceStage1LedgerUpdatePlan PlanLedgerUpdate(ProvinceStage1LedgerUpdateOptions options, Action<string> log)
        {
            ValidateLedgerUpdateOptions(options);

            log?.Invoke("正在预检" + ProvinceDisplayNames.GetName(options.Province) + "阶段一台账更新。");
            return _excel.PlanLedgerUpdate(options);
        }

        public ProvinceStage1LedgerUpdateResult UpdateLedger(ProvinceStage1LedgerUpdateOptions options, Action<string> log)
        {
            ValidateLedgerUpdateOptions(options);
            Directory.CreateDirectory(options.OutputDirectory);

            log?.Invoke("正在更新" + ProvinceDisplayNames.GetName(options.Province) + "阶段一台账。");
            var result = _excel.UpdateLedger(options);
            log?.Invoke("台账更新结果已生成：" + result.OutputLedgerPath);
            log?.Invoke("台账更新报告已生成：" + result.ReportPath);
            return result;
        }

        private static void ValidateCleanOptions(ProvinceStage1CleanOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Province != ProvinceCode.Chongqing)
            {
                throw new NotSupportedException("当前仅支持重庆阶段一电量清洗。");
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

            if (options.Province != ProvinceCode.Chongqing)
            {
                throw new NotSupportedException("当前仅支持重庆阶段一台账更新。");
            }

            if (options.Month <= 0)
            {
                throw new ArgumentException("请选择结算月份。");
            }

            if (string.IsNullOrWhiteSpace(options.LedgerPath))
            {
                throw new ArgumentException("请选择重庆售电结算台账。");
            }

            if (!File.Exists(options.LedgerPath))
            {
                throw new FileNotFoundException("重庆售电结算台账不存在。", options.LedgerPath);
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
