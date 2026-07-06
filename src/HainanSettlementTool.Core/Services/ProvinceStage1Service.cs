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
            Validate(options);
            Directory.CreateDirectory(options.OutputDirectory);

            log?.Invoke("正在清洗" + ProvinceName(options.Province) + "阶段一电量数据。");
            var result = _excel.CleanPowerData(options);
            log?.Invoke("电量处理表已生成：" + result.OutputWorkbookPath);
            log?.Invoke("校验报告已生成：" + result.ReportPath);
            return result;
        }

        private static void Validate(ProvinceStage1CleanOptions options)
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

        public static string ProvinceName(ProvinceCode province)
        {
            switch (province)
            {
                case ProvinceCode.Chongqing:
                    return "重庆";
                case ProvinceCode.Hainan:
                    return "海南";
                default:
                    return province.ToString();
            }
        }
    }
}
