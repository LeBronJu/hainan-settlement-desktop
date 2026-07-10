using System;
using System.Collections.Generic;
using System.IO;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class GuangdongStage2MonthPreparationService
    {
        private readonly IGuangdongStage2MonthPreparationExcelGateway _excel;

        public GuangdongStage2MonthPreparationService(IGuangdongStage2MonthPreparationExcelGateway excel)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
        }

        public GuangdongStage2PreflightReport Analyze(GuangdongStage2MonthPreparationOptions options)
        {
            Validate(options);
            return _excel.AnalyzeMonthPreparation(options);
        }

        public GuangdongStage2MonthPreparationReport Run(GuangdongStage2MonthPreparationOptions options, Action<string> log)
        {
            Validate(options);
            Directory.CreateDirectory(options.OutputDirectory);

            log?.Invoke("正在初始化广东代理/居间/退补分表的目标月份 sheet。");
            var report = _excel.GenerateMonthPreparation(options);
            log?.Invoke("广东分表月份初始化完成：" + report.OutputDirectory);
            log?.Invoke("广东分表月份初始化报告：" + report.ReportPath);
            return report;
        }

        private static void Validate(GuangdongStage2MonthPreparationOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Year < 2000 || options.Year > 2100)
            {
                throw new ArgumentException("广东分表月份初始化年份无效。");
            }

            if (options.Month <= 1 || options.Month > 12)
            {
                throw new ArgumentException("广东分表月份初始化当前支持2月至12月。");
            }

            var directories = new List<string>();
            AddDirectoryIfPresent(directories, options.ProxyDirectory, "广东代理分表文件夹");
            AddDirectoryIfPresent(directories, options.IntermediaryDirectory, "广东居间分表文件夹");
            AddDirectoryIfPresent(directories, options.RefundDirectory, "广东退补分表文件夹");
            if (directories.Count == 0)
            {
                throw new ArgumentException("请至少选择一个广东代理、居间或退补分表文件夹。");
            }

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new ArgumentException("请选择输出文件夹。");
            }

            var output = NormalizePath(options.OutputDirectory);
            foreach (var input in directories)
            {
                if (IsSameOrChildPath(output, NormalizePath(input)))
                {
                    throw new ArgumentException("输出文件夹不能位于广东分表输入文件夹内部：" + input);
                }
            }
        }

        private static void AddDirectoryIfPresent(ICollection<string> directories, string path, string label)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            if (!Directory.Exists(path))
            {
                throw new DirectoryNotFoundException(label + "不存在：" + path);
            }

            directories.Add(path);
        }

        private static string NormalizePath(string path)
        {
            return Path.GetFullPath(path)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        private static bool IsSameOrChildPath(string candidate, string parent)
        {
            if (string.Equals(candidate, parent, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return candidate.StartsWith(parent + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                || candidate.StartsWith(parent + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }
    }
}
