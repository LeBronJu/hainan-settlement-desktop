using System.IO;
using System.Linq;
using System.Text;
using HainanSettlementTool.Core.Models;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal sealed class GuangdongStage2ReportWriter
    {
        public void Write(GuangdongStage2MonthPreparationReport report)
        {
            report.ReportPath = Path.Combine(
                report.OutputDirectory,
                "广东" + report.Month + "月分表初始化报告.json");
            report.ValidationReportPath = Path.Combine(
                report.OutputDirectory,
                "广东" + report.Month + "月分表初始化校验报告.txt");

            File.WriteAllText(
                report.ReportPath,
                JsonConvert.SerializeObject(report, Formatting.Indented),
                Encoding.UTF8);
            File.WriteAllText(report.ValidationReportPath, BuildValidationText(report), Encoding.UTF8);
        }

        private static string BuildValidationText(GuangdongStage2MonthPreparationReport report)
        {
            var builder = new StringBuilder();
            builder.AppendLine("广东分表月份初始化校验报告");
            builder.AppendLine("结算月份：" + report.Year + "年" + report.Month + "月");
            builder.AppendLine("成功输出：" + report.SuccessfulCount);
            builder.AppendLine("从上月创建：" + report.CreatedCount);
            builder.AppendLine("整理现有目标月：" + report.NormalizedCount);
            builder.AppendLine("原本已经准备完成：" + report.AlreadyPreparedCount);
            builder.AppendLine("跳过：" + report.SkippedCount);
            builder.AppendLine("失败：" + report.FailedCount);

            var reviewItems = report.Workbooks
                .Where(item => item.Action == GuangdongStage2PreparationActions.Skipped
                    || item.Action == GuangdongStage2PreparationActions.Failed)
                .ToList();
            if (reviewItems.Count == 0)
            {
                builder.AppendLine();
                builder.AppendLine("没有需要人工复核的文件。");
                return builder.ToString();
            }

            builder.AppendLine();
            builder.AppendLine("需要人工复核：");
            foreach (var item in reviewItems)
            {
                builder.AppendLine("- [" + item.SettlementKind + "] " + item.RelativePath);
                builder.AppendLine("  " + item.Message);
            }

            return builder.ToString();
        }
    }
}
