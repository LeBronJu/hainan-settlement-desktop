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
            report.HtmlReportPath = Path.Combine(
                report.OutputDirectory,
                report.HasReviewItems
                    ? "【必须处理】" + (report.SkippedCount + report.FailedCount) + "个文件未生成新月份sheet.html"
                    : "广东" + report.Month + "月分表初始化结果.html");

            File.WriteAllText(
                report.ReportPath,
                JsonConvert.SerializeObject(report, Formatting.Indented),
                Encoding.UTF8);
            File.WriteAllText(report.ValidationReportPath, BuildValidationText(report), Encoding.UTF8);
            File.WriteAllText(report.HtmlReportPath, BuildHtml(report), Encoding.UTF8);
        }

        private static string BuildValidationText(GuangdongStage2MonthPreparationReport report)
        {
            var builder = new StringBuilder();
            builder.AppendLine("广东分表月份初始化校验报告");
            builder.AppendLine("结算月份：" + report.Year + "年" + report.Month + "月");
            builder.AppendLine("本批次状态：" + StatusText(report));
            builder.AppendLine();
            builder.AppendLine("扫描输入：" + report.InputCount);
            builder.AppendLine("正常输出：" + report.SuccessfulCount);
            builder.AppendLine("未自动处理但原文件已保留：" + report.PreservedReviewCopyCount);
            builder.AppendLine("失败：" + report.FailedCount);
            builder.AppendLine("输出目录可找到的 Excel：" + report.AvailableWorkbookCount);
            builder.AppendLine("输入分类守恒：" + (report.IsClassificationComplete ? "是" : "否"));
            builder.AppendLine("输入文件保留完整：" + (report.HasCompleteOutputSet ? "是" : "否"));
            builder.AppendLine();
            builder.AppendLine("从上月创建：" + report.CreatedCount);
            builder.AppendLine("整理现有目标月：" + report.NormalizedCount);
            builder.AppendLine("原本已经准备完成：" + report.AlreadyPreparedCount);
            builder.AppendLine("跳过：" + report.SkippedCount);

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
            builder.AppendLine("警告：在完成以下文件的人工处理前，本批次结果不能视为完整。");
            builder.AppendLine();
            builder.AppendLine("需要人工复核：");
            foreach (var item in reviewItems)
            {
                builder.AppendLine("- [" + item.SettlementKind + "] " + item.RelativePath);
                builder.AppendLine("  " + item.Message);
                builder.AppendLine(string.IsNullOrWhiteSpace(item.ReviewCopyPath)
                    ? "  原文件未能保留到输出目录。"
                    : "  原文件已保留：" + item.ReviewCopyPath);
            }

            return builder.ToString();
        }

        private static string BuildHtml(GuangdongStage2MonthPreparationReport report)
        {
            var title = "广东" + report.Month + "月分表初始化结果";
            var document = new ReadableReportDocument
            {
                Title = title,
                PeriodLabel = "结算月份：" + report.Year + "年" + report.Month + "月",
                Status = report.HasCriticalFailures
                    ? ReadableReportStatus.Critical
                    : report.HasReviewItems ? ReadableReportStatus.Review : ReadableReportStatus.Success,
                StatusText = StatusText(report)
            };
            document.Metrics.Add(new ReadableReportMetric("扫描输入", report.InputCount.ToString()));
            document.Metrics.Add(new ReadableReportMetric("正常输出", report.SuccessfulCount.ToString()));
            document.Metrics.Add(new ReadableReportMetric("原文件已保留", report.PreservedReviewCopyCount.ToString()));
            document.Metrics.Add(new ReadableReportMetric("失败", report.FailedCount.ToString()));

            var reviewItems = report.Workbooks
                .Where(item => item.Action == GuangdongStage2PreparationActions.Skipped
                    || item.Action == GuangdongStage2PreparationActions.Failed)
                .ToList();
            if (reviewItems.Count > 0)
            {
                document.Notices.Add(new ReadableReportNotice(
                    "警告",
                    "在完成以下文件的人工处理前，本批次结果不能视为完整。",
                    report.HasCriticalFailures
                        ? ReadableReportNoticeTone.Critical
                        : ReadableReportNoticeTone.Review));
                var reviewSection = new ReadableReportSection(
                    "需要人工复核",
                    "类型",
                    "文件",
                    "原因",
                    "原文件保留位置");
                foreach (var item in reviewItems)
                {
                    reviewSection.Rows.Add(new ReadableReportRow(
                        item.SettlementKind,
                        item.RelativePath,
                        item.Message,
                        string.IsNullOrWhiteSpace(item.ReviewCopyPath) ? "原文件未能保留" : item.ReviewCopyPath));
                }
                document.Sections.Add(reviewSection);
            }
            else
            {
                document.Sections.Add(new ReadableReportSection("需要人工复核")
                {
                    EmptyMessage = "没有需要人工复核的文件。"
                });
            }

            var detailSection = new ReadableReportSection("处理明细", "项目", "结果");
            detailSection.Rows.Add(new ReadableReportRow("从上月创建", report.CreatedCount.ToString()));
            detailSection.Rows.Add(new ReadableReportRow("整理现有目标月", report.NormalizedCount.ToString()));
            detailSection.Rows.Add(new ReadableReportRow("原本已经准备完成", report.AlreadyPreparedCount.ToString()));
            detailSection.Rows.Add(new ReadableReportRow("跳过", report.SkippedCount.ToString()));
            detailSection.Rows.Add(new ReadableReportRow("输入分类守恒", report.IsClassificationComplete ? "是" : "否"));
            detailSection.Rows.Add(new ReadableReportRow("输入文件保留完整", report.HasCompleteOutputSet ? "是" : "否"));
            document.Sections.Add(detailSection);
            return ReadableHtmlReportRenderer.Render(document);
        }

        private static string StatusText(GuangdongStage2MonthPreparationReport report)
        {
            if (report.HasCriticalFailures)
            {
                return "执行异常，必须处理";
            }

            return report.HasReviewItems
                ? "部分完成，必须人工复核"
                : "全部完成";
        }
    }
}
