using System.IO;
using System.Linq;
using System.Net;
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
            var builder = new StringBuilder();
            var title = "广东" + report.Month + "月分表初始化结果";
            var statusText = StatusText(report);
            var statusClass = report.HasCriticalFailures
                ? "danger"
                : report.HasReviewItems ? "warning" : "success";
            builder.AppendLine("<!DOCTYPE html>");
            builder.AppendLine("<html lang='zh-CN'><head><meta charset='utf-8'>");
            builder.AppendLine("<meta name='viewport' content='width=device-width, initial-scale=1'>");
            builder.AppendLine("<title>" + Html(title) + "</title>");
            builder.AppendLine("<style>");
            builder.AppendLine("body{font-family:'Microsoft YaHei UI','Microsoft YaHei',sans-serif;background:#f4f6f8;color:#17202a;margin:0;padding:32px}");
            builder.AppendLine(".page{max-width:1100px;margin:auto;background:#fff;border-radius:14px;box-shadow:0 8px 28px rgba(20,35,55,.10);overflow:hidden}");
            builder.AppendLine("header{padding:28px 32px;background:#153a5b;color:#fff}h1{margin:0 0 8px;font-size:26px}.period{opacity:.85}");
            builder.AppendLine(".status{margin:24px 32px 0;padding:16px 18px;border-radius:10px;font-size:20px;font-weight:700}.warning{background:#fff3cd;color:#7a4b00;border:1px solid #f0cf68}.danger{background:#fdecea;color:#9a2018;border:1px solid #e9a29d}.success{background:#e8f5e9;color:#236b2b;border:1px solid #9bd0a0}");
            builder.AppendLine(".cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;padding:22px 32px}.card{background:#f7f9fb;border:1px solid #e1e7ed;border-radius:10px;padding:16px}.label{color:#607080;font-size:13px}.value{font-size:28px;font-weight:700;margin-top:4px}");
            builder.AppendLine("section{padding:0 32px 30px}h2{font-size:18px;margin:8px 0 14px}table{width:100%;border-collapse:collapse;font-size:14px}th,td{text-align:left;vertical-align:top;padding:11px 10px;border-bottom:1px solid #e6ebef}th{background:#eef3f7;color:#33495c}td.path{word-break:break-all;color:#465f73}.empty{padding:18px;background:#eef8ef;border-radius:8px;color:#236b2b}");
            builder.AppendLine(".notice{margin:0 32px 24px;padding:14px 16px;background:#fff7e6;border-left:5px solid #e89a18}.footer{padding:18px 32px;background:#f7f9fb;color:#677786;font-size:12px}");
            builder.AppendLine("@media(max-width:760px){body{padding:12px}.cards{grid-template-columns:repeat(2,1fr);padding:16px}.status,section,.notice{margin-left:16px;margin-right:16px}section{padding-left:0;padding-right:0}header{padding:22px 18px}}");
            builder.AppendLine("</style></head><body><div class='page'>");
            builder.AppendLine("<header><h1>" + Html(title) + "</h1><div class='period'>结算月份：" + report.Year + "年" + report.Month + "月</div></header>");
            builder.AppendLine("<div class='status " + statusClass + "'>" + Html(statusText) + "</div>");
            builder.AppendLine("<div class='cards'>");
            AppendCard(builder, "扫描输入", report.InputCount);
            AppendCard(builder, "正常输出", report.SuccessfulCount);
            AppendCard(builder, "原文件已保留", report.PreservedReviewCopyCount);
            AppendCard(builder, "失败", report.FailedCount);
            builder.AppendLine("</div>");

            var reviewItems = report.Workbooks
                .Where(item => item.Action == GuangdongStage2PreparationActions.Skipped
                    || item.Action == GuangdongStage2PreparationActions.Failed)
                .ToList();
            if (reviewItems.Count > 0)
            {
                builder.AppendLine("<div class='notice'><strong>警告：</strong>在完成以下文件的人工处理前，本批次结果不能视为完整。</div>");
                builder.AppendLine("<section><h2>需要人工复核</h2><table><thead><tr><th>类型</th><th>文件</th><th>原因</th><th>原文件保留位置</th></tr></thead><tbody>");
                foreach (var item in reviewItems)
                {
                    builder.AppendLine("<tr><td>" + Html(item.SettlementKind) + "</td><td>" + Html(item.RelativePath)
                        + "</td><td>" + Html(item.Message) + "</td><td class='path'>"
                        + Html(string.IsNullOrWhiteSpace(item.ReviewCopyPath) ? "原文件未能保留" : item.ReviewCopyPath)
                        + "</td></tr>");
                }
                builder.AppendLine("</tbody></table></section>");
            }
            else
            {
                builder.AppendLine("<section><div class='empty'>没有需要人工复核的文件。</div></section>");
            }

            builder.AppendLine("<section><h2>处理明细</h2><table><tbody>");
            builder.AppendLine("<tr><th>从上月创建</th><td>" + report.CreatedCount + "</td></tr>");
            builder.AppendLine("<tr><th>整理现有目标月</th><td>" + report.NormalizedCount + "</td></tr>");
            builder.AppendLine("<tr><th>原本已经准备完成</th><td>" + report.AlreadyPreparedCount + "</td></tr>");
            builder.AppendLine("<tr><th>跳过</th><td>" + report.SkippedCount + "</td></tr>");
            builder.AppendLine("<tr><th>输入分类守恒</th><td>" + (report.IsClassificationComplete ? "是" : "否") + "</td></tr>");
            builder.AppendLine("<tr><th>输入文件保留完整</th><td>" + (report.HasCompleteOutputSet ? "是" : "否") + "</td></tr>");
            builder.AppendLine("</tbody></table></section>");
            builder.AppendLine("<div class='footer'>本报告由结算自动化工具生成。JSON 与 TXT 报告保留用于追溯。</div>");
            builder.AppendLine("</div></body></html>");
            return builder.ToString();
        }

        private static void AppendCard(StringBuilder builder, string label, int value)
        {
            builder.AppendLine("<div class='card'><div class='label'>" + Html(label)
                + "</div><div class='value'>" + value + "</div></div>");
        }

        private static string Html(string value)
        {
            return WebUtility.HtmlEncode(value ?? string.Empty);
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
