using System;
using System.Net;
using System.Text;

namespace HainanSettlementTool.Excel
{
    internal static class ReadableHtmlReportRenderer
    {
        internal static string Render(ReadableReportDocument document)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            Validate(document);
            var builder = new StringBuilder();
            builder.AppendLine("<!DOCTYPE html>");
            builder.AppendLine("<html lang='zh-CN'><head><meta charset='utf-8'>");
            builder.AppendLine("<meta name='viewport' content='width=device-width,initial-scale=1'>");
            builder.AppendLine("<meta http-equiv='Content-Security-Policy' content=\"default-src 'none'; style-src 'unsafe-inline'\">");
            builder.AppendLine("<title>" + Html(document.Title) + "</title>");
            AppendStyle(builder);
            builder.AppendLine("</head><body><main class='page'>");
            builder.AppendLine("<header><h1>" + Html(document.Title) + "</h1><div class='period'>"
                + Html(document.PeriodLabel) + "</div></header>");

            var statusClass = StatusClass(document.Status);
            builder.AppendLine("<div class='status " + statusClass + "'><div class='status-title'>"
                + Html(document.StatusText) + "</div>");
            if (!string.IsNullOrWhiteSpace(document.StatusDetail))
            {
                builder.AppendLine("<div class='status-detail'>" + Html(document.StatusDetail) + "</div>");
            }
            builder.AppendLine("</div>");

            if (document.Metrics.Count > 0)
            {
                builder.AppendLine("<div class='metrics'>");
                foreach (var metric in document.Metrics)
                {
                    builder.AppendLine("<div class='metric'><div class='metric-label'>"
                        + Html(metric.Label) + "</div><div class='metric-value'>"
                        + Html(metric.Value) + "</div></div>");
                }
                builder.AppendLine("</div>");
            }

            foreach (var notice in document.Notices)
            {
                builder.AppendLine("<div class='notice " + NoticeClass(notice.Tone) + "'><strong>"
                    + Html(notice.Title) + "</strong><div>" + Html(notice.Body) + "</div></div>");
            }

            foreach (var section in document.Sections)
            {
                AppendSection(builder, section);
            }

            builder.AppendLine("<footer>" + Html(string.IsNullOrWhiteSpace(document.Footer)
                ? "本报告由结算自动化工具生成。JSON 与 TXT 报告继续保留用于追溯。"
                : document.Footer) + "</footer>");
            builder.AppendLine("</main></body></html>");
            return builder.ToString();
        }

        private static void AppendSection(StringBuilder builder, ReadableReportSection section)
        {
            if (section.IsCollapsed)
            {
                builder.AppendLine("<details class='section'><summary>" + Html(section.Title) + "</summary>");
            }
            else
            {
                builder.AppendLine("<section class='section'><h2>" + Html(section.Title) + "</h2>");
            }

            if (section.Rows.Count == 0)
            {
                builder.AppendLine("<div class='empty'>" + Html(string.IsNullOrWhiteSpace(section.EmptyMessage)
                    ? "没有相关项目。"
                    : section.EmptyMessage) + "</div>");
            }
            else
            {
                builder.AppendLine("<div class='table-wrap'><table>");
                if (section.Headers.Count > 0)
                {
                    builder.AppendLine("<thead><tr>");
                    foreach (var header in section.Headers)
                    {
                        builder.AppendLine("<th>" + Html(header) + "</th>");
                    }
                    builder.AppendLine("</tr></thead>");
                }

                builder.AppendLine("<tbody>");
                foreach (var row in section.Rows)
                {
                    builder.AppendLine("<tr>");
                    foreach (var cell in row.Cells)
                    {
                        builder.AppendLine("<td>" + Html(cell) + "</td>");
                    }
                    builder.AppendLine("</tr>");
                }
                builder.AppendLine("</tbody></table></div>");
            }

            builder.AppendLine(section.IsCollapsed ? "</details>" : "</section>");
        }

        private static void Validate(ReadableReportDocument document)
        {
            if (string.IsNullOrWhiteSpace(document.Title))
            {
                throw new InvalidOperationException("可读报告标题不能为空。");
            }

            if (string.IsNullOrWhiteSpace(document.StatusText))
            {
                throw new InvalidOperationException("可读报告状态不能为空。");
            }

            foreach (var section in document.Sections)
            {
                if (section == null || string.IsNullOrWhiteSpace(section.Title))
                {
                    throw new InvalidOperationException("可读报告区块标题不能为空。");
                }

                if (section.Headers.Count == 0)
                {
                    continue;
                }

                foreach (var row in section.Rows)
                {
                    if (row.Cells.Count != section.Headers.Count)
                    {
                        throw new InvalidOperationException(
                            "可读报告区块“" + section.Title + "”的表头和数据列数不一致。");
                    }
                }
            }
        }

        private static string Html(string value)
        {
            return WebUtility.HtmlEncode(value ?? string.Empty);
        }

        private static string StatusClass(ReadableReportStatus status)
        {
            switch (status)
            {
                case ReadableReportStatus.Critical:
                    return "critical";
                case ReadableReportStatus.Review:
                    return "review";
                default:
                    return "success";
            }
        }

        private static string NoticeClass(ReadableReportNoticeTone tone)
        {
            switch (tone)
            {
                case ReadableReportNoticeTone.Critical:
                    return "notice-critical";
                case ReadableReportNoticeTone.Review:
                    return "notice-review";
                default:
                    return "notice-information";
            }
        }

        private static void AppendStyle(StringBuilder builder)
        {
            builder.AppendLine("<style>");
            builder.AppendLine("*{box-sizing:border-box}body{margin:0;padding:32px;background:#f3f6f8;color:#17212b;font-family:'Microsoft YaHei UI','Microsoft YaHei',sans-serif;line-height:1.55}");
            builder.AppendLine(".page{max-width:1180px;margin:auto;background:#fff;border:1px solid #dfe6ec;border-radius:14px;box-shadow:0 8px 28px rgba(20,35,55,.10);overflow:hidden}");
            builder.AppendLine("header{padding:28px 32px;background:#153a5b;color:#fff}h1{margin:0 0 6px;font-size:26px}.period{opacity:.86}");
            builder.AppendLine(".status{margin:24px 32px 0;padding:16px 18px;border-radius:10px;border:1px solid}.status-title{font-size:20px;font-weight:700}.status-detail{margin-top:5px;white-space:pre-wrap}");
            builder.AppendLine(".status.success{background:#e8f5e9;color:#236b2b;border-color:#9bd0a0}.status.review{background:#fff3cd;color:#7a4b00;border-color:#f0cf68}.status.critical{background:#fdecea;color:#9a2018;border-color:#e9a29d}");
            builder.AppendLine(".metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;padding:22px 32px}.metric{padding:16px;background:#f7f9fb;border:1px solid #e1e7ed;border-radius:10px}.metric-label{font-size:13px;color:#607080}.metric-value{margin-top:4px;font-size:26px;font-weight:700;word-break:break-word}");
            builder.AppendLine(".notice{margin:0 32px 16px;padding:14px 16px;border-left:5px solid;border-radius:6px;white-space:pre-wrap}.notice strong{display:block;margin-bottom:3px}.notice-review{background:#fff7e6;border-color:#e89a18}.notice-critical{background:#fdecea;border-color:#c0392b}.notice-information{background:#eef6fc;border-color:#4c88b5}");
            builder.AppendLine(".section{display:block;margin:0;padding:8px 32px 28px}.section h2,.section summary{font-size:18px;font-weight:700;margin:0 0 14px}.section summary{cursor:pointer;padding:10px 0}.table-wrap{overflow-x:auto;border:1px solid #e1e7ed;border-radius:8px}table{width:100%;border-collapse:collapse;font-size:14px}th,td{text-align:left;vertical-align:top;padding:10px;border-bottom:1px solid #e6ebef;white-space:pre-wrap;overflow-wrap:anywhere}th{background:#eef3f7;color:#33495c}tbody tr:last-child td{border-bottom:0}.empty{padding:16px;background:#eef8ef;border-radius:8px;color:#236b2b}");
            builder.AppendLine("footer{padding:18px 32px;background:#f7f9fb;color:#677786;font-size:12px}");
            builder.AppendLine("@media(max-width:760px){body{padding:12px}header{padding:22px 18px}.status,.notice{margin-left:16px;margin-right:16px}.metrics{padding:16px}.section{padding-left:16px;padding-right:16px}th,td{min-width:110px}}");
            builder.AppendLine("@media print{body{padding:0;background:#fff}.page{border:0;box-shadow:none}.section{break-inside:avoid}details{display:block}details>summary{list-style:none}}");
            builder.AppendLine("</style>");
        }
    }
}
