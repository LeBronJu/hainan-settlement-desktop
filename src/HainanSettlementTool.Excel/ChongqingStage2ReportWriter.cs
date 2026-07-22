using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal static class ChongqingStage2ReportWriter
    {
        public static ChongqingStage2Report CreateReport(
            ChongqingStage2Options options,
            IList<ChongqingSettlementDetail> details,
            IList<GroupSettlementTotal> groups,
            string summaryPath,
            IList<string> warnings,
            IList<ChongqingStage2CheckIssue> auditIssues)
        {
            var report = new ChongqingStage2Report
            {
                Month = options.Month,
                Ledger = options.LedgerPath,
                ProxyTemplateDirectory = options.ProxyTemplateDirectory,
                IntermediaryTemplateDirectory = options.IntermediaryTemplateDirectory,
                RefundTemplateDirectory = options.RefundTemplateDirectory,
                SummaryTemplate = options.SummaryTemplatePath,
                OutputDirectory = options.OutputDirectory,
                ProxyOutputDirectory = ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Proxy),
                IntermediaryOutputDirectory = ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Intermediary),
                RefundOutputDirectory = ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Refund),
                Summary = summaryPath,
                ReportPath = Path.Combine(options.OutputDirectory, options.Month + "月重庆阶段二生成报告.json"),
                ValidationReportPath = Path.Combine(options.OutputDirectory, "重庆阶段二校验报告.txt"),
                HtmlReportPath = Path.Combine(
                    options.OutputDirectory,
                    "重庆" + options.Month + "月阶段二结算报告.html"),
                PreflightSignature = options.ExpectedPreflightSignature,
                InputFingerprint = options.ExpectedInputFingerprint,
                ProxyRows = details.Count(detail => detail.Kind == ChongqingStage2SettlementKinds.Proxy),
                IntermediaryRows = details.Count(detail => detail.Kind == ChongqingStage2SettlementKinds.Intermediary),
                RefundRows = details.Count(detail => detail.Kind == ChongqingStage2SettlementKinds.Refund),
                ProxyGroups = groups.Count(group => group.Kind == ChongqingStage2SettlementKinds.Proxy),
                IntermediaryGroups = groups.Count(group => group.Kind == ChongqingStage2SettlementKinds.Intermediary),
                RefundGroups = groups.Count(group => group.Kind == ChongqingStage2SettlementKinds.Refund),
                ProxyTotal = Math.Round(groups.Where(group => group.Kind == ChongqingStage2SettlementKinds.Proxy).Sum(group => group.ExpectedNet), 4),
                IntermediaryTotal = Math.Round(groups.Where(group => group.Kind == ChongqingStage2SettlementKinds.Intermediary).Sum(group => group.ExpectedNet), 4),
                RefundTotal = Math.Round(groups.Where(group => group.Kind == ChongqingStage2SettlementKinds.Refund).Sum(group => group.ExpectedNet), 4)
            };
            report.Groups.AddRange(groups);
            report.Warnings.AddRange(warnings);
            report.AuditIssues.AddRange(auditIssues);
            return report;
        }

        public static void Write(ChongqingStage2Report report)
        {
            Write(report, report.ReportPath, report.ValidationReportPath, report.HtmlReportPath);
        }

        public static void Write(
            ChongqingStage2Report report,
            string physicalReportPath,
            string physicalValidationReportPath,
            string physicalHtmlReportPath = null)
        {
            File.WriteAllText(physicalReportPath, JsonConvert.SerializeObject(report, Formatting.Indented), Encoding.UTF8);
            WriteValidationReport(report, physicalValidationReportPath);
            if (!string.IsNullOrWhiteSpace(physicalHtmlReportPath))
            {
                File.WriteAllText(
                    physicalHtmlReportPath,
                    ReadableHtmlReportRenderer.Render(BuildReadableReport(report)),
                    Encoding.UTF8);
            }
        }

        private static void WriteValidationReport(ChongqingStage2Report report, string physicalPath)
        {
            var lines = new List<string>
            {
                "重庆阶段二校验报告",
                "结算月份：" + ChongqingStage2Layout.Year + "年" + report.Month + "月",
                "说明：文件已生成；分表和汇总表金额按分表公式结果写入。",
                string.Empty
            };

            if (report.AuditIssues.Count == 0 && report.Warnings.Count == 0)
            {
                lines.Add("未发现需要提示的自动化校验项。");
            }

            if (report.AuditIssues.Count > 0)
            {
                lines.Add("一、校验问题");
                for (var index = 0; index < report.AuditIssues.Count; index++)
                {
                    var issue = report.AuditIssues[index];
                    lines.Add((index + 1) + ". [" + issue.Severity + "] " + issue.Category);
                    lines.Add("   类型：" + issue.SettlementKind);
                    lines.Add("   主体：" + issue.Entity);
                    lines.Add("   客户：" + issue.Customer);
                    lines.Add("   台账行：" + issue.LedgerRow);
                    lines.Add("   对比：" + issue.PreviousValue + "；" + issue.CurrentValue);
                    lines.Add("   建议：" + issue.Suggestion);
                    lines.Add("   文件：" + issue.TemplateFile);
                    lines.Add("   工作表：" + issue.SheetName);
                }

                lines.Add(string.Empty);
            }

            if (report.Warnings.Count > 0)
            {
                lines.Add("二、自动生成提示");
                foreach (var warning in report.Warnings)
                {
                    lines.Add("- " + warning);
                }
            }

            File.WriteAllLines(physicalPath, lines, Encoding.UTF8);
        }

        private static ReadableReportDocument BuildReadableReport(ChongqingStage2Report report)
        {
            var reviewIssues = report.AuditIssues
                .Where(issue => issue != null && issue.Disposition == Stage2PreflightDisposition.Review)
                .ToList();
            var informationIssues = report.AuditIssues
                .Where(issue => issue != null && issue.Disposition == Stage2PreflightDisposition.Information)
                .ToList();
            var warnings = report.Warnings.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            var needsReview = reviewIssues.Count > 0 || warnings.Count > 0;
            var document = new ReadableReportDocument
            {
                Title = "重庆" + report.Month + "月阶段二结算报告",
                PeriodLabel = "结算月份：" + ChongqingStage2Layout.Year + "年" + report.Month + "月",
                Status = needsReview ? ReadableReportStatus.Review : ReadableReportStatus.Success,
                StatusText = needsReview ? "生成完成，但需要人工复核" : "全部生成完成",
                StatusDetail = needsReview
                    ? "正式分表、退补表和汇总表已完整生成；付款前请处理下方复核项目。"
                    : "未发现需要人工复核的自动化校验项目。"
            };
            document.Metrics.Add(new ReadableReportMetric("代理主体", report.ProxyGroups.ToString(CultureInfo.InvariantCulture)));
            document.Metrics.Add(new ReadableReportMetric("代理合计（万元）", Amount(report.ProxyTotal)));
            document.Metrics.Add(new ReadableReportMetric("居间主体", report.IntermediaryGroups.ToString(CultureInfo.InvariantCulture)));
            document.Metrics.Add(new ReadableReportMetric("居间合计（万元）", Amount(report.IntermediaryTotal)));
            document.Metrics.Add(new ReadableReportMetric("退补主体", report.RefundGroups.ToString(CultureInfo.InvariantCulture)));
            document.Metrics.Add(new ReadableReportMetric("退补合计（万元）", Amount(report.RefundTotal)));

            if (needsReview)
            {
                document.Notices.Add(new ReadableReportNotice(
                    "付款前请检查",
                    "重庆退补额外扣减块及其它人工边界仍按校验提示处理；HTML 报告不会替代对分表和汇总表的复核。",
                    ReadableReportNoticeTone.Review));
            }

            var reviewSection = new ReadableReportSection(
                "付款前需要检查",
                "类别",
                "类型",
                "主体/客户",
                "问题",
                "处理建议")
            {
                EmptyMessage = "没有需要人工复核的项目。"
            };
            foreach (var issue in reviewIssues)
            {
                reviewSection.Rows.Add(IssueRow(issue));
            }
            foreach (var warning in warnings)
            {
                reviewSection.Rows.Add(new ReadableReportRow("自动生成提示", string.Empty, string.Empty, warning, "请核对生成文件。"));
            }
            document.Sections.Add(reviewSection);

            if (informationIssues.Count > 0)
            {
                var informationSection = new ReadableReportSection(
                    "本次自动处理记录",
                    "类别",
                    "类型",
                    "主体/客户",
                    "说明",
                    "建议");
                foreach (var issue in informationIssues)
                {
                    informationSection.Rows.Add(IssueRow(issue));
                }
                document.Sections.Add(informationSection);
            }

            var groupSection = new ReadableReportSection(
                "结算主体明细",
                "费用类型",
                "负责人",
                "主体",
                "客户行数",
                "金额（万元）",
                "分表文件");
            foreach (var group in report.Groups.OrderBy(item => item.Kind).ThenBy(item => item.Owner).ThenBy(item => item.Entity))
            {
                groupSection.Rows.Add(new ReadableReportRow(
                    group.Kind,
                    group.Owner,
                    group.Entity,
                    group.Rows.ToString(CultureInfo.InvariantCulture),
                    Amount(group.ExpectedNet),
                    group.OutputFile));
            }
            document.Sections.Add(groupSection);

            var technical = new ReadableReportSection("文件与技术详情", "项目", "内容")
            {
                IsCollapsed = true
            };
            technical.Rows.Add(new ReadableReportRow("台账", report.Ledger));
            technical.Rows.Add(new ReadableReportRow("汇总表", report.Summary));
            technical.Rows.Add(new ReadableReportRow("JSON 报告", report.ReportPath));
            technical.Rows.Add(new ReadableReportRow("文本校验报告", report.ValidationReportPath));
            technical.Rows.Add(new ReadableReportRow("预检签名", report.PreflightSignature));
            technical.Rows.Add(new ReadableReportRow("输入指纹", report.InputFingerprint));
            foreach (var issue in report.AuditIssues.Where(item => item != null))
            {
                technical.Rows.Add(new ReadableReportRow(
                    string.IsNullOrWhiteSpace(issue.Category) ? "校验项目" : issue.Category,
                    JoinNonEmpty(
                        issue.LedgerRow > 0 ? "台账行：" + issue.LedgerRow.ToString(CultureInfo.InvariantCulture) : string.Empty,
                        JoinNonEmpty(issue.TemplateFile, issue.SheetName))));
            }
            document.Sections.Add(technical);
            return document;
        }

        private static ReadableReportRow IssueRow(Stage2PreflightIssue issue)
        {
            return new ReadableReportRow(
                issue.Category,
                string.IsNullOrWhiteSpace(issue.SettlementKind) ? issue.Kind : issue.SettlementKind,
                JoinNonEmpty(JoinNonEmpty(issue.Owner, issue.Entity), issue.Customer),
                JoinNonEmpty(issue.Message, JoinNonEmpty(issue.PreviousValue, issue.CurrentValue)),
                issue.Suggestion);
        }

        private static string JoinNonEmpty(string first, string second)
        {
            if (string.IsNullOrWhiteSpace(first))
            {
                return second ?? string.Empty;
            }
            if (string.IsNullOrWhiteSpace(second))
            {
                return first;
            }
            return first + "\n" + second;
        }

        private static string Amount(double value)
        {
            return value.ToString("0.####", CultureInfo.InvariantCulture);
        }
    }
}
