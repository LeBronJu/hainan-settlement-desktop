using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using HainanSettlementTool.Core.Models;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal static class HainanStage2ReportWriter
    {
        internal static HainanStage2Report CreateReport(
            HainanStage2Options options,
            IList<HainanStage2DetailSettlementRow> proxyRows,
            IList<HainanStage2DetailSettlementRow> interRows,
            IList<GroupSettlementTotal> totals,
            string summaryPath,
            IList<string> warnings,
            IList<string> missingOwners,
            IList<HainanStage2CheckIssue> auditIssues)
        {
            var reportPath = Path.Combine(options.OutputDirectory, options.Month + "月结算生成总报告.json");
            var report = new HainanStage2Report
            {
                Month = options.Month,
                Ledger = options.LedgerPath,
                ProxyTemplateDirectory = options.ProxyTemplateDirectory,
                IntermediaryTemplateDirectory = options.IntermediaryTemplateDirectory,
                SummaryTemplate = options.SummaryTemplatePath,
                OutputDirectory = options.OutputDirectory,
                Summary = summaryPath,
                ReportPath = reportPath,
                ValidationReportPath = Path.Combine(options.OutputDirectory, "阶段二校验报告.txt"),
                HtmlReportPath = Path.Combine(
                    options.OutputDirectory,
                    "海南" + options.Month + "月阶段二结算报告.html"),
                GeneratedSummaryReviewPath = Path.Combine(options.OutputDirectory, "自动生成汇总提示.txt"),
                PreflightSignature = options.ExpectedPreflightSignature,
                InputFingerprint = options.ExpectedInputFingerprint,
                ProxyRows = proxyRows.Count,
                IntermediaryRows = interRows.Count,
                ProxyGroups = totals.Count(total => total.Kind == "代理费"),
                IntermediaryGroups = totals.Count(total => total.Kind == "居间费"),
                ProxyTotal = Math.Round(totals.Where(total => total.Kind == "代理费").Sum(total => total.ExpectedNet), 4),
                IntermediaryTotal = Math.Round(totals.Where(total => total.Kind == "居间费").Sum(total => total.ExpectedNet), 4)
            };
            report.Groups.AddRange(totals);
            report.Warnings.AddRange(warnings);
            report.MissingOwners.AddRange(missingOwners);
            report.AuditIssues.AddRange(auditIssues);
            return report;
        }

        internal static void WriteReport(HainanStage2Options options, HainanStage2Report report)
        {
            WriteReport(report, report.ReportPath);
        }

        internal static void WriteReport(HainanStage2Report report, string physicalPath)
        {
            File.WriteAllText(physicalPath, JsonConvert.SerializeObject(report, Formatting.Indented), Encoding.UTF8);
        }

        internal static void WriteHtmlReport(HainanStage2Report report, string physicalPath)
        {
            File.WriteAllText(
                physicalPath,
                ReadableHtmlReportRenderer.Render(BuildReadableReport(report)),
                Encoding.UTF8);
        }

        internal static void WriteWarnings(HainanStage2Options options, IList<string> warnings)
        {
            var path = Path.Combine(options.OutputDirectory, "自动生成汇总提示.txt");
            WriteWarnings(warnings, path);
        }

        internal static void WriteWarnings(IList<string> warnings, string physicalPath)
        {
            if (warnings.Count > 0)
            {
                File.WriteAllLines(physicalPath, warnings, Encoding.UTF8);
            }
            else
            {
                File.WriteAllText(physicalPath, "本批无自动生成汇总提示。", Encoding.UTF8);
            }
        }

        internal static void WriteAuditReport(HainanStage2Options options, HainanStage2Report report)
        {
            var path = string.IsNullOrWhiteSpace(report.ValidationReportPath)
                ? Path.Combine(options.OutputDirectory, "阶段二校验报告.txt")
                : report.ValidationReportPath;
            WriteAuditReport(options, report, path);
        }

        internal static void WriteAuditReport(
            HainanStage2Options options,
            HainanStage2Report report,
            string physicalPath)
        {
            if (report.AuditIssues.Count == 0 && report.Warnings.Count == 0 && report.MissingOwners.Count == 0)
            {
                File.WriteAllLines(
                    physicalPath,
                    new[]
                    {
                        "阶段二校验报告",
                        "结算月份：2026年" + options.Month + "月",
                        "本批无需复核的校验项。"
                    },
                    Encoding.UTF8);
                return;
            }

            var lines = new List<string>
            {
                "阶段二校验报告",
                "结算月份：2026年" + options.Month + "月",
                "说明：文件已照常生成；当前分表和汇总表金额采用分表自算结果。",
                "提示：如果确认台账金额才是正确结果，请同步检查/修改对应分表和汇总表。",
                string.Empty
            };

            if (report.AuditIssues.Count > 0)
            {
                lines.Add("一、校验问题");
                for (var index = 0; index < report.AuditIssues.Count; index++)
                {
                    var issue = report.AuditIssues[index];
                    lines.Add((index + 1) + ". [" + issue.Severity + "] " + issue.Category);
                    if (!string.IsNullOrWhiteSpace(issue.Kind))
                    {
                        lines.Add("   类型：" + issue.Kind);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Owner) || !string.IsNullOrWhiteSpace(issue.Entity))
                    {
                        lines.Add("   负责人/主体：" + issue.Owner + " / " + issue.Entity);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Customer))
                    {
                        lines.Add("   客户：" + issue.Customer);
                    }
                    if (issue.LedgerRow > 0)
                    {
                        lines.Add("   台账行：" + issue.LedgerRow);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.PreviousValue) || !string.IsNullOrWhiteSpace(issue.CurrentValue))
                    {
                        lines.Add("   对比：" + issue.PreviousValue + "；" + issue.CurrentValue);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Message))
                    {
                        lines.Add("   问题：" + issue.Message);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Suggestion))
                    {
                        lines.Add("   建议：" + issue.Suggestion);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.TemplateFile))
                    {
                        lines.Add("   文件：" + issue.TemplateFile);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.SheetName))
                    {
                        lines.Add("   工作表：" + issue.SheetName);
                    }
                }
                lines.Add(string.Empty);
            }

            if (report.Warnings.Count > 0)
            {
                lines.Add("二、自动生成汇总提示");
                foreach (var warning in report.Warnings)
                {
                    lines.Add("- " + warning);
                }
                lines.Add(string.Empty);
            }

            if (report.MissingOwners.Count > 0)
            {
                lines.Add("三、负责人缺失");
                foreach (var missingOwner in report.MissingOwners)
                {
                    lines.Add("- " + missingOwner);
                }
            }

            File.WriteAllLines(physicalPath, lines, Encoding.UTF8);
        }

        private static ReadableReportDocument BuildReadableReport(HainanStage2Report report)
        {
            var reviewIssues = report.AuditIssues
                .Where(issue => issue != null && issue.Disposition == Stage2PreflightDisposition.Review)
                .ToList();
            var informationIssues = report.AuditIssues
                .Where(issue => issue != null && issue.Disposition == Stage2PreflightDisposition.Information)
                .ToList();
            var warnings = report.Warnings.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            var missingOwners = report.MissingOwners.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            var needsReview = reviewIssues.Count > 0 || warnings.Count > 0 || missingOwners.Count > 0;
            var document = new ReadableReportDocument
            {
                Title = "海南" + report.Month + "月阶段二结算报告",
                PeriodLabel = "结算月份：2026年" + report.Month + "月",
                Status = needsReview ? ReadableReportStatus.Review : ReadableReportStatus.Success,
                StatusText = needsReview ? "生成完成，但需要人工复核" : "全部生成完成",
                StatusDetail = needsReview
                    ? "正式分表和汇总表已完整生成；付款前请处理下方复核项目。"
                    : "未发现需要人工复核的自动化校验项目。"
            };
            document.Metrics.Add(new ReadableReportMetric("代理费主体", report.ProxyGroups.ToString(CultureInfo.InvariantCulture)));
            document.Metrics.Add(new ReadableReportMetric("居间费主体", report.IntermediaryGroups.ToString(CultureInfo.InvariantCulture)));
            document.Metrics.Add(new ReadableReportMetric("代理费合计（万元）", Amount(report.ProxyTotal)));
            document.Metrics.Add(new ReadableReportMetric("居间费合计（万元）", Amount(report.IntermediaryTotal)));

            if (needsReview)
            {
                document.Notices.Add(new ReadableReportNotice(
                    "付款前请检查",
                    "本报告只帮助定位自动化复核项目；如需人工修改金额，请同步检查对应分表和汇总表。",
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
            foreach (var missingOwner in missingOwners)
            {
                reviewSection.Rows.Add(new ReadableReportRow("负责人缺失", string.Empty, string.Empty, missingOwner, "请检查台账负责人。"));
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

            AddGroupSection(document, report.Groups);
            AddTechnicalSection(
                document,
                report.Ledger,
                report.Summary,
                report.ReportPath,
                report.ValidationReportPath,
                report.GeneratedSummaryReviewPath,
                report.PreflightSignature,
                report.InputFingerprint,
                report.AuditIssues);
            return document;
        }

        private static ReadableReportRow IssueRow(Stage2PreflightIssue issue)
        {
            var comparison = JoinNonEmpty(
                issue.Message,
                JoinNonEmpty(issue.PreviousValue, issue.CurrentValue));
            return new ReadableReportRow(
                issue.Category,
                string.IsNullOrWhiteSpace(issue.SettlementKind) ? issue.Kind : issue.SettlementKind,
                JoinNonEmpty(JoinNonEmpty(issue.Owner, issue.Entity), issue.Customer),
                comparison,
                issue.Suggestion);
        }

        private static void AddGroupSection(
            ReadableReportDocument document,
            IEnumerable<GroupSettlementTotal> groups)
        {
            var section = new ReadableReportSection(
                "结算主体明细",
                "费用类型",
                "负责人",
                "主体",
                "客户行数",
                "金额（万元）",
                "分表文件");
            foreach (var group in groups.OrderBy(item => item.Kind).ThenBy(item => item.Owner).ThenBy(item => item.Entity))
            {
                section.Rows.Add(new ReadableReportRow(
                    group.Kind,
                    group.Owner,
                    group.Entity,
                    group.Rows.ToString(CultureInfo.InvariantCulture),
                    Amount(group.ExpectedNet),
                    group.OutputFile));
            }
            document.Sections.Add(section);
        }

        private static void AddTechnicalSection(
            ReadableReportDocument document,
            string ledger,
            string summary,
            string jsonReport,
            string validationReport,
            string generatedSummaryReview,
            string preflightSignature,
            string inputFingerprint,
            IEnumerable<HainanStage2CheckIssue> issues)
        {
            var section = new ReadableReportSection("文件与技术详情", "项目", "内容")
            {
                IsCollapsed = true
            };
            section.Rows.Add(new ReadableReportRow("台账", ledger));
            section.Rows.Add(new ReadableReportRow("汇总表", summary));
            section.Rows.Add(new ReadableReportRow("JSON 报告", jsonReport));
            section.Rows.Add(new ReadableReportRow("文本校验报告", validationReport));
            section.Rows.Add(new ReadableReportRow("自动生成提示", generatedSummaryReview));
            section.Rows.Add(new ReadableReportRow("预检签名", preflightSignature));
            section.Rows.Add(new ReadableReportRow("输入指纹", inputFingerprint));
            foreach (var issue in issues.Where(item => item != null))
            {
                section.Rows.Add(new ReadableReportRow(
                    string.IsNullOrWhiteSpace(issue.Category) ? "校验项目" : issue.Category,
                    JoinNonEmpty(
                        issue.LedgerRow > 0 ? "台账行：" + issue.LedgerRow.ToString(CultureInfo.InvariantCulture) : string.Empty,
                        JoinNonEmpty(issue.TemplateFile, issue.SheetName))));
            }
            document.Sections.Add(section);
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
