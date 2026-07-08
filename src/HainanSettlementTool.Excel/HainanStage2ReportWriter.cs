using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Models;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal static class HainanStage2ReportWriter
    {
        internal static HainanStage2Report CreateReport(
            HainanStage2Options options,
            IList<DetailSettlementRow> proxyRows,
            IList<DetailSettlementRow> interRows,
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
            File.WriteAllText(report.ReportPath, JsonConvert.SerializeObject(report, Formatting.Indented), System.Text.Encoding.UTF8);
        }

        internal static void WriteWarnings(HainanStage2Options options, IList<string> warnings)
        {
            var path = Path.Combine(options.OutputDirectory, "自动生成汇总提示.txt");
            if (warnings.Count > 0)
            {
                File.WriteAllLines(path, warnings, System.Text.Encoding.UTF8);
            }
            else if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        internal static void WriteAuditReport(HainanStage2Options options, HainanStage2Report report)
        {
            var path = Path.Combine(options.OutputDirectory, "阶段二校验报告.txt");
            if (report.AuditIssues.Count == 0 && report.Warnings.Count == 0 && report.MissingOwners.Count == 0)
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
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

            File.WriteAllLines(path, lines, System.Text.Encoding.UTF8);
        }
    }
}
