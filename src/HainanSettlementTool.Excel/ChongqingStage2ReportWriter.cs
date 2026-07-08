using System;
using System.Collections.Generic;
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
            File.WriteAllText(report.ReportPath, JsonConvert.SerializeObject(report, Formatting.Indented), Encoding.UTF8);
            WriteValidationReport(report);
        }

        private static void WriteValidationReport(ChongqingStage2Report report)
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

            File.WriteAllLines(report.ValidationReportPath, lines, Encoding.UTF8);
        }
    }
}
