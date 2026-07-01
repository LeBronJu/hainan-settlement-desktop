using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Wpf
{
    public partial class Stage2PreflightWindow : Window
    {
        public Stage2PreflightWindow(Stage2PreflightReport report)
        {
            InitializeComponent();
            SummaryText.Text = "结算月份：2026年" + report.Month + "月；共发现 " + report.Issues.Count + " 条需要确认的变化。";
            IssueGroupsList.ItemsSource = BuildIssueGroups(report);
        }

        private static List<PreflightIssueGroupViewModel> BuildIssueGroups(Stage2PreflightReport report)
        {
            return report.Issues
                .GroupBy(issue => issue.Category)
                .OrderBy(group => group.Key)
                .Select(group =>
                {
                    var issues = group.Select(BuildIssueItem).ToList();
                    return new PreflightIssueGroupViewModel
                    {
                        Category = group.Key,
                        CountText = issues.Count + " 条",
                        Issues = issues
                    };
                })
                .ToList();
        }

        private static PreflightIssueItemViewModel BuildIssueItem(Stage2CheckIssue issue)
        {
            return new PreflightIssueItemViewModel
            {
                PrimaryText = "[" + issue.Severity + "] " + issue.Message,
                FileText = BuildFileText(issue),
                HandlingText = BuildHandlingText(issue),
                ContextText = BuildContextText(issue)
            };
        }

        private static string BuildFileText(Stage2CheckIssue issue)
        {
            if (string.IsNullOrWhiteSpace(issue.TemplateFile))
            {
                return "分表文件：未匹配到上月分表模板";
            }

            return "分表文件：" + Path.GetFileName(issue.TemplateFile);
        }

        private static string BuildHandlingText(Stage2CheckIssue issue)
        {
            if (string.IsNullOrWhiteSpace(issue.Suggestion))
            {
                return "处理方式：请人工确认后决定是否继续生成。";
            }

            return "处理方式：" + issue.Suggestion;
        }

        private static string BuildContextText(Stage2CheckIssue issue)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(issue.Kind))
            {
                parts.Add("类型：" + issue.Kind);
            }

            if (!string.IsNullOrWhiteSpace(issue.Owner) || !string.IsNullOrWhiteSpace(issue.Entity))
            {
                parts.Add("负责人/主体：" + issue.Owner + " / " + issue.Entity);
            }

            if (!string.IsNullOrWhiteSpace(issue.Customer))
            {
                parts.Add("客户/明细：" + issue.Customer);
            }

            if (issue.LedgerRow > 0)
            {
                parts.Add("台账行：" + issue.LedgerRow);
            }

            if (!string.IsNullOrWhiteSpace(issue.PreviousValue) || !string.IsNullOrWhiteSpace(issue.CurrentValue))
            {
                parts.Add("对比：" + issue.PreviousValue + "；" + issue.CurrentValue);
            }

            if (!string.IsNullOrWhiteSpace(issue.SheetName))
            {
                parts.Add("工作表：" + issue.SheetName);
            }

            if (!string.IsNullOrWhiteSpace(issue.TemplateFile))
            {
                parts.Add("完整路径：" + issue.TemplateFile);
            }

            return string.Join("；", parts);
        }

        public sealed class PreflightIssueGroupViewModel
        {
            public string Category { get; set; }
            public string CountText { get; set; }
            public List<PreflightIssueItemViewModel> Issues { get; set; }
        }

        public sealed class PreflightIssueItemViewModel
        {
            public string PrimaryText { get; set; }
            public string FileText { get; set; }
            public string HandlingText { get; set; }
            public string ContextText { get; set; }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
