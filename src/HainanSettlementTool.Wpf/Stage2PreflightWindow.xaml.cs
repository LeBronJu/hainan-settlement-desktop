using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Wpf
{
    public partial class Stage2PreflightWindow : Window
    {
        private readonly List<PreflightIssueItemViewModel> _issueRows;

        public Stage2PreflightWindow(Stage2PreflightReport report)
        {
            InitializeComponent();
            SummaryText.Text = "结算月份：2026年" + report.Month + "月；共发现 " + report.Issues.Count + " 条需要确认的变化。";
            var issueGroups = BuildIssueGroups(report);
            _issueRows = issueGroups.SelectMany(group => group.Issues).ToList();
            IssueGroupsList.ItemsSource = issueGroups;
        }

        public List<Stage2SummarySubjectDecision> SummarySubjectDecisions { get; private set; } =
            new List<Stage2SummarySubjectDecision>();

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
            IEnumerable<string> paymentParties = issue.AvailablePaymentParties.Count > 0
                ? issue.AvailablePaymentParties.AsEnumerable()
                : Stage2PaymentParties.Supported.AsEnumerable();
            return new PreflightIssueItemViewModel
            {
                PrimaryText = "[" + issue.Severity + "] " + issue.Message,
                FileText = BuildFileText(issue),
                HandlingText = BuildHandlingText(issue),
                ContextText = BuildContextText(issue),
                Kind = issue.Kind,
                Entity = issue.Entity,
                RequiresPaymentPartySelection = issue.RequiresPaymentPartySelection,
                PaymentSelectionVisibility = issue.RequiresPaymentPartySelection ? Visibility.Visible : Visibility.Collapsed,
                PaymentPartyOptions = paymentParties
                    .Select(party => new PaymentPartyOption { Value = party, DisplayText = party })
                    .ToList()
            };
        }

        private static string BuildFileText(Stage2CheckIssue issue)
        {
            if (issue.RequiresPaymentPartySelection)
            {
                return string.IsNullOrWhiteSpace(issue.TemplateFile)
                    ? "汇总表模板：未找到"
                    : "汇总表模板：" + Path.GetFileName(issue.TemplateFile);
            }

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
            public string Kind { get; set; }
            public string Entity { get; set; }
            public bool RequiresPaymentPartySelection { get; set; }
            public Visibility PaymentSelectionVisibility { get; set; }
            public List<PaymentPartyOption> PaymentPartyOptions { get; set; }
            public string SelectedPaymentParty { get; set; }
        }

        public sealed class PaymentPartyOption
        {
            public string Value { get; set; }
            public string DisplayText { get; set; }

            public override string ToString()
            {
                return DisplayText;
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var missing = _issueRows
                .Where(row => row.RequiresPaymentPartySelection && string.IsNullOrWhiteSpace(row.SelectedPaymentParty))
                .Select(row => row.Kind + " " + row.Entity)
                .ToList();
            if (missing.Count > 0)
            {
                ErrorText.Text = "请先为新增汇总主体选择支付方：" + BuildShortSubjectList(missing);
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            SummarySubjectDecisions = _issueRows
                .Where(row => row.RequiresPaymentPartySelection)
                .Select(row => new Stage2SummarySubjectDecision
                {
                    SettlementKind = row.Kind,
                    Entity = row.Entity,
                    PaymentParty = row.SelectedPaymentParty
                })
                .ToList();
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private static string BuildShortSubjectList(IEnumerable<string> subjects)
        {
            var names = subjects.Take(5).ToList();
            var suffix = subjects.Skip(5).Any() ? " 等" : string.Empty;
            return string.Join("、", names) + suffix;
        }
    }
}
